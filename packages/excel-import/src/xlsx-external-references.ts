import { XMLParser } from 'fast-xml-parser'

import { getZipText, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface ExternalCachedNumber {
  readonly kind: 'number'
  readonly value: number
}

interface ExternalCachedString {
  readonly kind: 'string'
  readonly value: string
}

interface ExternalCachedBoolean {
  readonly kind: 'boolean'
  readonly value: boolean
}

interface ExternalCachedError {
  readonly kind: 'error'
  readonly value: string
}

type ExternalCachedValue = ExternalCachedNumber | ExternalCachedString | ExternalCachedBoolean | ExternalCachedError
type ExternalCachedCells = Map<string, ExternalCachedValue>
type ExternalCachedSheets = Map<string, ExternalCachedCells>
export type ImportedExternalLinkCaches = Map<number, ExternalCachedSheets>

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
}

export interface ImportedFormulaExternalReferenceTranslation {
  readonly formula: string
  readonly resolvedCount: number
  readonly unresolvedCount: number
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function stringValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  return null
}

function integerValue(value: unknown): number | null {
  const raw = stringValue(value)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const parsed = Number(raw)
  return Number.isInteger(parsed) ? parsed : null
}

function normalizeSheetName(sheetName: string): string {
  return sheetName.toLocaleLowerCase('en-US')
}

function normalizeCellAddress(address: string): string {
  return address.replaceAll('$', '').toUpperCase()
}

function resolveTargetPath(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.replace(/^\/+/, '')
  }
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(xml)
  return asArray(recordChild(parsed, 'Relationships')?.['Relationship']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['Id'] !== 'string' || typeof entry['Target'] !== 'string' || typeof entry['Type'] !== 'string') {
      return []
    }
    return [{ id: entry['Id'], target: entry['Target'], type: entry['Type'] }]
  })
}

function readSheetNames(externalBook: Record<string, unknown>): string[] {
  return asArray(recordChild(externalBook, 'sheetNames')?.['sheetName']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const name = stringValue(entry['val'])
    return name === null || name.trim().length === 0 ? [] : [name]
  })
}

function readExternalCachedValue(cell: Record<string, unknown>): ExternalCachedValue | null {
  const rawValue = stringValue(cell['v'])
  if (rawValue === null) {
    return null
  }
  const type = stringValue(cell['t'])?.trim().toLowerCase()
  if (type === 'b') {
    return { kind: 'boolean', value: rawValue === '1' || rawValue.toLowerCase() === 'true' }
  }
  if (type === 'e') {
    return { kind: 'error', value: rawValue }
  }
  if (type === 'str' || type === 's' || type === 'inlineStr') {
    return { kind: 'string', value: rawValue }
  }
  const numberValue = Number(rawValue)
  if (Number.isFinite(numberValue)) {
    return { kind: 'number', value: numberValue }
  }
  return { kind: 'string', value: rawValue }
}

function readExternalLinkCache(xml: string): ExternalCachedSheets {
  const parsed: unknown = xmlParser.parse(xml)
  const externalBook = recordChild(recordChild(parsed, 'externalLink'), 'externalBook')
  if (!externalBook) {
    return new Map()
  }
  const sheetNames = readSheetNames(externalBook)
  const sheetCaches: ExternalCachedSheets = new Map()
  for (const sheetData of asArray(recordChild(externalBook, 'sheetDataSet')?.['sheetData'])) {
    if (!isRecord(sheetData)) {
      continue
    }
    const sheetId = integerValue(sheetData['sheetId'])
    if (sheetId === null) {
      continue
    }
    const sheetName = sheetNames[sheetId]
    if (!sheetName) {
      continue
    }
    const cells: ExternalCachedCells = new Map()
    for (const row of asArray(sheetData['row'])) {
      if (!isRecord(row)) {
        continue
      }
      for (const cell of asArray(row['cell'])) {
        if (!isRecord(cell)) {
          continue
        }
        const address = stringValue(cell['r'])
        const value = readExternalCachedValue(cell)
        if (address && value) {
          cells.set(normalizeCellAddress(address), value)
        }
      }
    }
    if (cells.size > 0) {
      sheetCaches.set(normalizeSheetName(sheetName), cells)
    }
  }
  return sheetCaches
}

function readWorkbookExternalLinkTargets(zip: Record<string, Uint8Array>): Map<number, string> {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return new Map()
  }
  const relationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const parsed: unknown = xmlParser.parse(workbookXml)
  const targets = new Map<number, string>()
  let bookIndex = 1
  for (const entry of asArray(recordChild(recordChild(parsed, 'workbook'), 'externalReferences')?.['externalReference'])) {
    if (isRecord(entry) && typeof entry['id'] === 'string') {
      const relationship = relationships.find(
        (candidate) => candidate.id === entry['id'] && candidate.type === externalLinkRelationshipType,
      )
      if (relationship) {
        targets.set(bookIndex, resolveTargetPath('xl/workbook.xml', relationship.target))
      }
    }
    bookIndex += 1
  }
  return targets
}

function readFallbackExternalLinkTargets(zip: Record<string, Uint8Array>): Map<number, string> {
  const targets = new Map<number, string>()
  for (const path of Object.keys(zip)) {
    const match = /^xl\/externalLinks\/externalLink([1-9][0-9]*)\.xml$/u.exec(path)
    if (match) {
      targets.set(Number(match[1]), path)
    }
  }
  return targets
}

export function readImportedExternalLinkCaches(source: XlsxZipSource): ImportedExternalLinkCaches {
  const zip = readXlsxZipEntries(source)
  const caches: ImportedExternalLinkCaches = new Map()
  const workbookTargets = readWorkbookExternalLinkTargets(zip)
  const linkTargets = workbookTargets.size > 0 ? workbookTargets : readFallbackExternalLinkTargets(zip)
  for (const [bookIndex, path] of [...linkTargets.entries()].toSorted((left, right) => left[0] - right[0])) {
    const xml = getZipText(zip, path)
    if (!xml) {
      continue
    }
    const cache = readExternalLinkCache(xml)
    if (cache.size > 0) {
      caches.set(bookIndex, cache)
    }
  }
  return caches
}

function skipDoubleQuotedString(source: string, startIndex: number): number {
  let index = startIndex + 1
  while (index < source.length) {
    if (source[index] === '"') {
      if (source[index + 1] === '"') {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return source.length
}

function readSingleQuotedIdentifier(source: string, startIndex: number): { readonly value: string; readonly endIndex: number } | null {
  let index = startIndex + 1
  let value = ''
  while (index < source.length) {
    const current = source[index]
    if (current === "'") {
      if (source[index + 1] === "'") {
        value += "'"
        index += 2
        continue
      }
      return { value, endIndex: index + 1 }
    }
    value += current ?? ''
    index += 1
  }
  return null
}

function readCellAddress(source: string, startIndex: number): { readonly address: string; readonly endIndex: number } | null {
  const match = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}/u.exec(source.slice(startIndex))
  if (!match) {
    return null
  }
  const address = match[0]
  const endIndex = startIndex + address.length
  if (/^[A-Za-z0-9_.$]/u.test(source[endIndex] ?? '')) {
    return null
  }
  return { address, endIndex }
}

function readExternalSheetName(value: string): { readonly bookIndex: number; readonly sheetName: string } | null {
  const match = /^\[([1-9][0-9]*)\](.+)$/u.exec(value)
  if (!match) {
    return null
  }
  return { bookIndex: Number(match[1]), sheetName: match[2] ?? '' }
}

function readUnquotedExternalSheetReference(
  source: string,
  startIndex: number,
): { readonly bookIndex: number; readonly sheetName: string; readonly endIndex: number } | null {
  const prefix = /^\[([1-9][0-9]*)\]/u.exec(source.slice(startIndex))
  if (!prefix) {
    return null
  }
  const bookIndex = Number(prefix[1])
  let index = startIndex + prefix[0].length
  const sheetStart = index
  while (index < source.length && /[A-Za-z0-9_.-]/u.test(source[index] ?? '')) {
    index += 1
  }
  if (index === sheetStart || source[index] !== '!') {
    return null
  }
  return { bookIndex, sheetName: source.slice(sheetStart, index), endIndex: index }
}

function isRangeOrSpillReference(source: string, addressStartIndex: number, addressEndIndex: number): boolean {
  return source[addressStartIndex - 1] === ':' || source[addressEndIndex] === ':' || source[addressEndIndex] === '#'
}

function formatFormulaLiteral(value: ExternalCachedValue): string {
  switch (value.kind) {
    case 'number':
      return String(value.value)
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE'
    case 'error':
      return value.value
    case 'string':
      return `"${value.value.replaceAll('"', '""')}"`
  }
}

function cachedExternalValue(
  caches: ImportedExternalLinkCaches,
  bookIndex: number,
  sheetName: string,
  address: string,
): ExternalCachedValue | null {
  return caches.get(bookIndex)?.get(normalizeSheetName(sheetName))?.get(normalizeCellAddress(address)) ?? null
}

export function translateImportedFormulaExternalReferences(
  formula: string,
  caches: ImportedExternalLinkCaches,
): ImportedFormulaExternalReferenceTranslation {
  if (caches.size === 0 || !formula.includes('[')) {
    return { formula, resolvedCount: 0, unresolvedCount: 0 }
  }
  let output = ''
  let index = 0
  let resolvedCount = 0
  let unresolvedCount = 0

  while (index < formula.length) {
    const character = formula[index]
    if (character === '"') {
      const endIndex = skipDoubleQuotedString(formula, index)
      output += formula.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === "'") {
      const quoted = readSingleQuotedIdentifier(formula, index)
      if (!quoted || formula[quoted.endIndex] !== '!') {
        const endIndex = quoted?.endIndex ?? index + 1
        output += formula.slice(index, endIndex)
        index = endIndex
        continue
      }
      const externalSheet = readExternalSheetName(quoted.value)
      const address = readCellAddress(formula, quoted.endIndex + 1)
      if (!externalSheet || !address || isRangeOrSpillReference(formula, quoted.endIndex + 1, address.endIndex)) {
        output += formula.slice(index, address?.endIndex ?? quoted.endIndex + 1)
        index = address?.endIndex ?? quoted.endIndex + 1
        continue
      }
      const value = cachedExternalValue(caches, externalSheet.bookIndex, externalSheet.sheetName, address.address)
      if (value) {
        output += formatFormulaLiteral(value)
        resolvedCount += 1
      } else {
        output += formula.slice(index, address.endIndex)
        unresolvedCount += 1
      }
      index = address.endIndex
      continue
    }
    if (character === '[') {
      const externalSheet = readUnquotedExternalSheetReference(formula, index)
      const address = externalSheet ? readCellAddress(formula, externalSheet.endIndex + 1) : null
      if (!externalSheet || !address || isRangeOrSpillReference(formula, externalSheet.endIndex + 1, address.endIndex)) {
        output += character
        index += 1
        continue
      }
      const value = cachedExternalValue(caches, externalSheet.bookIndex, externalSheet.sheetName, address.address)
      if (value) {
        output += formatFormulaLiteral(value)
        resolvedCount += 1
      } else {
        output += formula.slice(index, address.endIndex)
        unresolvedCount += 1
      }
      index = address.endIndex
      continue
    }
    output += character
    index += 1
  }

  return { formula: output, resolvedCount, unresolvedCount }
}
