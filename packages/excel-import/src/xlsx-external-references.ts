import { XMLParser } from 'fast-xml-parser'

import type { LiteralInput, WorkbookExternalWorkbookReferenceSnapshot, WorkbookSnapshot } from '@bilig/protocol'
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
interface ExternalCachedSheet {
  readonly sheetName: string
  readonly cells: ExternalCachedCells
}
type ExternalCachedSheets = Map<string, ExternalCachedSheet>
export type ImportedExternalLinkCaches = Map<number, ExternalCachedSheets>

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
  readonly targetMode?: string
}

export interface ImportedFormulaExternalReferenceTranslation {
  readonly formula: string
  readonly resolvedCount: number
  readonly unresolvedCount: number
  readonly materializedExternalCacheSheetKeys: readonly string[]
}

export type ImportedExternalWorkbookReferences = Map<number, WorkbookExternalWorkbookReferenceSnapshot>
export type ImportedExternalCacheSheetMap = ReadonlyMap<string, string>

export interface ImportedExternalCacheSheetSnapshot {
  readonly key: string
  readonly name: string
  readonly cells: WorkbookSnapshot['sheets'][number]['cells']
}

export interface ImportedExternalCacheSheetPlan {
  readonly sheetsByExternalSheet: ReadonlyMap<string, ImportedExternalCacheSheetSnapshot>
  readonly sheetNamesByExternalSheet: ImportedExternalCacheSheetMap
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const externalLinkPathRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'

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

interface ParsedCellAddress {
  readonly row: number
  readonly col: number
}

interface CachedExternalRange {
  readonly formulaLiteral: string
  readonly resolvedCount: number
}

interface MaterializedCachedExternalRange {
  readonly formulaReference: string
  readonly resolvedCount: number
  readonly sheetKey: string
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
    return [
      {
        id: entry['Id'],
        target: entry['Target'],
        type: entry['Type'],
        ...(typeof entry['TargetMode'] === 'string' ? { targetMode: entry['TargetMode'] } : {}),
      },
    ]
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
      sheetCaches.set(normalizeSheetName(sheetName), { sheetName, cells })
    }
  }
  return sheetCaches
}

function externalCachedValueToLiteralInput(value: ExternalCachedValue): LiteralInput | undefined {
  switch (value.kind) {
    case 'number':
    case 'string':
    case 'boolean':
      return value.value
    case 'error':
      return undefined
  }
}

function sheetNameKey(sheetName: string): string {
  return sheetName.toLocaleLowerCase('en-US')
}

function sanitizeWorksheetNamePart(value: string): string {
  const sanitized = value
    .replace(/[:\\/?*[\]]/gu, '_')
    .replace(/'/gu, '_')
    .replace(/\s+/gu, '_')
    .replace(/_+/gu, '_')
    .replace(/^_+|_+$/gu, '')
  return sanitized.length > 0 ? sanitized : 'Sheet'
}

function uniqueExternalCacheSheetName(bookIndex: number, sheetName: string, usedSheetNames: Set<string>): string {
  const prefix = `__bilig_ext_${bookIndex}_`
  const suffix = sanitizeWorksheetNamePart(sheetName)
  const base = `${prefix}${suffix}`.slice(0, 31)
  let candidate = base.length > 0 ? base : '__bilig_ext'
  let disambiguator = 2
  while (usedSheetNames.has(sheetNameKey(candidate))) {
    const suffixText = `_${disambiguator}`
    candidate = `${base.slice(0, Math.max(0, 31 - suffixText.length))}${suffixText}`
    disambiguator += 1
  }
  usedSheetNames.add(sheetNameKey(candidate))
  return candidate
}

function compareExternalCacheCellAddresses(left: string, right: string): number {
  const leftAddress = parseCachedCellAddress(left)
  const rightAddress = parseCachedCellAddress(right)
  if (!leftAddress || !rightAddress) {
    return left.localeCompare(right)
  }
  return leftAddress.row - rightAddress.row || leftAddress.col - rightAddress.col
}

function externalCacheSheetKey(bookIndex: number, sheetName: string): string {
  return `${bookIndex}\u0000${normalizeSheetName(sheetName)}`
}

function buildImportedExternalCacheSheetCells(sheet: ExternalCachedSheet): WorkbookSnapshot['sheets'][number]['cells'] {
  return [...sheet.cells.entries()]
    .flatMap(([address, value]) => {
      const literal = externalCachedValueToLiteralInput(value)
      const parsed = parseCachedCellAddress(address)
      if (literal === undefined || !parsed) {
        return []
      }
      return [
        {
          address: normalizeCellAddress(address),
          row: parsed.row,
          col: parsed.col,
          value: literal,
        },
      ]
    })
    .toSorted((left, right) => compareExternalCacheCellAddresses(left.address, right.address))
}

export function buildImportedExternalCacheSheetPlan(
  caches: ImportedExternalLinkCaches,
  existingSheetNames: readonly string[],
): ImportedExternalCacheSheetPlan {
  const usedSheetNames = new Set(existingSheetNames.map(sheetNameKey))
  const sheetsByExternalSheet = new Map<string, ImportedExternalCacheSheetSnapshot>()
  const sheetNamesByExternalSheet = new Map<string, string>()

  for (const [bookIndex, sheets] of [...caches.entries()].toSorted((left, right) => left[0] - right[0])) {
    for (const sheet of [...sheets.values()].toSorted((left, right) => left.sheetName.localeCompare(right.sheetName))) {
      const cells = buildImportedExternalCacheSheetCells(sheet)
      if (cells.length === 0) {
        continue
      }
      const key = externalCacheSheetKey(bookIndex, sheet.sheetName)
      const name = uniqueExternalCacheSheetName(bookIndex, sheet.sheetName, usedSheetNames)
      const snapshot: ImportedExternalCacheSheetSnapshot = { key, name, cells }
      sheetsByExternalSheet.set(key, snapshot)
      sheetNamesByExternalSheet.set(key, name)
    }
  }

  return { sheetsByExternalSheet, sheetNamesByExternalSheet }
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

function externalLinkRelationshipsPartPath(partPath: string): string {
  const fileName = partPath.slice(partPath.lastIndexOf('/') + 1)
  return `xl/externalLinks/_rels/${fileName}.rels`
}

function readExternalBookRelationshipId(xml: string): string | null {
  const parsed: unknown = xmlParser.parse(xml)
  const externalBook = recordChild(recordChild(parsed, 'externalLink'), 'externalBook')
  return externalBook ? stringValue(externalBook['id']) : null
}

function readExternalBookSheetNames(xml: string): string[] {
  const parsed: unknown = xmlParser.parse(xml)
  const externalBook = recordChild(recordChild(parsed, 'externalLink'), 'externalBook')
  return externalBook ? readSheetNames(externalBook) : []
}

function workbookNameFromExternalTarget(target: string | undefined): string | undefined {
  if (!target) {
    return undefined
  }
  const targetPath = target.split(/[?#]/u)[0] ?? target
  const normalizedTarget = targetPath.replace(/^file:\/+/iu, '').replace(/\\/gu, '/')
  const targetSegments = normalizedTarget.split('/')
  let lastSegment: string | undefined
  for (let index = targetSegments.length - 1; index >= 0; index -= 1) {
    const segment = targetSegments[index]
    if (segment && segment.length > 0) {
      lastSegment = segment
      break
    }
  }
  if (!lastSegment) {
    return undefined
  }
  try {
    return decodeURIComponent(lastSegment)
  } catch {
    return lastSegment
  }
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

export function readImportedExternalWorkbookReferences(source: XlsxZipSource): ImportedExternalWorkbookReferences {
  const zip = readXlsxZipEntries(source)
  const references: ImportedExternalWorkbookReferences = new Map()
  const workbookTargets = readWorkbookExternalLinkTargets(zip)
  const linkTargets = workbookTargets.size > 0 ? workbookTargets : readFallbackExternalLinkTargets(zip)
  for (const [bookIndex, path] of [...linkTargets.entries()].toSorted((left, right) => left[0] - right[0])) {
    const xml = getZipText(zip, path)
    if (!xml) {
      continue
    }
    const relationships = parseRelationships(getZipText(zip, externalLinkRelationshipsPartPath(path)))
    const externalBookRelationshipId = readExternalBookRelationshipId(xml)
    const linkedWorkbookRelationship =
      (externalBookRelationshipId ? relationships.find((relationship) => relationship.id === externalBookRelationshipId) : undefined) ??
      relationships.find((relationship) => relationship.type === externalLinkPathRelationshipType)
    const target = linkedWorkbookRelationship?.target
    const workbookName = workbookNameFromExternalTarget(target)
    references.set(bookIndex, {
      bookIndex,
      packagePath: path,
      ...(target ? { target } : {}),
      ...(linkedWorkbookRelationship?.targetMode ? { targetMode: linkedWorkbookRelationship.targetMode } : {}),
      ...(workbookName ? { workbookName } : {}),
      sheetNames: readExternalBookSheetNames(xml),
    })
  }
  return references
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

export function collectImportedFormulaExternalWorkbookReferences(
  formula: string,
  references: ImportedExternalWorkbookReferences,
): readonly WorkbookExternalWorkbookReferenceSnapshot[] {
  if (!formula.includes('[')) {
    return []
  }
  const bookIndices = new Set<number>()
  let index = 0
  while (index < formula.length) {
    const character = formula[index]
    if (character === '"') {
      index = skipDoubleQuotedString(formula, index)
      continue
    }
    if (character === "'") {
      const quoted = readSingleQuotedIdentifier(formula, index)
      const externalSheet = quoted ? readExternalSheetName(quoted.value) : null
      if (externalSheet) {
        bookIndices.add(externalSheet.bookIndex)
      }
      index = quoted?.endIndex ?? index + 1
      continue
    }
    if (character === '[') {
      const externalSheet = readUnquotedExternalSheetReference(formula, index)
      if (externalSheet) {
        bookIndices.add(externalSheet.bookIndex)
        index = externalSheet.endIndex
        continue
      }
      const externalBook = /^\[([1-9][0-9]*)\]/u.exec(formula.slice(index))
      if (externalBook) {
        bookIndices.add(Number(externalBook[1]))
        index += externalBook[0].length
        continue
      }
    }
    index += 1
  }
  return [...bookIndices].toSorted((left, right) => left - right).map((bookIndex) => references.get(bookIndex) ?? { bookIndex })
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

function parseCachedCellAddress(address: string): ParsedCellAddress | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/u.exec(normalizeCellAddress(address))
  if (!match) {
    return null
  }
  let col = 0
  for (const character of match[1]!) {
    col = col * 26 + character.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  return col > 0 && Number.isSafeInteger(row) ? { row: row - 1, col: col - 1 } : null
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

function readCachedExternalRange(
  caches: ImportedExternalLinkCaches,
  bookIndex: number,
  sheetName: string,
  startAddress: string,
  endAddress: string,
): CachedExternalRange | null {
  const start = parseCachedCellAddress(startAddress)
  const end = parseCachedCellAddress(endAddress)
  if (!start || !end) {
    return null
  }
  const rowStart = Math.min(start.row, end.row)
  const rowEnd = Math.max(start.row, end.row)
  const colStart = Math.min(start.col, end.col)
  const colEnd = Math.max(start.col, end.col)
  const rows: string[][] = []
  for (let row = rowStart; row <= rowEnd; row += 1) {
    const values: string[] = []
    for (let col = colStart; col <= colEnd; col += 1) {
      const value = cachedExternalValue(caches, bookIndex, sheetName, formatA1Address(row, col))
      if (!value) {
        return null
      }
      values.push(formatFormulaLiteral(value))
    }
    rows.push(values)
  }
  return {
    formulaLiteral: `{${rows.map((row) => row.join(',')).join(';')}}`,
    resolvedCount: (rowEnd - rowStart + 1) * (colEnd - colStart + 1),
  }
}

function formatA1Address(row: number, col: number): string {
  let current = col + 1
  let column = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    column = String.fromCharCode(65 + remainder) + column
    current = Math.floor((current - 1) / 26)
  }
  return `${column}${String(row + 1)}`
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
  return caches.get(bookIndex)?.get(normalizeSheetName(sheetName))?.cells.get(normalizeCellAddress(address)) ?? null
}

function quoteFormulaSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function formatAbsoluteA1Address(row: number, col: number): string {
  const address = formatA1Address(row, col)
  const match = /^([A-Z]+)([1-9][0-9]*)$/u.exec(address)
  return match ? `$${match[1]}$${match[2]}` : address
}

function readMaterializedCachedExternalRange(
  caches: ImportedExternalLinkCaches,
  cacheSheetNames: ImportedExternalCacheSheetMap | undefined,
  bookIndex: number,
  sheetName: string,
  startAddress: string,
  endAddress: string,
): MaterializedCachedExternalRange | null {
  if (!cacheSheetNames) {
    return null
  }
  const start = parseCachedCellAddress(startAddress)
  const end = parseCachedCellAddress(endAddress)
  if (!start || !end) {
    return null
  }
  const sheetKey = externalCacheSheetKey(bookIndex, sheetName)
  const localSheetName = cacheSheetNames.get(sheetKey)
  if (!localSheetName) {
    return null
  }
  const rowStart = Math.min(start.row, end.row)
  const rowEnd = Math.max(start.row, end.row)
  const colStart = Math.min(start.col, end.col)
  const colEnd = Math.max(start.col, end.col)
  let resolvedCount = 0
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      const value = cachedExternalValue(caches, bookIndex, sheetName, formatA1Address(row, col))
      if (!value || externalCachedValueToLiteralInput(value) === undefined) {
        return null
      }
      resolvedCount += 1
    }
  }
  const formulaReference = `${quoteFormulaSheetName(localSheetName)}!${formatAbsoluteA1Address(rowStart, colStart)}:${formatAbsoluteA1Address(
    rowEnd,
    colEnd,
  )}`
  return { formulaReference, resolvedCount, sheetKey }
}

function formulaNeedsPivotReferenceContext(formula: string): boolean {
  return /\bGETPIVOTDATA\s*\(/iu.test(formula)
}

function formulaNeedsCriteriaRangeReferenceContext(formula: string): boolean {
  return /\b(?:AVERAGEIF|AVERAGEIFS|COUNTIF|COUNTIFS|MAXIFS|MINIFS|SUMIF|SUMIFS)\s*\(/iu.test(formula)
}

export function translateImportedFormulaExternalReferences(
  formula: string,
  caches: ImportedExternalLinkCaches,
  cacheSheetNames?: ImportedExternalCacheSheetMap,
): ImportedFormulaExternalReferenceTranslation {
  const unchanged = (): ImportedFormulaExternalReferenceTranslation => ({
    formula,
    resolvedCount: 0,
    unresolvedCount: 0,
    materializedExternalCacheSheetKeys: [],
  })
  if (caches.size === 0 || !formula.includes('[')) {
    return unchanged()
  }
  if (formulaNeedsPivotReferenceContext(formula)) {
    return unchanged()
  }
  const needsCriteriaRangeReferenceContext = formulaNeedsCriteriaRangeReferenceContext(formula)
  if (needsCriteriaRangeReferenceContext && !cacheSheetNames) {
    return unchanged()
  }
  let output = ''
  let index = 0
  let resolvedCount = 0
  let unresolvedCount = 0
  const materializedExternalCacheSheetKeys = new Set<string>()

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
      if (!externalSheet || !address) {
        output += formula.slice(index, address?.endIndex ?? quoted.endIndex + 1)
        index = address?.endIndex ?? quoted.endIndex + 1
        continue
      }
      if (formula[address.endIndex] === ':') {
        const endAddress = readCellAddress(formula, address.endIndex + 1)
        if (endAddress) {
          const materializedRange = readMaterializedCachedExternalRange(
            caches,
            cacheSheetNames,
            externalSheet.bookIndex,
            externalSheet.sheetName,
            address.address,
            endAddress.address,
          )
          if (materializedRange) {
            output += materializedRange.formulaReference
            resolvedCount += materializedRange.resolvedCount
            materializedExternalCacheSheetKeys.add(materializedRange.sheetKey)
          } else if (!needsCriteriaRangeReferenceContext) {
            const range = readCachedExternalRange(
              caches,
              externalSheet.bookIndex,
              externalSheet.sheetName,
              address.address,
              endAddress.address,
            )
            if (range) {
              output += range.formulaLiteral
              resolvedCount += range.resolvedCount
            } else {
              output += formula.slice(index, endAddress.endIndex)
              unresolvedCount += 1
            }
          } else {
            output += formula.slice(index, endAddress.endIndex)
            unresolvedCount += 1
          }
          index = endAddress.endIndex
          continue
        }
      }
      if (isRangeOrSpillReference(formula, quoted.endIndex + 1, address.endIndex)) {
        output += formula.slice(index, address.endIndex)
        index = address.endIndex
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
      if (!externalSheet || !address) {
        output += character
        index += 1
        continue
      }
      if (formula[address.endIndex] === ':') {
        const endAddress = readCellAddress(formula, address.endIndex + 1)
        if (endAddress) {
          const materializedRange = readMaterializedCachedExternalRange(
            caches,
            cacheSheetNames,
            externalSheet.bookIndex,
            externalSheet.sheetName,
            address.address,
            endAddress.address,
          )
          if (materializedRange) {
            output += materializedRange.formulaReference
            resolvedCount += materializedRange.resolvedCount
            materializedExternalCacheSheetKeys.add(materializedRange.sheetKey)
          } else if (!needsCriteriaRangeReferenceContext) {
            const range = readCachedExternalRange(
              caches,
              externalSheet.bookIndex,
              externalSheet.sheetName,
              address.address,
              endAddress.address,
            )
            if (range) {
              output += range.formulaLiteral
              resolvedCount += range.resolvedCount
            } else {
              output += formula.slice(index, endAddress.endIndex)
              unresolvedCount += 1
            }
          } else {
            output += formula.slice(index, endAddress.endIndex)
            unresolvedCount += 1
          }
          index = endAddress.endIndex
          continue
        }
      }
      if (isRangeOrSpillReference(formula, externalSheet.endIndex + 1, address.endIndex)) {
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

  return {
    formula: output,
    resolvedCount,
    unresolvedCount,
    materializedExternalCacheSheetKeys: [...materializedExternalCacheSheetKeys].toSorted(),
  }
}
