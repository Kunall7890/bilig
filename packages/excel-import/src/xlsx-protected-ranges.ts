import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type { CellRangeRef, WorkbookRangeProtectionSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { escapeXmlAttribute } from './xlsx-export-xml.js'
import { workbookSheetPathEntriesFromSource } from './xlsx-workbook-sheet-paths.js'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const worksheetProtectedRangesTailElements = [
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'drawing',
  'legacyDrawing',
  'tableParts',
  'pivotTableDefinition',
] as const

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

const xmlNamePattern = /^[A-Za-z_:][\w:.-]*$/u
const protectedRangeSemanticAttributes = new Set(['name', 'sqref'])
const xmlNamedEntities: Readonly<Record<string, string>> = {
  amp: '&',
  apos: "'",
  gt: '>',
  lt: '<',
  quot: '"',
}

function unescapeXmlAttribute(value: string): string {
  return value.replace(
    /&#x([0-9a-fA-F]+);|&#([0-9]+);|&(quot|apos|lt|gt|amp);/gu,
    (match, hex: string | undefined, decimal: string | undefined, named: string | undefined) => {
      if (hex || decimal) {
        const codePoint = Number.parseInt(hex ?? decimal ?? '', hex ? 16 : 10)
        return Number.isSafeInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : match
      }
      return named ? (xmlNamedEntities[named] ?? match) : match
    },
  )
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getZipText(zip: ZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

function setZipText(zip: ZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function rangeRefA1(range: CellRangeRef): string | null {
  try {
    const decoded = XLSX.utils.decode_range(`${range.startAddress}:${range.endAddress}`.replaceAll('$', ''))
    return XLSX.utils.encode_range(decoded)
  } catch {
    return null
  }
}

function parseRangeRef(sheetName: string, ref: string): CellRangeRef | null {
  try {
    const decoded = XLSX.utils.decode_range(ref.replaceAll('$', ''))
    return {
      sheetName,
      startAddress: XLSX.utils.encode_cell(decoded.s),
      endAddress: XLSX.utils.encode_cell(decoded.e),
    }
  } catch {
    return null
  }
}

function insertWorksheetProtectedRanges(sheetXml: string, rangeXml: readonly string[]): string {
  const protectedRanges = `<protectedRanges>${rangeXml.join('')}</protectedRanges>`
  if (/<protectedRanges\b/u.test(sheetXml)) {
    return sheetXml.replace(/<protectedRanges\b[^>]*(?:\/>|>[\s\S]*?<\/protectedRanges>)/u, protectedRanges)
  }

  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetProtectedRangesTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${protectedRanges}${sheetXml.slice(insertIndex)}`
}

function exportProtectedRangeXml(sheetName: string, protection: WorkbookRangeProtectionSnapshot): string | null {
  const name = protection.id.trim()
  const ref = protection.range.sheetName === sheetName ? rangeRefA1(protection.range) : null
  if (!name || !ref) {
    return null
  }
  const attributes = [
    { name: 'name', value: name },
    { name: 'sqref', value: ref },
    ...(protection.xmlAttributes ?? []).filter(
      (attribute) => xmlNamePattern.test(attribute.name) && !protectedRangeSemanticAttributes.has(attribute.name),
    ),
  ]
  return `<protectedRange${attributes.map((attribute) => ` ${attribute.name}="${escapeXmlAttribute(attribute.value)}"`).join('')}/>`
}

export function addExportProtectedRangesToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => (sheet.metadata?.protectedRanges ?? []).length > 0)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const rangeXml = (sheet.metadata?.protectedRanges ?? []).flatMap((protection) => {
        const xml = exportProtectedRangeXml(sheet.name, protection)
        return xml ? [xml] : []
      })
      if (rangeXml.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetProtectedRanges(sheetXml, rangeXml))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

function uniqueProtectionId(input: { sheetName: string; name: unknown; range: CellRangeRef; usedIds: Set<string> }): string {
  const name = typeof input.name === 'string' ? input.name.trim() : ''
  const base = name || `protected-range:${input.sheetName}:${input.range.startAddress}:${input.range.endAddress}`
  let candidate = base
  let suffix = 2
  while (input.usedIds.has(candidate)) {
    candidate = `${base}:${String(suffix)}`
    suffix += 1
  }
  input.usedIds.add(candidate)
  return candidate
}

function readProtectedRangeXmlAttributes(sheetXml: string): Array<WorkbookRangeProtectionSnapshot['xmlAttributes']> {
  return [
    ...sheetXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?protectedRange\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?protectedRange>)/gu),
  ].map((match) => {
    const attributes = [...match[0].matchAll(/\s([A-Za-z_:][\w:.-]*)=(?:"([^"]*)"|'([^']*)')/gu)]
      .map((attributeMatch) => ({
        name: attributeMatch[1] ?? '',
        value: unescapeXmlAttribute(attributeMatch[2] ?? attributeMatch[3] ?? ''),
      }))
      .filter((attribute) => !protectedRangeSemanticAttributes.has(attribute.name))
    return attributes.length > 0 ? attributes : undefined
  })
}

function parseProtectedRangeEntry(
  sheetName: string,
  entry: unknown,
  usedIds: Set<string>,
  xmlAttributes: WorkbookRangeProtectionSnapshot['xmlAttributes'],
): WorkbookRangeProtectionSnapshot[] {
  if (!isRecord(entry) || typeof entry['sqref'] !== 'string') {
    return []
  }
  return entry['sqref'].split(/\s+/u).flatMap((rawRef) => {
    const ref = rawRef.trim()
    if (!ref) {
      return []
    }
    const range = parseRangeRef(sheetName, ref)
    if (!range) {
      return []
    }
    return [
      {
        id: uniqueProtectionId({ sheetName, name: entry['name'], range, usedIds }),
        range,
        ...(xmlAttributes ? { xmlAttributes: structuredClone(xmlAttributes) } : {}),
      },
    ]
  })
}

export function readImportedWorkbookProtectedRanges(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookRangeProtectionSnapshot[]> {
  const zip = readXlsxZipEntries(source)
  const protectedRangesBySheet = new Map<string, WorkbookRangeProtectionSnapshot[]>()

  workbookSheetPathEntriesFromSource(zip, sheetNames).forEach((sheet) => {
    const sheetXml = getZipText(zip, sheet.path)
    if (!sheetXml || !/<protectedRanges\b/u.test(sheetXml)) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const xmlAttributes = readProtectedRangeXmlAttributes(sheetXml)
    const usedIds = new Set<string>()
    const protectedRanges = asArray(
      recordChild(recordChild(recordChild(parsed, 'worksheet'), 'protectedRanges'), 'protectedRange'),
    ).flatMap((entry, index) => parseProtectedRangeEntry(sheet.name, entry, usedIds, xmlAttributes[index]))
    if (protectedRanges.length > 0) {
      protectedRangesBySheet.set(sheet.name, protectedRanges)
    }
  })

  return protectedRangesBySheet
}
