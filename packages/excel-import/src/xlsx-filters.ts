import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type { CellRangeRef, WorkbookSnapshot } from '@bilig/protocol'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const worksheetAutoFilterTailElements = [
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

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
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

function insertWorksheetAutoFilter(sheetXml: string, ref: string): string {
  const autoFilter = `<autoFilter ref="${escapeXml(ref)}"/>`
  if (/<autoFilter\b/u.test(sheetXml)) {
    return sheetXml.replace(/<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/u, autoFilter)
  }

  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetAutoFilterTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${autoFilter}${sheetXml.slice(insertIndex)}`
}

export function addExportFiltersToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => (sheet.metadata?.filters ?? []).length > 0)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const filter = (sheet.metadata?.filters ?? []).find((candidate) => candidate.sheetName === sheet.name)
      const ref = filter ? rangeRefA1(filter) : null
      if (!ref) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetAutoFilter(sheetXml, ref))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

export function readImportedWorkbookFilters(bytes: Uint8Array, sheetNames: readonly string[]): Map<string, CellRangeRef[]> {
  const zip = unzipSync(bytes)
  const filtersBySheet = new Map<string, CellRangeRef[]>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const autoFilters = asArray(recordChild(recordChild(parsed, 'worksheet'), 'autoFilter')).flatMap((entry) => {
      if (!isRecord(entry) || typeof entry['ref'] !== 'string') {
        return []
      }
      const range = parseRangeRef(sheetName, entry['ref'])
      return range ? [range] : []
    })
    if (autoFilters.length > 0) {
      filtersBySheet.set(sheetName, autoFilters)
    }
  })

  return filtersBySheet
}
