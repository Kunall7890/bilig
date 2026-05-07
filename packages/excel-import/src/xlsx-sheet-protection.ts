import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type { WorkbookSheetProtectionSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const worksheetProtectionTailElements = [
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
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

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
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

function isFalseAttribute(value: unknown): boolean {
  return value === false || value === '0' || value === 'false'
}

function insertWorksheetSheetProtection(sheetXml: string): string {
  const sheetProtection = '<sheetProtection sheet="1"/>'
  if (/<sheetProtection\b/u.test(sheetXml)) {
    return sheetXml.replace(/<sheetProtection\b[^>]*(?:\/>|>[\s\S]*?<\/sheetProtection>)/u, sheetProtection)
  }

  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetProtectionTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${sheetProtection}${sheetXml.slice(insertIndex)}`
}

export function addExportSheetProtectionsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => sheet.metadata?.sheetProtection?.sheetName === sheet.name)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      if (sheet.metadata?.sheetProtection?.sheetName !== sheet.name) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetSheetProtection(sheetXml))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

export function readImportedWorkbookSheetProtections(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetProtectionSnapshot> {
  const zip = readXlsxZipEntries(source)
  const protectionsBySheet = new Map<string, WorkbookSheetProtectionSnapshot>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !/<sheetProtection\b/u.test(sheetXml)) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const sheetProtection = recordChild(recordChild(parsed, 'worksheet'), 'sheetProtection')
    if (!sheetProtection || isFalseAttribute(sheetProtection['sheet'])) {
      return
    }
    protectionsBySheet.set(sheetName, { sheetName })
  })

  return protectionsBySheet
}
