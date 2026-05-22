import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type { WorkbookSheetProtectionSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { escapeXmlAttribute } from './xlsx-export-xml.js'
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

const xmlNamePattern = /^[A-Za-z_:][\w:.-]*$/u
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

function readSheetProtectionXmlAttributes(sheetXml: string): WorkbookSheetProtectionSnapshot['xmlAttributes'] | undefined {
  const match = /<sheetProtection\b[^>]*(?:\/>|>[\s\S]*?<\/sheetProtection>)/u.exec(sheetXml)
  if (!match) {
    return undefined
  }
  const attributes = [...match[0].matchAll(/\s([A-Za-z_:][\w:.-]*)=(?:"([^"]*)"|'([^']*)')/gu)].map((attributeMatch) => ({
    name: attributeMatch[1] ?? '',
    value: unescapeXmlAttribute(attributeMatch[2] ?? attributeMatch[3] ?? ''),
  }))
  if (attributes.length === 1 && attributes[0]?.name === 'sheet' && attributes[0].value === '1') {
    return undefined
  }
  return attributes.length > 0 ? attributes : undefined
}

function buildWorksheetSheetProtectionXml(protection: WorkbookSheetProtectionSnapshot): string {
  const attributes = (protection.xmlAttributes ?? []).filter((attribute) => xmlNamePattern.test(attribute.name))
  if (attributes.length === 0) {
    return '<sheetProtection sheet="1"/>'
  }
  return `<sheetProtection${attributes.map((attribute) => ` ${attribute.name}="${escapeXmlAttribute(attribute.value)}"`).join('')}/>`
}

function insertWorksheetSheetProtection(sheetXml: string, protection: WorkbookSheetProtectionSnapshot): string {
  const sheetProtection = buildWorksheetSheetProtectionXml(protection)
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
      const sheetProtection = sheet.metadata.sheetProtection
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetSheetProtection(sheetXml, sheetProtection))
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
    const xmlAttributes = readSheetProtectionXmlAttributes(sheetXml)
    protectionsBySheet.set(sheetName, {
      sheetName,
      ...(xmlAttributes ? { xmlAttributes } : {}),
    })
  })

  return protectionsBySheet
}
