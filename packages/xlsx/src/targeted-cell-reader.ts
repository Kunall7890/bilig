import { normalizeCellAddress } from './address.js'
import { workbookSheetPathEntriesFromSource } from './workbook-sheet-paths.js'
import { decodeXmlText, getXmlElementText, readXmlAttribute, worksheetCellElementPattern } from './xml.js'
import { getZipText, readXlsxZipEntries, type XlsxZipSource } from './zip-reader.js'

export type XlsxTargetCellValue = string | number | boolean

export interface XlsxTargetCellReadback {
  readonly address: string
  readonly formula: string | null
  readonly rawValue: string | null
  readonly styleIndex: number | null
  readonly type: string | null
  readonly value: XlsxTargetCellValue | null
}

export function readXlsxTargetCell(source: XlsxZipSource, sheetName: string, address: string): XlsxTargetCellReadback | null {
  const zip = readXlsxZipEntries(source)
  const sheetPath = workbookSheetPathEntriesFromSource(zip, [sheetName])[0]?.path
  if (!sheetPath) {
    throw new Error(`XLSX workbook is missing sheet: ${sheetName}`)
  }
  const sheetXml = getZipText(zip, sheetPath)
  if (!sheetXml) {
    throw new Error(`XLSX workbook is missing worksheet XML for sheet: ${sheetName}`)
  }
  const normalizedAddress = normalizeCellAddress(address)
  const cellXml = readWorksheetCellXml(sheetXml, normalizedAddress)
  if (!cellXml) {
    return null
  }
  return readTargetCellXml(cellXml, normalizedAddress, (index) => readSharedStringAt(zip, index))
}

function readTargetCellXml(
  cellXml: string,
  normalizedAddress: string,
  sharedStringAt: (index: number) => string | null,
): XlsxTargetCellReadback {
  const openingTag = /^<[^>]+>/u.exec(cellXml)?.[0] ?? ''
  const type = readXmlAttribute(openingTag, 't')
  const rawValue = getXmlElementText(cellXml, 'v')
  const styleIndex = readCellStyleIndex(openingTag)
  const formulaText = getXmlElementText(cellXml, 'f')
  return {
    address: normalizedAddress,
    formula: formulaText === null ? null : decodeXmlText(formulaText),
    rawValue: rawValue === null ? null : decodeXmlText(rawValue),
    styleIndex,
    type,
    value: readTargetCellValue(cellXml, type, rawValue, sharedStringAt),
  }
}

function readWorksheetCellXml(sheetXml: string, normalizedAddress: string): string | null {
  for (const match of sheetXml.matchAll(worksheetCellElementPattern)) {
    const cellXml = match[0]
    const openingTag = /^<[^>]+>/u.exec(cellXml)?.[0] ?? ''
    if (readXmlAttribute(openingTag, 'r') === normalizedAddress) {
      return cellXml
    }
  }
  return null
}

function readTargetCellValue(
  cellXml: string,
  type: string | null,
  rawValue: string | null,
  sharedStringAt: (index: number) => string | null,
): XlsxTargetCellValue | null {
  if (type === 'inlineStr') {
    const inlineText = readRichTextValue(cellXml)
    return inlineText.length > 0 ? inlineText : null
  }
  if (rawValue === null) {
    return null
  }
  const value = decodeXmlText(rawValue)
  if (type === 's') {
    const sharedStringIndex = Number(value)
    return Number.isSafeInteger(sharedStringIndex) && sharedStringIndex >= 0 ? sharedStringAt(sharedStringIndex) : null
  }
  if (type === 'b') {
    return value === '1' || value.toLowerCase() === 'true'
  }
  if (type === 'str' || type === 'e') {
    return value
  }
  const numericValue = Number(value)
  return Number.isFinite(numericValue) ? numericValue : value
}

function readSharedStringAt(zip: XlsxZipSource, index: number): string | null {
  const xml = getZipText(readXlsxZipEntries(zip), 'xl/sharedStrings.xml')
  if (!xml) {
    return null
  }
  let currentIndex = 0
  for (const match of xml.matchAll(/<si\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/si>/gu)) {
    if (currentIndex === index) {
      return readRichTextValue(match[0])
    }
    currentIndex += 1
  }
  return null
}

function readRichTextValue(xml: string): string {
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
    .map((match) => decodeXmlText(match[1] ?? ''))
    .join('')
}

function readCellStyleIndex(openingTag: string): number | null {
  const styleIndex = Number(readXmlAttribute(openingTag, 's'))
  return Number.isInteger(styleIndex) && styleIndex >= 0 ? styleIndex : null
}
