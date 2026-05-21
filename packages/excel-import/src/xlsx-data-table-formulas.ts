import { unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { WorkbookSheetDataTableFormulasSnapshot, WorkbookSheetDataTableFormulaSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import { escapeXml, setZipText } from './xlsx-pivot-artifacts.js'

const cellElementPattern = /<c\b(?<attributes>[^>]*)>(?<body>[\s\S]*?)<\/c>/gu
const formulaElementPattern = /<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>)/u
const formulaElementGlobalPattern = /<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/gu

export interface ImportedDataTableFormulaCells {
  readonly formulaCells: Map<string, string>
  readonly unsupportedCount: number
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function decodeXmlAttribute(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function cellColumnIndex(address: string): number {
  try {
    return XLSX.utils.decode_cell(address).c
  } catch {
    return Number.MAX_SAFE_INTEGER
  }
}

function cellRowNumber(address: string): number | null {
  const match = /^[A-Z]+([1-9][0-9]*)$/u.exec(address)
  return match ? Number(match[1]) : null
}

function readDataTableFormulaSnapshots(sheetXml: string | null): WorkbookSheetDataTableFormulaSnapshot[] {
  if (!sheetXml) {
    return []
  }
  if (!sheetXml.includes('t="dataTable"') && !sheetXml.includes("t='dataTable'")) {
    return []
  }
  const formulas: WorkbookSheetDataTableFormulaSnapshot[] = []
  cellElementPattern.lastIndex = 0
  for (const match of sheetXml.matchAll(cellElementPattern)) {
    const attributes = match.groups?.['attributes'] ?? ''
    const body = match.groups?.['body'] ?? ''
    const address = readAttribute(attributes, 'r')
    formulaElementGlobalPattern.lastIndex = 0
    const formulaXml = [...body.matchAll(formulaElementGlobalPattern)].find((formulaMatch) => {
      return readAttribute(formulaMatch[0], 't') === 'dataTable'
    })?.[0]
    if (address && formulaXml) {
      formulas.push({ address, formulaXml })
    }
  }
  return formulas
}

function readFormulaAttribute(formulaXml: string, attributeName: string): string | null {
  const value = readAttribute(formulaXml, attributeName)
  return value === null ? null : decodeXmlAttribute(value)
}

function isEnabledXmlFlag(value: string | null): boolean {
  return value === '1' || value === 'true'
}

function decodeFormulaRefRange(formula: WorkbookSheetDataTableFormulaSnapshot): XLSX.Range | null {
  const ref = readFormulaAttribute(formula.formulaXml, 'ref')
  if (!ref) {
    return null
  }
  try {
    return XLSX.utils.decode_range(ref)
  } catch {
    return null
  }
}

function buildTwoVariableDataTableFormulaCells(formula: WorkbookSheetDataTableFormulaSnapshot): Map<string, string> | null {
  if (!isEnabledXmlFlag(readFormulaAttribute(formula.formulaXml, 'dt2D'))) {
    return null
  }
  const rowInput = readFormulaAttribute(formula.formulaXml, 'r1')
  const columnInput = readFormulaAttribute(formula.formulaXml, 'r2')
  const range = decodeFormulaRefRange(formula)
  if (!rowInput || !columnInput || !range || range.s.r < 1 || range.s.c < 1) {
    return null
  }
  if (formula.address !== XLSX.utils.encode_cell(range.s)) {
    return null
  }

  const sourceFormulaAddress = XLSX.utils.encode_cell({ r: range.s.r - 1, c: range.s.c - 1 })
  const formulaCells = new Map<string, string>()
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const columnReplacement = XLSX.utils.encode_cell({ r: row, c: range.s.c - 1 })
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column })
      const rowReplacement = XLSX.utils.encode_cell({ r: range.s.r - 1, c: column })
      formulaCells.set(
        address,
        `MULTIPLE.OPERATIONS(${sourceFormulaAddress},${rowInput},${rowReplacement},${columnInput},${columnReplacement})`,
      )
    }
  }
  return formulaCells
}

export function buildImportedDataTableFormulaCells(
  dataTableFormulas: WorkbookSheetDataTableFormulasSnapshot | undefined,
): ImportedDataTableFormulaCells {
  const formulaCells = new Map<string, string>()
  let unsupportedCount = 0
  for (const formula of dataTableFormulas?.formulas ?? []) {
    const loweredFormulaCells = buildTwoVariableDataTableFormulaCells(formula)
    if (!loweredFormulaCells) {
      unsupportedCount += 1
      continue
    }
    for (const [address, loweredFormula] of loweredFormulaCells) {
      formulaCells.set(address, loweredFormula)
    }
  }
  return { formulaCells, unsupportedCount }
}

function insertFormulaIntoCellBody(body: string, formulaXml: string): string {
  if (formulaElementPattern.test(body)) {
    return body.replace(formulaElementPattern, formulaXml)
  }
  const valueIndex = body.search(/<(?:v|is|extLst)\b/u)
  return valueIndex >= 0 ? `${body.slice(0, valueIndex)}${formulaXml}${body.slice(valueIndex)}` : `${formulaXml}${body}`
}

function upsertFormulaInExistingCell(sheetXml: string, formula: WorkbookSheetDataTableFormulaSnapshot): string | null {
  const addressPattern = escapeRegExp(formula.address)
  const cellPattern = new RegExp(`<c\\b(?<attributes>[^>]*\\br=(["'])${addressPattern}\\2[^>]*)>(?<body>[\\s\\S]*?)<\\/c>`, 'u')
  if (!cellPattern.test(sheetXml)) {
    return null
  }
  return sheetXml.replace(cellPattern, (_cellXml: string, attributes: string, _quote: string, body: string) => {
    return `<c${attributes}>${insertFormulaIntoCellBody(body, formula.formulaXml)}</c>`
  })
}

function insertFormulaCellIntoRow(sheetXml: string, formula: WorkbookSheetDataTableFormulaSnapshot): string | null {
  const rowNumber = cellRowNumber(formula.address)
  if (rowNumber === null) {
    return null
  }
  const rowPattern = new RegExp(`<row\\b(?<attributes>[^>]*\\br=(["'])${String(rowNumber)}\\2[^>]*)>(?<body>[\\s\\S]*?)<\\/row>`, 'u')
  if (!rowPattern.test(sheetXml)) {
    return null
  }
  const nextCellXml = `<c r="${escapeXml(formula.address)}">${formula.formulaXml}</c>`
  return sheetXml.replace(rowPattern, (_rowXml: string, attributes: string, _quote: string, body: string) => {
    const targetColumn = cellColumnIndex(formula.address)
    let insertIndex = body.length
    for (const match of body.matchAll(/<c\b(?<attributes>[^>]*)>(?:[\s\S]*?)<\/c>/gu)) {
      const address = readAttribute(match.groups?.['attributes'] ?? '', 'r')
      if (address && cellColumnIndex(address) > targetColumn) {
        insertIndex = match.index
        break
      }
    }
    return `<row${attributes}>${body.slice(0, insertIndex)}${nextCellXml}${body.slice(insertIndex)}</row>`
  })
}

function insertFormulaCellIntoSheetData(sheetXml: string, formula: WorkbookSheetDataTableFormulaSnapshot): string {
  const rowNumber = cellRowNumber(formula.address) ?? 1
  const rowXml = `<row r="${String(rowNumber)}"><c r="${escapeXml(formula.address)}">${formula.formulaXml}</c></row>`
  return sheetXml.includes('</sheetData>') ? sheetXml.replace('</sheetData>', `${rowXml}</sheetData>`) : sheetXml
}

function upsertDataTableFormula(sheetXml: string, formula: WorkbookSheetDataTableFormulaSnapshot): string {
  return (
    upsertFormulaInExistingCell(sheetXml, formula) ??
    insertFormulaCellIntoRow(sheetXml, formula) ??
    insertFormulaCellIntoSheetData(sheetXml, formula)
  )
}

function dataTableFormulaOutputAddresses(formula: WorkbookSheetDataTableFormulaSnapshot): string[] {
  const range = decodeFormulaRefRange(formula)
  if (!range) {
    return []
  }
  const addresses: string[] = []
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      addresses.push(XLSX.utils.encode_cell({ r: row, c: column }))
    }
  }
  return addresses
}

function removeFormulaElementsFromCells(sheetXml: string, addresses: ReadonlySet<string>): string {
  if (addresses.size === 0) {
    return sheetXml
  }
  cellElementPattern.lastIndex = 0
  return sheetXml.replace(cellElementPattern, (cellXml: string, attributes: string, body: string) => {
    const address = readAttribute(attributes, 'r')
    if (!address || !addresses.has(address)) {
      return cellXml
    }
    return `<c${attributes}>${body.replace(formulaElementGlobalPattern, '')}</c>`
  })
}

export function readImportedWorkbookDataTableFormulas(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetDataTableFormulasSnapshot> {
  const zip = readXlsxZipEntries(source)
  const formulasBySheet = new Map<string, WorkbookSheetDataTableFormulasSnapshot>()
  sheetNames.forEach((sheetName, sheetIndex) => {
    const formulas = readDataTableFormulaSnapshots(getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`))
    if (formulas.length > 0) {
      formulasBySheet.set(sheetName, { formulas })
    }
  })
  return formulasBySheet
}

export function addExportDataTableFormulasToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const sheetsWithDataTableFormulas = snapshot.sheets.filter((sheet) => (sheet.metadata?.dataTableFormulas?.formulas.length ?? 0) > 0)
  if (sheetsWithDataTableFormulas.length === 0) {
    return bytes
  }

  const zip: XlsxZipEntries = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const dataTableFormulas = sheet.metadata?.dataTableFormulas?.formulas
      if (!dataTableFormulas || dataTableFormulas.length === 0) {
        return
      }
      const sheetPath = normalizeZipPath(`xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const dataTableOutputAddresses = new Set(dataTableFormulas.flatMap(dataTableFormulaOutputAddresses))
      const strippedSheetXml = removeFormulaElementsFromCells(sheetXml, dataTableOutputAddresses)
      const nextSheetXml = dataTableFormulas.reduce(upsertDataTableFormula, strippedSheetXml)
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, sheetPath, nextSheetXml)
        changed = true
      }
    })

  return changed ? zipSync(zip) : bytes
}
