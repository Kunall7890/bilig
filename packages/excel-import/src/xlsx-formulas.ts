import * as XLSX from 'xlsx'

import { translateFormulaReferences } from '@bilig/formula'
import { getZipText, type XlsxZipEntries } from './xlsx-zip.js'
import { workbookSheetPath } from './xlsx-workbook-sheet-paths.js'

interface SharedFormulaBase {
  readonly row: number
  readonly column: number
  readonly formula: string
}

interface WorksheetFormulaCell {
  readonly address: string
  readonly row: number
  readonly column: number
  readonly formulaType: string | null
  readonly sharedIndex: string | null
  readonly formula: string
}

const cellElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?c)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const cellOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u
const formulaElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?f)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/u
const formulaOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u
const formulaTextPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/u
const cellAddressPattern = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$/u

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function normalizeAddress(address: string): string | null {
  if (!cellAddressPattern.test(address)) {
    return null
  }
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(address.replaceAll('$', '')))
  } catch {
    return null
  }
}

function readFormulaXml(cellXml: string): string | null {
  return formulaElementPattern.exec(cellXml)?.[0] ?? null
}

function readFormulaText(formulaXml: string): string {
  return decodeXmlText(formulaTextPattern.exec(formulaXml)?.[1] ?? '')
}

function readWorksheetFormulaCells(sheetXml: string | null): WorksheetFormulaCell[] {
  if (!sheetXml) {
    return []
  }
  const cells: WorksheetFormulaCell[] = []
  for (const match of sheetXml.matchAll(cellElementPattern)) {
    const cellXml = match[0]
    const cellOpeningTag = cellOpeningTagPattern.exec(cellXml)?.[0]
    const rawAddress = cellOpeningTag ? readXmlAttribute(cellOpeningTag, 'r') : null
    const address = rawAddress ? normalizeAddress(rawAddress) : null
    if (!address) {
      continue
    }
    const formulaXml = readFormulaXml(cellXml)
    const formulaOpeningTag = formulaXml ? formulaOpeningTagPattern.exec(formulaXml)?.[0] : null
    if (!formulaXml || !formulaOpeningTag) {
      continue
    }
    const decodedAddress = XLSX.utils.decode_cell(address)
    cells.push({
      address,
      row: decodedAddress.r,
      column: decodedAddress.c,
      formulaType: readXmlAttribute(formulaOpeningTag, 't'),
      sharedIndex: readXmlAttribute(formulaOpeningTag, 'si'),
      formula: readFormulaText(formulaXml),
    })
  }
  return cells
}

function readWorksheetSharedFormulas(sheetXml: string | null): Map<string, string> {
  const cells = readWorksheetFormulaCells(sheetXml)
  const sharedBases = new Map<string, SharedFormulaBase>()
  const formulas = new Map<string, string>()

  for (const cell of cells) {
    if (cell.formulaType !== 'shared' || cell.sharedIndex === null || cell.formula.trim().length === 0) {
      continue
    }
    sharedBases.set(cell.sharedIndex, {
      row: cell.row,
      column: cell.column,
      formula: cell.formula,
    })
    formulas.set(cell.address, cell.formula)
  }

  for (const cell of cells) {
    if (cell.formulaType !== 'shared' || cell.sharedIndex === null || cell.formula.trim().length > 0) {
      continue
    }
    const base = sharedBases.get(cell.sharedIndex)
    if (!base) {
      continue
    }
    try {
      formulas.set(cell.address, translateFormulaReferences(base.formula, cell.row - base.row, cell.column - base.column))
    } catch {
      // SheetJS has already provided a best-effort formula for unsupported syntax.
    }
  }

  return formulas
}

export function readImportedWorksheetFormulas(
  zip: XlsxZipEntries,
  sheetNames: readonly string[],
  sheetPathsByName: ReadonlyMap<string, string>,
  fallbackSheetPaths: readonly string[],
): Map<string, Map<string, string>> {
  const formulasBySheet = new Map<string, Map<string, string>>()
  sheetNames.forEach((sheetName, index) => {
    const sheetPath = workbookSheetPath(sheetPathsByName, fallbackSheetPaths, sheetName, index)
    if (!sheetPath) {
      return
    }
    const formulas = readWorksheetSharedFormulas(getZipText(zip, sheetPath))
    if (formulas.size > 0) {
      formulasBySheet.set(sheetName, formulas)
    }
  })
  return formulasBySheet
}
