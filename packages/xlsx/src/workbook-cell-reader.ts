import { decodeCellAddress, decodeCellRange, encodeCellAddress, normalizeCellAddress, type XlsxCellRange } from './address.js'
import { workbookSheetPathEntriesForSource } from './workbook-sheet-paths.js'
import { decodeXmlText, getXmlElementText, readXmlAttribute, worksheetCellElementPattern, worksheetCellOpeningTagPattern } from './xml.js'
import { getZipText, readXlsxZipEntries, readXlsxZipEntriesLazy, type XlsxZipSource } from './zip-reader.js'

export type XlsxWorkbookCellValue = string | number | boolean

export interface XlsxWorkbookCell {
  readonly address: string
  readonly row: number
  readonly col: number
  readonly formula: string | null
  readonly hasValue: boolean
  readonly rawValue: string | null
  readonly styleIndex: number | null
  readonly type: string | null
  readonly value: XlsxWorkbookCellValue | null
}

export interface XlsxWorkbookSheetCells {
  readonly name: string
  readonly index: number
  readonly path: string
  readonly cells: readonly XlsxWorkbookCell[]
  readonly range: XlsxCellRange
  readonly rowCount: number
  readonly columnCount: number
}

export interface XlsxWorkbookCells {
  readonly sheets: readonly XlsxWorkbookSheetCells[]
  readonly hasFormulaMarkup: boolean
  readonly hasUnresolvedFormulaMarkup: boolean
}

const worksheetRowElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?row\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?row)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const worksheetRowOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?row\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u

export function readXlsxWorkbookCells(source: XlsxZipSource): XlsxWorkbookCells {
  const zip = source instanceof Uint8Array ? readXlsxZipEntriesLazy(source) : readXlsxZipEntries(source)
  let workbookHasFormulaMarkup = false
  let workbookHasUnresolvedFormulaMarkup = false
  const sharedStringAt = lazySharedStringReader(zip)
  const sheets = workbookSheetPathEntriesForSource(zip).map((sheet) => {
    const sheetXml = getZipText(zip, sheet.path) ?? ''
    const cells = readWorksheetCells(sheetXml, sharedStringAt)
    const formulaElementCount = sheetXml.match(/<f\b/gu)?.length ?? 0
    const cellFormulaCount = cells.filter((cell) => cell.formula !== null).length
    if (cellFormulaCount > 0 || formulaElementCount > 0) {
      workbookHasFormulaMarkup = true
    }
    if (formulaElementCount > cellFormulaCount || cells.some((cell) => cell.formula !== null && cell.formula.trim().length === 0)) {
      workbookHasUnresolvedFormulaMarkup = true
    }
    const observedRange = rangeFromCells(cells)
    const range = readWorksheetDimension(sheetXml) ?? observedRange
    const rowCount = Math.max(observedRange.e.r + 1, 1)
    const columnCount = Math.max(observedRange.e.c + 1, 1)
    return {
      name: sheet.name,
      index: sheet.index,
      path: sheet.path,
      cells,
      range,
      rowCount,
      columnCount,
    }
  })
  return {
    sheets,
    hasFormulaMarkup: workbookHasFormulaMarkup,
    hasUnresolvedFormulaMarkup: workbookHasUnresolvedFormulaMarkup,
  }
}

function readWorksheetCells(sheetXml: string, sharedStringAt: (index: number) => string | null): XlsxWorkbookCell[] {
  const cells: XlsxWorkbookCell[] = []
  let sawRows = false
  let fallbackRow = 0
  for (const rowMatch of sheetXml.matchAll(worksheetRowElementPattern)) {
    sawRows = true
    const rowXml = rowMatch[0]
    const rowOpeningTag = worksheetRowOpeningTagPattern.exec(rowXml)?.[0] ?? ''
    const explicitRow = decodePositiveRowIndex(readXmlAttribute(rowOpeningTag, 'r'))
    const rowIndex = explicitRow ?? fallbackRow
    let fallbackCol = 0
    for (const cellMatch of rowXml.matchAll(worksheetCellElementPattern)) {
      const cell = readWorkbookCellXml(cellMatch[0], rowIndex, fallbackCol, sharedStringAt)
      if (cell) {
        cells.push(cell)
        fallbackCol = cell.col + 1
      } else {
        fallbackCol += 1
      }
    }
    fallbackRow = rowIndex + 1
  }
  if (sawRows) {
    return cells
  }
  let fallbackCol = 0
  for (const cellMatch of sheetXml.matchAll(worksheetCellElementPattern)) {
    const cell = readWorkbookCellXml(cellMatch[0], 0, fallbackCol, sharedStringAt)
    if (cell) {
      cells.push(cell)
      fallbackCol = cell.col + 1
    } else {
      fallbackCol += 1
    }
  }
  return cells
}

function readWorkbookCellXml(
  cellXml: string,
  fallbackRow: number,
  fallbackCol: number,
  sharedStringAt: (index: number) => string | null,
): XlsxWorkbookCell | null {
  const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0] ?? ''
  const explicitAddress = readXmlAttribute(openingTag, 'r')
  const decodedAddress = explicitAddress ? decodeCellAddressOrNull(explicitAddress) : { r: fallbackRow, c: fallbackCol }
  if (!decodedAddress) {
    return null
  }
  const address = explicitAddress ? normalizeCellAddress(explicitAddress) : encodeCellAddress(decodedAddress)
  const type = readXmlAttribute(openingTag, 't')
  const rawValueText = getXmlElementText(cellXml, 'v')
  const rawValue = rawValueText === null ? null : decodeXmlText(rawValueText)
  const formulaText = getXmlElementText(cellXml, 'f')
  const formula = formulaText === null ? null : decodeXmlText(formulaText)
  const value = readCellValue(cellXml, type, rawValue, sharedStringAt)
  return {
    address,
    row: decodedAddress.r,
    col: decodedAddress.c,
    formula,
    hasValue: rawValue !== null || type === 'inlineStr',
    rawValue,
    styleIndex: readCellStyleIndex(openingTag),
    type,
    value,
  }
}

function readCellValue(
  cellXml: string,
  type: string | null,
  rawValue: string | null,
  sharedStringAt: (index: number) => string | null,
): XlsxWorkbookCellValue | null {
  if (type === 'inlineStr') {
    return readRichTextValue(cellXml)
  }
  if (rawValue === null) {
    return null
  }
  if (type === 's') {
    const sharedStringIndex = Number(rawValue)
    return Number.isSafeInteger(sharedStringIndex) && sharedStringIndex >= 0 ? sharedStringAt(sharedStringIndex) : null
  }
  if (type === 'b') {
    return rawValue === '1' || rawValue.toLowerCase() === 'true'
  }
  if (type === 'str' || type === 'e') {
    return rawValue
  }
  if (rawValue.length === 0) {
    return null
  }
  const numericValue = Number(rawValue)
  return Number.isFinite(numericValue) ? numericValue : rawValue
}

function lazySharedStringReader(zip: XlsxZipSource): (index: number) => string | null {
  let sharedStrings: readonly string[] | undefined
  return (index) => {
    sharedStrings ??= readSharedStrings(zip)
    return sharedStrings[index] ?? null
  }
}

function readSharedStrings(zip: XlsxZipSource): readonly string[] {
  const xml = getZipText(readXlsxZipEntries(zip), 'xl/sharedStrings.xml')
  if (!xml) {
    return []
  }
  return [...xml.matchAll(/<si\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/si>/gu)].map((match) => readRichTextValue(match[0]))
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

function readWorksheetDimension(sheetXml: string): XlsxCellRange | null {
  const tag = /<(?:[A-Za-z_][\w.-]*:)?dimension\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/u.exec(sheetXml)?.[0] ?? ''
  const ref = readXmlAttribute(tag, 'ref')
  if (!ref) {
    return null
  }
  try {
    return decodeCellRange(ref.replaceAll('$', ''))
  } catch {
    return null
  }
}

function rangeFromCells(cells: readonly XlsxWorkbookCell[]): XlsxCellRange {
  if (cells.length === 0) {
    return { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
  }
  return cells.reduce(
    (range, cell) => ({
      s: {
        r: Math.min(range.s.r, cell.row),
        c: Math.min(range.s.c, cell.col),
      },
      e: {
        r: Math.max(range.e.r, cell.row),
        c: Math.max(range.e.c, cell.col),
      },
    }),
    {
      s: { r: cells[0]!.row, c: cells[0]!.col },
      e: { r: cells[0]!.row, c: cells[0]!.col },
    },
  )
}

function decodePositiveRowIndex(value: string | null): number | null {
  if (!value) {
    return null
  }
  const row = Number.parseInt(value, 10)
  return Number.isSafeInteger(row) && row > 0 ? row - 1 : null
}

function decodeCellAddressOrNull(address: string): { readonly r: number; readonly c: number } | null {
  try {
    return decodeCellAddress(address.replaceAll('$', ''))
  } catch {
    return null
  }
}
