import {
  decodeCellRange,
  encodeCellAddress,
  writeSimpleXlsxWorkbook,
  type SimpleXlsxCell,
  type SimpleXlsxDefinedName,
  type SimpleXlsxSheet,
  type SimpleXlsxStyle,
  type SimpleXlsxWorkbook,
} from '@bilig/xlsx'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { buildExportDefinedNames } from './xlsx-defined-names.js'
import { encodeFormulaForXlsx } from './xlsx-formula-translation.js'

const maxGeneratedRangeCells = 200_000

const supportedWorkbookMetadataKeys = new Set(['definedNames', 'styles', 'formats'])
const supportedSheetMetadataKeys = new Set(['rows', 'columns', 'styleRanges', 'formatRanges', 'merges', 'filters'])

function hasUnsupportedMetadata(metadata: object | undefined, supportedKeys: ReadonlySet<string>): boolean {
  if (!metadata) {
    return false
  }
  return Object.entries(metadata).some(([key, value]) => {
    if (supportedKeys.has(key)) {
      return false
    }
    if (value === undefined || value === null) {
      return false
    }
    if (Array.isArray(value)) {
      return value.length > 0
    }
    if (typeof value === 'object') {
      return Object.keys(value).length > 0
    }
    return true
  })
}

function simpleRangeCellCount(startAddress: string, endAddress: string): number | null {
  try {
    const range = decodeCellRange(`${startAddress}:${endAddress}`)
    return (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1)
  } catch {
    return null
  }
}

function canMaterializeRanges(sheet: WorkbookSnapshot['sheets'][number]): boolean {
  let generatedCells = 0
  for (const range of [...(sheet.metadata?.styleRanges ?? []), ...(sheet.metadata?.formatRanges ?? [])]) {
    if (range.range.sheetName !== sheet.name) {
      return false
    }
    const count = simpleRangeCellCount(range.range.startAddress, range.range.endAddress)
    if (count === null) {
      return false
    }
    generatedCells += count
    if (generatedCells > maxGeneratedRangeCells) {
      return false
    }
  }
  return true
}

function hasOnlySimpleFilters(sheet: WorkbookSnapshot['sheets'][number]): boolean {
  return (sheet.metadata?.filters ?? []).every((filter) => !filter.criteria || filter.criteria.length === 0)
}

function canUseBiligSimpleWriter(snapshot: WorkbookSnapshot): boolean {
  if (snapshot.sheets.length === 0) {
    return false
  }
  if (hasUnsupportedMetadata(snapshot.workbook.metadata, supportedWorkbookMetadataKeys)) {
    return false
  }
  return snapshot.sheets.every(
    (sheet) =>
      !hasUnsupportedMetadata(sheet.metadata, supportedSheetMetadataKeys) && hasOnlySimpleFilters(sheet) && canMaterializeRanges(sheet),
  )
}

const invalidExportSheetNameCharacters = ['[', ']', ':', '*', '?', '/', '\\'] as const

function normalizeExportSheetName(name: string, order: number, usedNames: Set<string>): string {
  let sanitized = name
  for (const character of invalidExportSheetNameCharacters) {
    sanitized = sanitized.split(character).join(' ')
  }
  const baseName = sanitized.length > 0 ? sanitized : `Sheet${order + 1}`
  let candidate = baseName.slice(0, 31)
  candidate = candidate.length > 0 ? candidate : `Sheet${order + 1}`
  let suffix = 1
  while (usedNames.has(candidate)) {
    const suffixText = ` ${String(suffix)}`
    candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`
    suffix += 1
  }
  usedNames.add(candidate)
  return candidate
}

function rangeAddresses(startAddress: string, endAddress: string): string[] {
  const range = decodeCellRange(`${startAddress}:${endAddress}`)
  const output: string[] = []
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      output.push(encodeCellAddress({ r: row, c: col }))
    }
  }
  return output
}

function readCellCoordinates(address: string): { readonly row: number; readonly col: number } {
  const decoded = decodeCellRange(address)
  return { row: decoded.s.r, col: decoded.s.c }
}

function ensureCell(cellsByAddress: Map<string, SimpleXlsxCell>, address: string): SimpleXlsxCell {
  const existing = cellsByAddress.get(address)
  if (existing) {
    return existing
  }
  const coordinates = readCellCoordinates(address)
  const cell: SimpleXlsxCell = { address, row: coordinates.row, col: coordinates.col }
  cellsByAddress.set(address, cell)
  return cell
}

function replaceCell(cellsByAddress: Map<string, SimpleXlsxCell>, cell: SimpleXlsxCell, patch: Partial<SimpleXlsxCell>): void {
  cellsByAddress.set(cell.address, {
    ...cell,
    ...patch,
  })
}

function simpleSheetCells(
  sheet: WorkbookSnapshot['sheets'][number],
  formatCodesById: ReadonlyMap<string, string>,
): readonly SimpleXlsxCell[] {
  const cellsByAddress = new Map<string, SimpleXlsxCell>()
  for (const cell of sheet.cells) {
    const decoded = readCellCoordinates(cell.address)
    cellsByAddress.set(cell.address, {
      address: cell.address,
      row: decoded.row,
      col: decoded.col,
      ...(cell.value !== undefined ? { value: cell.value } : {}),
      ...(cell.formula ? { formula: encodeFormulaForXlsx(cell.formula.replace(/^=/u, '')) } : {}),
      ...(cell.format?.trim() ? { numberFormat: cell.format.trim() } : {}),
    })
  }
  for (const styleRange of sheet.metadata?.styleRanges ?? []) {
    for (const address of rangeAddresses(styleRange.range.startAddress, styleRange.range.endAddress)) {
      const cell = ensureCell(cellsByAddress, address)
      replaceCell(cellsByAddress, cell, { styleId: styleRange.styleId })
    }
  }
  for (const formatRange of sheet.metadata?.formatRanges ?? []) {
    const format = formatCodesById.get(formatRange.formatId)?.trim()
    if (!format || format === 'General') {
      continue
    }
    for (const address of rangeAddresses(formatRange.range.startAddress, formatRange.range.endAddress)) {
      const cell = ensureCell(cellsByAddress, address)
      replaceCell(cellsByAddress, cell, { numberFormat: format })
    }
  }
  return [...cellsByAddress.values()].toSorted((left, right) => left.row - right.row || left.col - right.col)
}

function simpleStyles(snapshot: WorkbookSnapshot): readonly SimpleXlsxStyle[] | undefined {
  const styles = snapshot.workbook.metadata?.styles
  return styles && styles.length > 0 ? styles : undefined
}

function simpleDefinedNames(
  snapshot: WorkbookSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
  exportSheetIndexesByOriginalName: ReadonlyMap<string, number>,
): readonly SimpleXlsxDefinedName[] | undefined {
  const definedNames = buildExportDefinedNames(
    snapshot.workbook.metadata?.definedNames,
    exportSheetNamesByOriginalName,
    exportSheetIndexesByOriginalName,
  )
  if (!definedNames) {
    return undefined
  }
  const output: SimpleXlsxDefinedName[] = []
  for (const definedName of definedNames) {
    if (!definedName.Name || !definedName.Ref) {
      continue
    }
    output.push({
      name: definedName.Name,
      formula: definedName.Ref,
      ...(definedName.Sheet !== undefined ? { localSheetIndex: definedName.Sheet } : {}),
    })
  }
  return output.length > 0 ? output : undefined
}

function simpleSheet(
  sheet: WorkbookSnapshot['sheets'][number],
  exportName: string,
  formatCodesById: ReadonlyMap<string, string>,
): SimpleXlsxSheet {
  const rows = sheet.metadata?.rows?.map((row) => ({
    index: row.index,
    ...(typeof row.size === 'number' ? { size: row.size } : {}),
    ...(row.hidden === true || row.filterHidden === true ? { hidden: true } : {}),
  }))
  const columns = sheet.metadata?.columns?.map((column) => ({
    index: column.index,
    ...(typeof column.size === 'number' ? { size: column.size } : {}),
    ...(column.hidden === true ? { hidden: true } : {}),
  }))
  const merges = sheet.metadata?.merges?.map((merge) => ({
    startAddress: merge.startAddress,
    endAddress: merge.endAddress,
  }))
  const autoFilters = sheet.metadata?.filters?.map((filter) => ({
    startAddress: filter.startAddress,
    endAddress: filter.endAddress,
  }))
  return {
    name: exportName,
    cells: simpleSheetCells(sheet, formatCodesById),
    ...(rows && rows.length > 0 ? { rows } : {}),
    ...(columns && columns.length > 0 ? { columns } : {}),
    ...(merges && merges.length > 0 ? { merges } : {}),
    ...(autoFilters && autoFilters.length > 0 ? { autoFilters } : {}),
  }
}

function simpleWorkbook(snapshot: WorkbookSnapshot): SimpleXlsxWorkbook {
  const usedNames = new Set<string>()
  const exportSheetNamesByOriginalName = new Map<string, string>()
  const exportSheetIndexesByOriginalName = new Map<string, number>()
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  for (const sheet of orderedSheets) {
    const exportSheetName = normalizeExportSheetName(sheet.name, sheet.order, usedNames)
    exportSheetNamesByOriginalName.set(sheet.name, exportSheetName)
    exportSheetIndexesByOriginalName.set(sheet.name, exportSheetIndexesByOriginalName.size)
  }
  const formatCodesById = new Map((snapshot.workbook.metadata?.formats ?? []).map((format) => [format.id, format.code]))
  const styles = simpleStyles(snapshot)
  const definedNames = simpleDefinedNames(snapshot, exportSheetNamesByOriginalName, exportSheetIndexesByOriginalName)
  return {
    sheets: orderedSheets.map((sheet) => simpleSheet(sheet, exportSheetNamesByOriginalName.get(sheet.name) ?? sheet.name, formatCodesById)),
    ...(styles ? { styles } : {}),
    ...(definedNames ? { definedNames } : {}),
  }
}

export function tryExportBiligSimpleXlsx(snapshot: WorkbookSnapshot): Uint8Array | null {
  return canUseBiligSimpleWriter(snapshot) ? writeSimpleXlsxWorkbook(simpleWorkbook(snapshot)) : null
}
