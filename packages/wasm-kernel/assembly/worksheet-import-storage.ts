import { ensureF64, ensureU16, ensureU32, ensureU8 } from './vm-core-helpers'

let worksheetImportRows = new Uint32Array(64)
let worksheetImportColumns = new Uint16Array(64)
let worksheetImportValueKinds = new Uint8Array(64)
let worksheetImportNumbers = new Float64Array(64)
let worksheetImportSharedStringIds = new Uint32Array(64)
let worksheetImportCellCount = 0

let worksheetImportStyleRows = new Uint32Array(16)
let worksheetImportStyleColumns = new Uint16Array(16)
let worksheetImportStyleIds = new Uint32Array(16)
let worksheetImportStyleCount = 0

let worksheetImportFormulaCellIndexes = new Uint32Array(16)
let worksheetImportFormulaRows = new Uint32Array(16)
let worksheetImportFormulaColumns = new Uint16Array(16)
let worksheetImportFormulaTypeCodes = new Uint8Array(16)
let worksheetImportFormulaSharedIndexes = new Uint32Array(16)
let worksheetImportFormulaCount = 0

export const worksheetImportValueKindFormulaOnly: u8 = 0
export const worksheetImportValueKindNumber: u8 = 1
export const worksheetImportValueKindSharedString: u8 = 2
export const worksheetImportNoSharedFormulaIndex: u32 = 0xffffffff

export function resetWorksheetImportStorage(cellCapacity: i32, styleCapacity: i32, formulaCapacity: i32): void {
  const cells = positiveCapacity(cellCapacity)
  const styles = positiveCapacity(styleCapacity)
  const formulas = positiveCapacity(formulaCapacity)
  worksheetImportRows = new Uint32Array(cells)
  worksheetImportColumns = new Uint16Array(cells)
  worksheetImportValueKinds = new Uint8Array(cells)
  worksheetImportNumbers = new Float64Array(cells)
  worksheetImportSharedStringIds = new Uint32Array(cells)
  worksheetImportCellCount = 0

  worksheetImportStyleRows = new Uint32Array(styles)
  worksheetImportStyleColumns = new Uint16Array(styles)
  worksheetImportStyleIds = new Uint32Array(styles)
  worksheetImportStyleCount = 0

  worksheetImportFormulaCellIndexes = new Uint32Array(formulas)
  worksheetImportFormulaRows = new Uint32Array(formulas)
  worksheetImportFormulaColumns = new Uint16Array(formulas)
  worksheetImportFormulaTypeCodes = new Uint8Array(formulas)
  worksheetImportFormulaSharedIndexes = new Uint32Array(formulas)
  worksheetImportFormulaCount = 0
}

export function releaseWorksheetImportStorage(): void {
  resetWorksheetImportStorage(1, 1, 1)
}

export function addWorksheetImportNumberCell(row: u32, column: u32, value: f64): i32 {
  const index = worksheetImportCellCount
  ensureWorksheetImportCellCapacity(index + 1)
  worksheetImportRows[index] = row
  worksheetImportColumns[index] = <u16>column
  worksheetImportValueKinds[index] = worksheetImportValueKindNumber
  worksheetImportNumbers[index] = value
  worksheetImportSharedStringIds[index] = worksheetImportNoSharedFormulaIndex
  worksheetImportCellCount += 1
  return index
}

export function addWorksheetImportFormulaOnlyCell(row: u32, column: u32): i32 {
  const index = worksheetImportCellCount
  ensureWorksheetImportCellCapacity(index + 1)
  worksheetImportRows[index] = row
  worksheetImportColumns[index] = <u16>column
  worksheetImportValueKinds[index] = worksheetImportValueKindFormulaOnly
  worksheetImportNumbers[index] = NaN
  worksheetImportSharedStringIds[index] = worksheetImportNoSharedFormulaIndex
  worksheetImportCellCount += 1
  return index
}

export function addWorksheetImportSharedStringCell(row: u32, column: u32, sharedStringIndex: u32): i32 {
  const index = worksheetImportCellCount
  ensureWorksheetImportCellCapacity(index + 1)
  worksheetImportRows[index] = row
  worksheetImportColumns[index] = <u16>column
  worksheetImportValueKinds[index] = worksheetImportValueKindSharedString
  worksheetImportNumbers[index] = NaN
  worksheetImportSharedStringIds[index] = sharedStringIndex
  worksheetImportCellCount += 1
  return index
}

export function addWorksheetImportStyle(row: u32, column: u32, styleId: u32): void {
  const index = worksheetImportStyleCount
  ensureWorksheetImportStyleCapacity(index + 1)
  worksheetImportStyleRows[index] = row
  worksheetImportStyleColumns[index] = <u16>column
  worksheetImportStyleIds[index] = styleId
  worksheetImportStyleCount += 1
}

export function addWorksheetImportFormulaRecord(cellIndex: u32, row: u32, column: u32, typeCode: u32, sharedIndex: u32): void {
  const index = worksheetImportFormulaCount
  ensureWorksheetImportFormulaCapacity(index + 1)
  worksheetImportFormulaCellIndexes[index] = cellIndex
  worksheetImportFormulaRows[index] = row
  worksheetImportFormulaColumns[index] = <u16>column
  worksheetImportFormulaTypeCodes[index] = <u8>typeCode
  worksheetImportFormulaSharedIndexes[index] = sharedIndex
  worksheetImportFormulaCount += 1
}

export function getWorksheetImportCellCount(): i32 {
  return worksheetImportCellCount
}

export function getWorksheetImportRowsPtr(): usize {
  return worksheetImportRows.dataStart
}

export function getWorksheetImportColumnsPtr(): usize {
  return worksheetImportColumns.dataStart
}

export function getWorksheetImportValueKindsPtr(): usize {
  return worksheetImportValueKinds.dataStart
}

export function getWorksheetImportNumbersPtr(): usize {
  return worksheetImportNumbers.dataStart
}

export function getWorksheetImportSharedStringIdsPtr(): usize {
  return worksheetImportSharedStringIds.dataStart
}

export function getWorksheetImportStyleCount(): i32 {
  return worksheetImportStyleCount
}

export function getWorksheetImportStyleRowsPtr(): usize {
  return worksheetImportStyleRows.dataStart
}

export function getWorksheetImportStyleColumnsPtr(): usize {
  return worksheetImportStyleColumns.dataStart
}

export function getWorksheetImportStyleIdsPtr(): usize {
  return worksheetImportStyleIds.dataStart
}

export function getWorksheetImportFormulaCount(): i32 {
  return worksheetImportFormulaCount
}

export function getWorksheetImportFormulaCellIndexesPtr(): usize {
  return worksheetImportFormulaCellIndexes.dataStart
}

export function getWorksheetImportFormulaRowsPtr(): usize {
  return worksheetImportFormulaRows.dataStart
}

export function getWorksheetImportFormulaColumnsPtr(): usize {
  return worksheetImportFormulaColumns.dataStart
}

export function getWorksheetImportFormulaTypeCodesPtr(): usize {
  return worksheetImportFormulaTypeCodes.dataStart
}

export function getWorksheetImportFormulaSharedIndexesPtr(): usize {
  return worksheetImportFormulaSharedIndexes.dataStart
}

function ensureWorksheetImportCellCapacity(nextCapacity: i32): void {
  worksheetImportRows = ensureU32(worksheetImportRows, nextCapacity)
  worksheetImportColumns = ensureU16(worksheetImportColumns, nextCapacity)
  worksheetImportValueKinds = ensureU8(worksheetImportValueKinds, nextCapacity)
  worksheetImportNumbers = ensureF64(worksheetImportNumbers, nextCapacity)
  worksheetImportSharedStringIds = ensureU32(worksheetImportSharedStringIds, nextCapacity)
}

function ensureWorksheetImportStyleCapacity(nextCapacity: i32): void {
  worksheetImportStyleRows = ensureU32(worksheetImportStyleRows, nextCapacity)
  worksheetImportStyleColumns = ensureU16(worksheetImportStyleColumns, nextCapacity)
  worksheetImportStyleIds = ensureU32(worksheetImportStyleIds, nextCapacity)
}

function ensureWorksheetImportFormulaCapacity(nextCapacity: i32): void {
  worksheetImportFormulaCellIndexes = ensureU32(worksheetImportFormulaCellIndexes, nextCapacity)
  worksheetImportFormulaRows = ensureU32(worksheetImportFormulaRows, nextCapacity)
  worksheetImportFormulaColumns = ensureU16(worksheetImportFormulaColumns, nextCapacity)
  worksheetImportFormulaTypeCodes = ensureU8(worksheetImportFormulaTypeCodes, nextCapacity)
  worksheetImportFormulaSharedIndexes = ensureU32(worksheetImportFormulaSharedIndexes, nextCapacity)
}

function positiveCapacity(value: i32): i32 {
  return value > 1 ? value : 1
}
