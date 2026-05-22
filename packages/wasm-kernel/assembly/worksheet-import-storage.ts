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
let worksheetImportScratch = new Uint8Array(64)

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

export function prepareWorksheetImportScratch(byteLength: i32): usize {
  const capacity = positiveCapacity(byteLength)
  if (capacity > worksheetImportScratch.length) {
    worksheetImportScratch = new Uint8Array(capacity)
  }
  return worksheetImportScratch.dataStart
}

export function addWorksheetImportNumberCellFromScratch(row: u32, column: u32, byteLength: i32): i32 {
  const value = parseWorksheetImportNumberScratch(byteLength)
  if (!isFiniteNumber(value)) {
    return -1
  }
  return addWorksheetImportNumberCell(row, column, value)
}

export function readWorksheetImportNonNegativeIntegerFromScratch(byteLength: i32): f64 {
  return parseWorksheetImportNonNegativeIntegerScratch(byteLength)
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

function parseWorksheetImportNonNegativeIntegerScratch(byteLength: i32): f64 {
  let start = 0
  let end = clampedScratchLength(byteLength)
  while (start < end && isAsciiWhitespace(worksheetImportScratch[start])) {
    start += 1
  }
  while (end > start && isAsciiWhitespace(worksheetImportScratch[end - 1])) {
    end -= 1
  }
  if (start == end) {
    return -1.0
  }
  let value = 0.0
  for (let index = start; index < end; index += 1) {
    const byte = worksheetImportScratch[index]
    if (byte < 48 || byte > 57) {
      return -1.0
    }
    value = value * 10.0 + <f64>(byte - 48)
    if (value > 4294967295.0) {
      return -1.0
    }
  }
  return value
}

function parseWorksheetImportNumberScratch(byteLength: i32): f64 {
  let start = 0
  let end = clampedScratchLength(byteLength)
  while (start < end && isAsciiWhitespace(worksheetImportScratch[start])) {
    start += 1
  }
  while (end > start && isAsciiWhitespace(worksheetImportScratch[end - 1])) {
    end -= 1
  }
  if (start == end) {
    return NaN
  }
  let index = start
  let sign = 1.0
  const first = worksheetImportScratch[index]
  if (first == 43 || first == 45) {
    sign = first == 45 ? -1.0 : 1.0
    index += 1
  }
  if (index == end) {
    return NaN
  }
  let significand = 0.0
  let decimalScale = 1.0
  let digitCount = 0
  let sawDigit = false
  let sawDecimalPoint = false
  while (index < end) {
    const byte = worksheetImportScratch[index]
    if (byte >= 48 && byte <= 57) {
      if (digitCount >= 15) {
        return NaN
      }
      sawDigit = true
      digitCount += 1
      significand = significand * 10.0 + <f64>(byte - 48)
      if (sawDecimalPoint) {
        decimalScale *= 10.0
      }
      index += 1
      continue
    }
    if (byte == 46 && !sawDecimalPoint) {
      sawDecimalPoint = true
      index += 1
      continue
    }
    break
  }
  if (!sawDigit) {
    return NaN
  }
  let exponent = 0
  if (index < end && (worksheetImportScratch[index] == 69 || worksheetImportScratch[index] == 101)) {
    index += 1
    let exponentSign = 1
    if (index < end && (worksheetImportScratch[index] == 43 || worksheetImportScratch[index] == 45)) {
      exponentSign = worksheetImportScratch[index] == 45 ? -1 : 1
      index += 1
    }
    if (index == end) {
      return NaN
    }
    let exponentDigits = 0
    while (index < end) {
      const byte = worksheetImportScratch[index]
      if (byte < 48 || byte > 57) {
        return NaN
      }
      exponent = exponent * 10 + byte - 48
      exponentDigits += 1
      if (exponent > 308) {
        return NaN
      }
      index += 1
    }
    if (exponentDigits == 0) {
      return NaN
    }
    exponent *= exponentSign
  }
  if (index != end) {
    return NaN
  }
  const value = sign * (significand / decimalScale) * Math.pow(10.0, <f64>exponent)
  if (!sawDecimalPoint && exponent == 0 && Math.abs(value) > 9007199254740991.0) {
    return NaN
  }
  return value
}

function clampedScratchLength(byteLength: i32): i32 {
  if (byteLength <= 0) {
    return 0
  }
  return byteLength < worksheetImportScratch.length ? byteLength : worksheetImportScratch.length
}

function isAsciiWhitespace(byte: u8): boolean {
  return byte == 32 || byte == 9 || byte == 10 || byte == 13
}

function isFiniteNumber(value: f64): boolean {
  return value == value && value != Infinity && value != -Infinity
}
