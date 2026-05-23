import { ValueTag, parseCellNumberFormatCode, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EngineRuntimeState } from '../runtime-state.js'

type RuntimeWorkbook = EngineRuntimeState['workbook']

export function readPrecisionAsDisplayedCellValue(
  state: Pick<EngineRuntimeState, 'workbook' | 'strings'>,
  cellIndex: number | undefined,
): CellValue {
  if (cellIndex === undefined) {
    return { tag: ValueTag.Empty }
  }
  const value = state.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : state.strings.get(stringId)))
  if (state.workbook.getCalculationSettings().fullPrecision !== false || !cellHasFormula(state.workbook, cellIndex)) {
    return value
  }
  return roundValueForPrecisionAsDisplayed(value, precisionFormatCodeForCell(state.workbook, cellIndex))
}

export function roundFormulaResultForPrecisionAsDisplayed(
  state: Pick<EngineRuntimeState, 'workbook'>,
  cellIndex: number,
  value: CellValue,
): CellValue {
  if (state.workbook.getCalculationSettings().fullPrecision !== false) {
    return value
  }
  return roundValueForPrecisionAsDisplayed(value, precisionFormatCodeForCell(state.workbook, cellIndex))
}

export function roundValueForPrecisionAsDisplayed(value: CellValue, formatCode: string | undefined): CellValue {
  if (value.tag !== ValueTag.Number || !Number.isFinite(value.value)) {
    return value
  }
  const places = displayedPrecisionDecimalPlaces(formatCode)
  if (places === null) {
    return value
  }
  return { tag: ValueTag.Number, value: roundHalfAwayFromZero(value.value, places) }
}

export function displayedPrecisionDecimalPlaces(formatCode: string | undefined): number | null {
  const normalized = (formatCode ?? 'general').trim()
  if (normalized.length === 0 || normalized.toLowerCase() === 'general') {
    return null
  }

  const preset = parseCellNumberFormatCode(normalized)
  switch (preset.kind) {
    case 'number':
    case 'currency':
    case 'accounting':
      return preset.decimals ?? 2
    case 'percent':
      return (preset.decimals ?? 2) + 2
    case 'general':
    case 'text':
    case 'date':
    case 'time':
    case 'datetime':
      break
  }

  return displayedPrecisionDecimalPlacesFromExcelFormat(normalized)
}

function precisionFormatCodeForCell(workbook: RuntimeWorkbook, cellIndex: number): string | undefined {
  const explicitFormat = workbook.getCellFormat(cellIndex)
  if (explicitFormat !== undefined) {
    return explicitFormat
  }
  const sheetId = workbook.cellStore.sheetIds[cellIndex]
  const row = workbook.cellStore.rows[cellIndex]
  const col = workbook.cellStore.cols[cellIndex]
  if (sheetId === undefined || row === undefined || col === undefined) {
    return undefined
  }
  const sheetName = workbook.getSheetNameById(sheetId)
  if (!sheetName) {
    return undefined
  }
  const formatId = workbook.getRangeFormatId(sheetName, row, col)
  return workbook.getCellNumberFormat(formatId)?.code
}

function cellHasFormula(workbook: RuntimeWorkbook, cellIndex: number): boolean {
  return ((workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.HasFormula) !== 0
}

function displayedPrecisionDecimalPlacesFromExcelFormat(formatCode: string): number | null {
  const section = firstFormatSection(formatCode)
  const stripped = stripExcelFormatLiterals(section)
  if (
    stripped.length === 0 ||
    /(?:^|[^A-Za-z])general(?:$|[^A-Za-z])/iu.test(stripped) ||
    /[@/?]/u.test(stripped) ||
    /[ymdhsa]/iu.test(stripped) ||
    /e[+-]/iu.test(stripped)
  ) {
    return null
  }

  const percentPlaces = (stripped.match(/%/gu)?.length ?? 0) * 2
  const numeric = stripped.replace(/%/gu, '')
  const decimalIndex = numeric.indexOf('.')
  const integerPattern = decimalIndex >= 0 ? numeric.slice(0, decimalIndex) : numeric
  const fractionPattern = decimalIndex >= 0 ? numeric.slice(decimalIndex + 1) : ''
  if (!/[0#?]/u.test(integerPattern) && !/[0#?]/u.test(fractionPattern)) {
    return null
  }
  return Math.min(15, countDecimalPlaceholders(fractionPattern) + percentPlaces)
}

function firstFormatSection(formatCode: string): string {
  let quoted = false
  for (let index = 0; index < formatCode.length; index += 1) {
    const char = formatCode[index]
    if (char === '"') {
      quoted = !quoted
      continue
    }
    if (!quoted && char === ';') {
      return formatCode.slice(0, index)
    }
    if (char === '\\' || char === '_' || char === '*') {
      index += 1
    }
  }
  return formatCode
}

function stripExcelFormatLiterals(formatCode: string): string {
  let output = ''
  let quoted = false
  let bracketDepth = 0
  for (let index = 0; index < formatCode.length; index += 1) {
    const char = formatCode[index]
    if (quoted) {
      if (char === '"') {
        quoted = false
      }
      continue
    }
    if (bracketDepth > 0) {
      if (char === ']') {
        bracketDepth -= 1
      }
      continue
    }
    if (char === '"') {
      quoted = true
      continue
    }
    if (char === '[') {
      bracketDepth += 1
      continue
    }
    if (char === '\\' || char === '_' || char === '*') {
      index += 1
      continue
    }
    output += char
  }
  return output
}

function countDecimalPlaceholders(pattern: string): number {
  let count = 0
  for (const char of pattern) {
    if (char === '0' || char === '#' || char === '?') {
      count += 1
    }
  }
  return count
}

function roundHalfAwayFromZero(value: number, decimalPlaces: number): number {
  const factor = 10 ** decimalPlaces
  if (!Number.isFinite(factor) || factor <= 0) {
    return value
  }
  const rounded = (Math.sign(value) * Math.round(Math.abs(value) * factor + Number.EPSILON)) / factor
  return Object.is(rounded, -0) ? 0 : rounded
}
