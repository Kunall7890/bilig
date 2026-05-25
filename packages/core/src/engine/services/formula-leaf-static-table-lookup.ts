import { parseRangeAddress, type JsPlanInstruction } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { emptyValue, errorValue } from '../../engine-value-utils.js'
import type { EngineRuntimeState, RuntimeFormula } from '../runtime-state.js'
import { normalizeExactLookupKey } from './direct-lookup-helpers.js'

type StaticExactTableLookupPlan = {
  readonly callee: 'VLOOKUP' | 'HLOOKUP'
  readonly rangeLookup: boolean
  readonly lookupInstruction: JsPlanInstruction
  readonly rangeInstruction: Extract<JsPlanInstruction, { opcode: 'push-range' }>
  readonly index: number
}

type StaticTableLookupState = Pick<EngineRuntimeState, 'workbook' | 'strings'>

interface StaticNumericApproximateVlookupCache {
  readonly col: number
  readonly columnVersion: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly sheetId: number
  readonly sheetName: string
  readonly structureVersion: number
  readonly values: Float64Array
}

const STATIC_NUMERIC_APPROXIMATE_VLOOKUP_CACHE = new WeakMap<RuntimeFormula, StaticNumericApproximateVlookupCache>()

export function tryEvaluateFormulaLeafStaticTableLookup(args: {
  readonly state: StaticTableLookupState
  readonly formula: RuntimeFormula
}): CellValue | undefined {
  const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[args.formula.cellIndex]!)
  if (!ownerSheetName) {
    return undefined
  }
  const plan = parseStaticExactTableLookupPlan(args.formula.compiled.jsPlan)
  if (!plan) {
    return undefined
  }
  const sheetName = plan.rangeInstruction.sheetName ?? ownerSheetName
  const parsedRange = parseRangeAddress(`${plan.rangeInstruction.start}:${plan.rangeInstruction.end}`, sheetName)
  if (parsedRange.kind !== 'cells') {
    return undefined
  }
  const lookupValue = readStaticLookupOperand(args.state, ownerSheetName, plan.lookupInstruction)
  if (lookupValue === undefined) {
    return undefined
  }
  if (lookupValue.tag === ValueTag.Error) {
    return lookupValue
  }

  const rowStart = parsedRange.start.row
  const rowEnd = parsedRange.end.row
  const colStart = parsedRange.start.col
  const colEnd = parsedRange.end.col
  if (plan.callee === 'VLOOKUP') {
    const resultCol = colStart + plan.index - 1
    if (plan.index < 1 || resultCol > colEnd) {
      return errorValue(ErrorCode.Value)
    }
    if (plan.rangeLookup) {
      return evaluateApproximateVlookup(args.state, args.formula, sheetName, rowStart, rowEnd, colStart, resultCol, lookupValue)
    }
    const lookupKey = exactLookupKey(args.state, lookupValue)
    if (lookupKey === undefined) {
      return errorValue(ErrorCode.NA)
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      if (exactLookupKey(args.state, readCellValueAt(args.state, sheetName, row, colStart)) === lookupKey) {
        return coerceLookupReturnValue(readCellValueAt(args.state, sheetName, row, resultCol))
      }
    }
    return errorValue(ErrorCode.NA)
  }

  const resultRow = rowStart + plan.index - 1
  if (plan.index < 1 || resultRow > rowEnd) {
    return errorValue(ErrorCode.Value)
  }
  if (plan.rangeLookup) {
    return evaluateApproximateHlookup(args.state, sheetName, colStart, colEnd, rowStart, resultRow, lookupValue)
  }
  const lookupKey = exactLookupKey(args.state, lookupValue)
  if (lookupKey === undefined) {
    return errorValue(ErrorCode.NA)
  }
  for (let col = colStart; col <= colEnd; col += 1) {
    if (exactLookupKey(args.state, readCellValueAt(args.state, sheetName, rowStart, col)) === lookupKey) {
      return coerceLookupReturnValue(readCellValueAt(args.state, sheetName, resultRow, col))
    }
  }
  return errorValue(ErrorCode.NA)
}

function parseStaticExactTableLookupPlan(plan: readonly JsPlanInstruction[]): StaticExactTableLookupPlan | undefined {
  if (plan.length !== 5 && plan.length !== 6) {
    return undefined
  }
  const lookupInstruction = plan[0]
  const rangeInstruction = plan[1]
  const indexInstruction = plan[2]
  const rangeLookupInstruction = plan.length === 6 ? plan[3] : undefined
  const callInstruction = plan.length === 6 ? plan[4] : plan[3]
  const returnInstruction = plan.length === 6 ? plan[5] : plan[4]
  if (
    lookupInstruction === undefined ||
    rangeInstruction?.opcode !== 'push-range' ||
    rangeInstruction.refKind !== 'cells' ||
    rangeInstruction.sheetEndName !== undefined ||
    indexInstruction?.opcode !== 'push-number' ||
    callInstruction?.opcode !== 'call' ||
    returnInstruction?.opcode !== 'return'
  ) {
    return undefined
  }
  const callee = callInstruction.callee.trim().toUpperCase()
  if ((callee !== 'VLOOKUP' && callee !== 'HLOOKUP') || (callInstruction.argc !== 3 && callInstruction.argc !== 4)) {
    return undefined
  }
  if ((callInstruction.argc === 3) !== (plan.length === 5)) {
    return undefined
  }
  const rangeLookup = rangeLookupInstruction === undefined ? true : staticLookupBoolean(rangeLookupInstruction)
  if (rangeLookup === undefined) {
    return undefined
  }
  const index = Math.trunc(indexInstruction.value)
  if (!Number.isFinite(index)) {
    return undefined
  }
  return {
    callee,
    rangeLookup,
    lookupInstruction,
    rangeInstruction,
    index,
  }
}

function staticLookupBoolean(instruction: JsPlanInstruction): boolean | undefined {
  if (instruction?.opcode === 'push-boolean') {
    return instruction.value
  }
  if (instruction?.opcode === 'push-number') {
    return instruction.value !== 0
  }
  if (instruction?.opcode === 'push-string') {
    const normalized = instruction.value.trim().toUpperCase()
    if (normalized === 'TRUE') {
      return true
    }
    if (normalized === 'FALSE') {
      return false
    }
  }
  return undefined
}

function readStaticLookupOperand(
  state: StaticTableLookupState,
  ownerSheetName: string,
  instruction: JsPlanInstruction,
): CellValue | undefined {
  switch (instruction.opcode) {
    case 'push-cell':
      return readCellValueByAddress(state, instruction.sheetName ?? ownerSheetName, instruction.address)
    case 'push-number':
      return { tag: ValueTag.Number, value: instruction.value }
    case 'push-boolean':
      return { tag: ValueTag.Boolean, value: instruction.value }
    case 'push-string':
      return { tag: ValueTag.String, value: instruction.value, stringId: 0 }
    case 'push-error':
      return { tag: ValueTag.Error, code: instruction.code }
    case 'begin-scope':
    case 'binary':
    case 'bind-name':
    case 'call':
    case 'end-scope':
    case 'invoke':
    case 'jump':
    case 'jump-if-false':
    case 'lookup-approximate-match':
    case 'lookup-exact-match':
    case 'make-array':
    case 'push-lambda':
    case 'push-name':
    case 'push-omitted':
    case 'push-range':
    case 'return':
    case 'unary':
      return undefined
  }
}

function readCellValueByAddress(state: StaticTableLookupState, sheetName: string, address: string): CellValue {
  const sheet = state.workbook.getSheet(sheetName)
  if (!sheet) {
    return errorValue(ErrorCode.Ref)
  }
  const parsed = parseRangeAddress(`${address}:${address}`, sheetName)
  if (parsed.kind !== 'cells') {
    return errorValue(ErrorCode.Ref)
  }
  return readCellValueAt(state, parsed.sheetName ?? sheetName, parsed.start.row, parsed.start.col)
}

function readCellValueAt(state: StaticTableLookupState, sheetName: string, row: number, col: number): CellValue {
  const sheet = state.workbook.getSheet(sheetName)
  if (!sheet) {
    return errorValue(ErrorCode.Ref)
  }
  const cellIndex = sheet.logical.getVisibleCell(row, col)
  if (cellIndex === undefined) {
    return emptyValue()
  }
  return state.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : state.strings.get(stringId)))
}

function evaluateApproximateVlookup(
  state: StaticTableLookupState,
  formula: RuntimeFormula,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  lookupCol: number,
  resultCol: number,
  lookupValue: CellValue,
): CellValue {
  const prepared = getStaticNumericApproximateVlookupCache(state, formula, sheetName, rowStart, rowEnd, lookupCol)
  if (prepared !== undefined) {
    const preparedMatch = findPreparedNumericApproximateVlookupRow(prepared, lookupValue)
    if (preparedMatch !== undefined) {
      return typeof preparedMatch === 'number'
        ? coerceLookupReturnValue(readCellValueAt(state, sheetName, preparedMatch, resultCol))
        : preparedMatch
    }
  }
  const matchRow = findStaticApproximateMatchIndex(
    rowStart,
    rowEnd,
    (row) => readCellValueAt(state, sheetName, row, lookupCol),
    lookupValue,
  )
  return typeof matchRow === 'number' ? coerceLookupReturnValue(readCellValueAt(state, sheetName, matchRow, resultCol)) : matchRow
}

function getStaticNumericApproximateVlookupCache(
  state: StaticTableLookupState,
  formula: RuntimeFormula,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  lookupCol: number,
): StaticNumericApproximateVlookupCache | undefined {
  const sheet = state.workbook.getSheet(sheetName)
  if (!sheet) {
    return undefined
  }
  const columnVersion = sheet.columnVersions[lookupCol] ?? 0
  const cached = STATIC_NUMERIC_APPROXIMATE_VLOOKUP_CACHE.get(formula)
  if (
    cached !== undefined &&
    cached.sheetId === sheet.id &&
    cached.sheetName === sheetName &&
    cached.rowStart === rowStart &&
    cached.rowEnd === rowEnd &&
    cached.col === lookupCol &&
    cached.structureVersion === sheet.structureVersion &&
    cached.columnVersion === columnVersion
  ) {
    return cached
  }
  const prepared = prepareStaticNumericApproximateVlookupCache(
    state,
    sheetName,
    sheet.id,
    sheet.structureVersion,
    columnVersion,
    rowStart,
    rowEnd,
    lookupCol,
  )
  if (prepared !== undefined) {
    STATIC_NUMERIC_APPROXIMATE_VLOOKUP_CACHE.set(formula, prepared)
  } else {
    STATIC_NUMERIC_APPROXIMATE_VLOOKUP_CACHE.delete(formula)
  }
  return prepared
}

function prepareStaticNumericApproximateVlookupCache(
  state: StaticTableLookupState,
  sheetName: string,
  sheetId: number,
  structureVersion: number,
  columnVersion: number,
  rowStart: number,
  rowEnd: number,
  lookupCol: number,
): StaticNumericApproximateVlookupCache | undefined {
  const length = rowEnd - rowStart + 1
  if (length <= 0) {
    return undefined
  }
  const values = new Float64Array(length)
  let previous = Number.NEGATIVE_INFINITY
  for (let offset = 0; offset < length; offset += 1) {
    const value = staticNumberValue(readCellValueAt(state, sheetName, rowStart + offset, lookupCol))
    if (value === undefined || value < previous) {
      return undefined
    }
    values[offset] = value
    previous = value
  }
  return {
    col: lookupCol,
    columnVersion,
    rowStart,
    rowEnd,
    sheetId,
    sheetName,
    structureVersion,
    values,
  }
}

function findPreparedNumericApproximateVlookupRow(
  prepared: StaticNumericApproximateVlookupCache,
  lookupValue: CellValue,
): number | CellValue | undefined {
  const lookupNumber = staticNumberValue(lookupValue)
  if (lookupNumber === undefined) {
    return undefined
  }
  let low = 0
  let high = prepared.values.length - 1
  let exact = -1
  let best = -1
  while (low <= high) {
    const mid = (low + high) >> 1
    const value = prepared.values[mid]!
    if (value === lookupNumber) {
      exact = mid
      high = mid - 1
    } else if (value < lookupNumber) {
      best = mid
      low = mid + 1
    } else {
      high = mid - 1
    }
  }
  if (exact >= 0) {
    return prepared.rowStart + exact
  }
  return best >= 0 ? prepared.rowStart + best : errorValue(ErrorCode.NA)
}

function evaluateApproximateHlookup(
  state: StaticTableLookupState,
  sheetName: string,
  colStart: number,
  colEnd: number,
  lookupRow: number,
  resultRow: number,
  lookupValue: CellValue,
): CellValue {
  const matchCol = findStaticApproximateMatchIndex(
    colStart,
    colEnd,
    (col) => readCellValueAt(state, sheetName, lookupRow, col),
    lookupValue,
  )
  return typeof matchCol === 'number' ? coerceLookupReturnValue(readCellValueAt(state, sheetName, resultRow, matchCol)) : matchCol
}

function findStaticApproximateMatchIndex(
  start: number,
  end: number,
  readLookupKey: (index: number) => CellValue,
  lookupValue: CellValue,
): number | CellValue {
  let matchedIndex = -1
  let approximateSearchDone = false
  let approximateError: CellValue | undefined
  for (let index = start; index <= end; index += 1) {
    const lookupKey = readLookupKey(index)
    const comparison = compareStaticLookupScalars(lookupKey, lookupValue)
    if (comparison === 0) {
      return index
    }
    if (approximateSearchDone) {
      continue
    }
    if (comparison === undefined) {
      if (lookupKey.tag === ValueTag.Empty) {
        continue
      }
      approximateError = errorValue(ErrorCode.Value)
      approximateSearchDone = true
      continue
    }
    if (comparison < 0) {
      matchedIndex = index
      continue
    }
    approximateSearchDone = true
  }
  if (approximateError !== undefined) {
    return approximateError
  }
  return matchedIndex === -1 ? errorValue(ErrorCode.NA) : matchedIndex
}

function exactLookupKey(state: StaticTableLookupState, value: CellValue): string | undefined {
  return normalizeExactLookupKey(
    value,
    (stringId) => (stringId === 0 ? '' : state.strings.get(stringId)),
    value.tag === ValueTag.String ? value.stringId : 0,
  )
}

function coerceLookupReturnValue(value: CellValue): CellValue {
  return value.tag === ValueTag.Empty ? { tag: ValueTag.Number, value: 0 } : value
}

function compareStaticLookupScalars(left: CellValue, right: CellValue): number | undefined {
  if ((left.tag === ValueTag.String || left.tag === ValueTag.Empty) && (right.tag === ValueTag.String || right.tag === ValueTag.Empty)) {
    const normalizedLeft = staticStringValue(left).toUpperCase()
    const normalizedRight = staticStringValue(right).toUpperCase()
    if (normalizedLeft === normalizedRight) {
      return 0
    }
    return normalizedLeft < normalizedRight ? -1 : 1
  }
  const leftNumber = staticNumberValue(left)
  const rightNumber = staticNumberValue(right)
  if (leftNumber === undefined || rightNumber === undefined) {
    return undefined
  }
  const normalizedLeft = Object.is(leftNumber, -0) ? 0 : leftNumber
  const normalizedRight = Object.is(rightNumber, -0) ? 0 : rightNumber
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}

function staticStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.String:
      return value.value
    case ValueTag.Number:
      return String(Object.is(value.value, -0) ? 0 : value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.Error:
      return ''
  }
}

function staticNumberValue(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return 0
    case ValueTag.Number:
      return Number.isFinite(value.value) ? value.value : undefined
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.String: {
      const trimmed = value.value.trim()
      if (trimmed.length === 0) {
        return 0
      }
      const parsed = Number(trimmed)
      return Number.isFinite(parsed) ? parsed : undefined
    }
    case ValueTag.Error:
      return undefined
  }
}
