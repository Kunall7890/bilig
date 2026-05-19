import { parseRangeAddress, type JsPlanInstruction } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { emptyValue, errorValue } from '../../engine-value-utils.js'
import type { EngineRuntimeState, RuntimeFormula } from '../runtime-state.js'
import { normalizeExactLookupKey } from './direct-lookup-helpers.js'

type StaticExactTableLookupPlan = {
  readonly callee: 'VLOOKUP' | 'HLOOKUP'
  readonly lookupInstruction: JsPlanInstruction
  readonly rangeInstruction: Extract<JsPlanInstruction, { opcode: 'push-range' }>
  readonly index: number
}

type StaticTableLookupState = Pick<EngineRuntimeState, 'workbook' | 'strings'>

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
  const lookupKey = exactLookupKey(args.state, lookupValue)
  if (lookupKey === undefined) {
    return errorValue(ErrorCode.NA)
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
  for (let col = colStart; col <= colEnd; col += 1) {
    if (exactLookupKey(args.state, readCellValueAt(args.state, sheetName, rowStart, col)) === lookupKey) {
      return coerceLookupReturnValue(readCellValueAt(args.state, sheetName, resultRow, col))
    }
  }
  return errorValue(ErrorCode.NA)
}

function parseStaticExactTableLookupPlan(plan: readonly JsPlanInstruction[]): StaticExactTableLookupPlan | undefined {
  if (plan.length !== 6) {
    return undefined
  }
  const [lookupInstruction, rangeInstruction, indexInstruction, rangeLookupInstruction, callInstruction, returnInstruction] = plan
  if (
    lookupInstruction === undefined ||
    rangeInstruction?.opcode !== 'push-range' ||
    rangeInstruction.refKind !== 'cells' ||
    rangeInstruction.sheetEndName !== undefined ||
    indexInstruction?.opcode !== 'push-number' ||
    !isExactLookupFalse(rangeLookupInstruction) ||
    callInstruction?.opcode !== 'call' ||
    returnInstruction?.opcode !== 'return'
  ) {
    return undefined
  }
  const callee = callInstruction.callee.trim().toUpperCase()
  if ((callee !== 'VLOOKUP' && callee !== 'HLOOKUP') || callInstruction.argc !== 4) {
    return undefined
  }
  const index = Math.trunc(indexInstruction.value)
  if (!Number.isFinite(index)) {
    return undefined
  }
  return {
    callee,
    lookupInstruction,
    rangeInstruction,
    index,
  }
}

function isExactLookupFalse(instruction: JsPlanInstruction | undefined): boolean {
  if (instruction?.opcode === 'push-boolean') {
    return !instruction.value
  }
  return instruction?.opcode === 'push-number' && Object.is(instruction.value, 0)
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
