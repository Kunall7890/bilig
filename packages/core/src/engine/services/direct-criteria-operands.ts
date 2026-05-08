import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin, isArrayValue, type EvaluationResult } from '@bilig/formula'
import type { RuntimeDirectCriteriaOperand } from '../runtime-state.js'

const dateBuiltin = getBuiltin('DATE')
const edateBuiltin = getBuiltin('EDATE')
const monthBuiltin = getBuiltin('MONTH')
const yearBuiltin = getBuiltin('YEAR')

function errorResult(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function scalarBuiltinResult(value: EvaluationResult | undefined): CellValue {
  if (!value) {
    return errorResult(ErrorCode.Name)
  }
  if (isArrayValue(value)) {
    return errorResult(ErrorCode.Value)
  }
  return value
}

function monthBoundaryValue(source: CellValue, offsetMonths: number): CellValue {
  if (source.tag === ValueTag.Error) {
    return source
  }
  const year = scalarBuiltinResult(yearBuiltin?.(source))
  if (year.tag === ValueTag.Error) {
    return year
  }
  const month = scalarBuiltinResult(monthBuiltin?.(source))
  if (month.tag === ValueTag.Error) {
    return month
  }
  const start = scalarBuiltinResult(dateBuiltin?.(year, month, numberResult(1)))
  if (start.tag === ValueTag.Error || offsetMonths === 0) {
    return start
  }
  return scalarBuiltinResult(edateBuiltin?.(start, numberResult(offsetMonths)))
}

export function readRuntimeDirectCriteriaOperandValue(args: {
  readonly operand: RuntimeDirectCriteriaOperand
  readonly readCellValueByIndex: (cellIndex: number | undefined) => CellValue
  readonly stringifyCriteriaValue: (value: CellValue) => string
}): CellValue {
  if (args.operand.kind === 'literal') {
    return args.operand.value
  }
  const value = args.readCellValueByIndex(args.operand.cellIndex)
  if (args.operand.kind === 'cell') {
    return value
  }
  if (value.tag === ValueTag.Error) {
    return value
  }
  const criteriaValue =
    args.operand.kind === 'cell-month-boundary-string-concat' ? monthBoundaryValue(value, args.operand.offsetMonths) : value
  if (criteriaValue.tag === ValueTag.Error) {
    return criteriaValue
  }
  return {
    tag: ValueTag.String,
    value: `${args.operand.prefix}${args.stringifyCriteriaValue(criteriaValue)}${args.operand.suffix}`,
    stringId: 0,
  }
}
