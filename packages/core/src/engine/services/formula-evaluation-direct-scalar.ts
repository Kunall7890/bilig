import { parseArithmeticScalarText, type ExcelDateSystem } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeDirectScalarOperand, RuntimeFormula } from '../runtime-state.js'
import { directErrorResult, directNumberResult } from './formula-evaluation-helpers.js'

function readDirectScalarOperand(
  operand: RuntimeDirectScalarOperand,
  readCellValueByIndex: (cellIndex: number | undefined) => CellValue,
): CellValue | undefined {
  if (operand.kind === 'literal-number') {
    return { tag: ValueTag.Number, value: operand.value }
  }
  if (operand.kind === 'error') {
    return directErrorResult(operand.code)
  }
  return readCellValueByIndex(operand.cellIndex)
}

function coerceDirectScalarNumeric(
  value: CellValue | undefined,
  dateSystem: ExcelDateSystem,
): { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode } | undefined {
  if (!value) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.Boolean:
      return { kind: 'number', value: value.value ? 1 : 0 }
    case ValueTag.Empty:
      return { kind: 'number', value: 0 }
    case ValueTag.Error:
      return { kind: 'error', code: value.code }
    case ValueTag.String: {
      const numeric = parseArithmeticScalarText(value.value, dateSystem)
      return numeric === undefined ? { kind: 'error', code: ErrorCode.Value } : { kind: 'number', value: numeric }
    }
    default:
      return undefined
  }
}

export function tryEvaluateDirectScalar(
  formula: RuntimeFormula,
  readCellValueByIndex: (cellIndex: number | undefined) => CellValue,
  dateSystem: ExcelDateSystem = '1900',
): CellValue | undefined {
  const directScalar = formula.directScalar
  if (!directScalar) {
    return undefined
  }
  if (directScalar.kind === 'abs') {
    const operand = coerceDirectScalarNumeric(readDirectScalarOperand(directScalar.operand, readCellValueByIndex), dateSystem)
    if (operand === undefined) {
      return undefined
    }
    if (operand.kind === 'error') {
      return directErrorResult(operand.code)
    }
    return directNumberResult(Math.abs(operand.value))
  }
  const left = coerceDirectScalarNumeric(readDirectScalarOperand(directScalar.left, readCellValueByIndex), dateSystem)
  const right = coerceDirectScalarNumeric(readDirectScalarOperand(directScalar.right, readCellValueByIndex), dateSystem)
  if (left === undefined || right === undefined) {
    return undefined
  }
  if (left.kind === 'error') {
    return directErrorResult(left.code)
  }
  if (right.kind === 'error') {
    return directErrorResult(right.code)
  }
  let result: number
  switch (directScalar.operator) {
    case '+':
      result = left.value + right.value
      break
    case '-':
      result = left.value - right.value
      break
    case '*':
      result = left.value * right.value
      break
    case '/':
      if (right.value === 0) {
        return directErrorResult(ErrorCode.Div0)
      }
      result = left.value / right.value
      break
  }
  return directNumberResult(result + (directScalar.resultOffset ?? 0))
}
