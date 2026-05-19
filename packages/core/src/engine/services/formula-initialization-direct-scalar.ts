import { parseArithmeticNumericText } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor, RuntimeDirectScalarOperand } from '../runtime-state.js'

type DirectScalarCellRead = { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode }

function readInitialStringCell(state: Pick<EngineRuntimeState, 'workbook' | 'strings'>, cellIndex: number): string {
  const stringId = state.workbook.cellStore.stringIds[cellIndex] ?? 0
  return stringId === 0 ? '' : state.strings.get(stringId)
}

function coerceInitialDirectScalarCell(
  state: Pick<EngineRuntimeState, 'workbook' | 'strings'>,
  cellIndex: number,
): DirectScalarCellRead | undefined {
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
  switch (tag) {
    case ValueTag.Number:
      return { kind: 'number', value: cellStore.numbers[cellIndex] ?? 0 }
    case ValueTag.Boolean:
      return { kind: 'number', value: (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0 }
    case ValueTag.Empty:
      return { kind: 'number', value: 0 }
    case ValueTag.Error:
      return { kind: 'error', code: (cellStore.errors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
    case ValueTag.String: {
      const numeric = parseArithmeticNumericText(readInitialStringCell(state, cellIndex))
      return numeric === undefined ? { kind: 'error', code: ErrorCode.Value } : { kind: 'number', value: numeric }
    }
    default:
      return undefined
  }
}

function readInitialDirectScalarOperand(
  state: Pick<EngineRuntimeState, 'workbook' | 'strings'>,
  operand: RuntimeDirectScalarOperand,
): DirectScalarCellRead | undefined {
  switch (operand.kind) {
    case 'literal-number':
      return { kind: 'number', value: operand.value }
    case 'error':
      return { kind: 'error', code: operand.code }
    case 'cell':
      return coerceInitialDirectScalarCell(state, operand.cellIndex)
  }
}

export function evaluateInitialDirectScalar(
  state: Pick<EngineRuntimeState, 'workbook' | 'strings'>,
  directScalar: RuntimeDirectScalarDescriptor,
): CellValue | undefined {
  if (directScalar.kind === 'abs') {
    const operand = readInitialDirectScalarOperand(state, directScalar.operand)
    if (!operand) {
      return undefined
    }
    return operand.kind === 'error' ? { tag: ValueTag.Error, code: operand.code } : { tag: ValueTag.Number, value: Math.abs(operand.value) }
  }
  const left = readInitialDirectScalarOperand(state, directScalar.left)
  const right = readInitialDirectScalarOperand(state, directScalar.right)
  if (!left || !right) {
    return undefined
  }
  if (left.kind === 'error') {
    return { tag: ValueTag.Error, code: left.code }
  }
  if (right.kind === 'error') {
    return { tag: ValueTag.Error, code: right.code }
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
        return { tag: ValueTag.Error, code: ErrorCode.Div0 }
      }
      result = left.value / right.value
      break
  }
  return { tag: ValueTag.Number, value: result + (directScalar.resultOffset ?? 0) }
}

function coerceInitialDirectScalarNumber(state: Pick<EngineRuntimeState, 'workbook' | 'strings'>, cellIndex: number): number | undefined {
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
  switch (tag) {
    case ValueTag.Number:
      return cellStore.numbers[cellIndex] ?? 0
    case ValueTag.Boolean:
      return (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
      return parseArithmeticNumericText(readInitialStringCell(state, cellIndex))
    case ValueTag.Error:
      return undefined
    default:
      return undefined
  }
}

function readInitialDirectScalarNumberOperand(
  state: Pick<EngineRuntimeState, 'workbook' | 'strings'>,
  operand: RuntimeDirectScalarOperand,
): number | undefined {
  switch (operand.kind) {
    case 'literal-number':
      return operand.value
    case 'cell':
      return coerceInitialDirectScalarNumber(state, operand.cellIndex)
    case 'error':
      return undefined
  }
}

export function evaluateInitialDirectScalarNumber(
  state: Pick<EngineRuntimeState, 'workbook' | 'strings'>,
  directScalar: RuntimeDirectScalarDescriptor,
): number | undefined {
  if (directScalar.kind === 'abs') {
    const operand = readInitialDirectScalarNumberOperand(state, directScalar.operand)
    return operand === undefined ? undefined : Math.abs(operand)
  }
  const left = readInitialDirectScalarNumberOperand(state, directScalar.left)
  const right = readInitialDirectScalarNumberOperand(state, directScalar.right)
  if (left === undefined || right === undefined) {
    return undefined
  }
  let result: number
  switch (directScalar.operator) {
    case '+':
      result = left + right
      break
    case '-':
      result = left - right
      break
    case '*':
      result = left * right
      break
    case '/':
      if (right === 0) {
        return undefined
      }
      result = left / right
      break
  }
  return result + (directScalar.resultOffset ?? 0)
}
