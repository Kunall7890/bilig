import { parseNumericText, parseRangeAddress, type JsPlanInstruction } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { emptyValue, errorValue } from '../../engine-value-utils.js'
import type { EngineRuntimeState, RuntimeFormula } from '../runtime-state.js'

type StaticAggregateExpressionState = Pick<EngineRuntimeState, 'workbook' | 'strings'>

type StaticAggregateRange = {
  readonly kind: 'range'
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

type StaticAggregateStackValue = { readonly kind: 'value'; readonly value: CellValue } | StaticAggregateRange

type StaticAggregateKind = 'sum' | 'average' | 'count' | 'min' | 'max'

const MAX_STATIC_AGGREGATE_PLAN_LENGTH = 64

export function tryEvaluateFormulaLeafStaticAggregateExpression(args: {
  readonly state: StaticAggregateExpressionState
  readonly formula: RuntimeFormula
}): CellValue | undefined {
  const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[args.formula.cellIndex]!)
  if (!ownerSheetName) {
    return undefined
  }
  const plan = args.formula.compiled.jsPlan
  if (plan.length < 3 || plan.length > MAX_STATIC_AGGREGATE_PLAN_LENGTH) {
    return undefined
  }

  const stack: StaticAggregateStackValue[] = []
  let aggregateCallCount = 0
  for (let index = 0; index < plan.length; index += 1) {
    const instruction = plan[index]!
    switch (instruction.opcode) {
      case 'push-number':
        stack.push({ kind: 'value', value: { tag: ValueTag.Number, value: instruction.value } })
        break
      case 'push-boolean':
        stack.push({ kind: 'value', value: { tag: ValueTag.Boolean, value: instruction.value } })
        break
      case 'push-string':
        stack.push({ kind: 'value', value: { tag: ValueTag.String, value: instruction.value, stringId: 0 } })
        break
      case 'push-error':
        stack.push({ kind: 'value', value: errorValue(instruction.code) })
        break
      case 'push-cell':
        stack.push({
          kind: 'value',
          value: readCellValueByAddress(args.state, instruction.sheetName ?? ownerSheetName, instruction.address),
        })
        break
      case 'push-range': {
        const range = parseStaticAggregateRange(instruction, ownerSheetName)
        if (!range || !args.state.workbook.getSheet(range.sheetName)) {
          return undefined
        }
        stack.push(range)
        break
      }
      case 'unary': {
        const operand = stack.pop()
        if (!operand || operand.kind !== 'value') {
          return undefined
        }
        const number = arithmeticNumber(operand.value)
        if (number.kind === 'error') {
          stack.push({ kind: 'value', value: number.value })
          break
        }
        stack.push({
          kind: 'value',
          value: { tag: ValueTag.Number, value: instruction.operator === '-' ? -number.value : number.value },
        })
        break
      }
      case 'binary': {
        const right = stack.pop()
        const left = stack.pop()
        if (!left || !right || left.kind !== 'value' || right.kind !== 'value') {
          return undefined
        }
        const value = evaluateArithmeticBinary(left.value, right.value, instruction.operator)
        if (!value) {
          return undefined
        }
        stack.push({ kind: 'value', value })
        break
      }
      case 'call': {
        const aggregateKind = staticAggregateKind(instruction.callee)
        if (!aggregateKind || instruction.argc !== 1) {
          return undefined
        }
        const range = stack.pop()
        if (!range || range.kind !== 'range') {
          return undefined
        }
        aggregateCallCount += 1
        stack.push({
          kind: 'value',
          value: evaluateStaticAggregateRange(args.state, range, aggregateKind),
        })
        break
      }
      case 'return': {
        if (index !== plan.length - 1 || stack.length !== 1 || stack[0]?.kind !== 'value' || aggregateCallCount === 0) {
          return undefined
        }
        return stack[0].value
      }
      case 'begin-scope':
      case 'bind-name':
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
        return undefined
    }
  }
  return undefined
}

function parseStaticAggregateRange(
  instruction: Extract<JsPlanInstruction, { readonly opcode: 'push-range' }>,
  ownerSheetName: string,
): StaticAggregateRange | undefined {
  if (instruction.refKind !== 'cells' || instruction.sheetEndName !== undefined) {
    return undefined
  }
  const sheetName = instruction.sheetName ?? ownerSheetName
  const parsedRange = parseRangeAddress(`${instruction.start}:${instruction.end}`, sheetName)
  if (parsedRange.kind !== 'cells') {
    return undefined
  }
  return {
    kind: 'range',
    sheetName: parsedRange.sheetName ?? sheetName,
    rowStart: parsedRange.start.row,
    rowEnd: parsedRange.end.row,
    colStart: parsedRange.start.col,
    colEnd: parsedRange.end.col,
  }
}

function staticAggregateKind(callee: string): StaticAggregateKind | undefined {
  switch (callee.trim().toUpperCase()) {
    case 'SUM':
      return 'sum'
    case 'AVERAGE':
    case 'AVG':
      return 'average'
    case 'COUNT':
      return 'count'
    case 'MIN':
      return 'min'
    case 'MAX':
      return 'max'
    default:
      return undefined
  }
}

function evaluateStaticAggregateRange(
  state: StaticAggregateExpressionState,
  range: StaticAggregateRange,
  aggregateKind: StaticAggregateKind,
): CellValue {
  let sum = 0
  let count = 0
  let averageCount = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  for (let col = range.colStart; col <= range.colEnd; col += 1) {
    for (let row = range.rowStart; row <= range.rowEnd; row += 1) {
      const value = readCellValueAt(state, range.sheetName, row, col)
      switch (value.tag) {
        case ValueTag.Number:
          sum += value.value
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, value.value)
          maximum = Math.max(maximum, value.value)
          break
        case ValueTag.Boolean: {
          const booleanNumber = value.value ? 1 : 0
          sum += booleanNumber
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, booleanNumber)
          maximum = Math.max(maximum, booleanNumber)
          break
        }
        case ValueTag.Error:
          if (aggregateKind === 'sum' || aggregateKind === 'average') {
            return value
          }
          break
        case ValueTag.Empty:
        case ValueTag.String:
          break
      }
    }
  }
  switch (aggregateKind) {
    case 'sum':
      return { tag: ValueTag.Number, value: sum }
    case 'count':
      return { tag: ValueTag.Number, value: count }
    case 'average':
      return averageCount === 0 ? errorValue(ErrorCode.Div0) : { tag: ValueTag.Number, value: sum / averageCount }
    case 'min':
      return { tag: ValueTag.Number, value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum }
    case 'max':
      return { tag: ValueTag.Number, value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum }
  }
}

function evaluateArithmeticBinary(
  leftValue: CellValue,
  rightValue: CellValue,
  operator: Extract<JsPlanInstruction, { readonly opcode: 'binary' }>['operator'],
): CellValue | undefined {
  if (operator !== '+' && operator !== '-' && operator !== '*' && operator !== '/' && operator !== '^') {
    return undefined
  }
  const left = arithmeticNumber(leftValue)
  if (left.kind === 'error') {
    return left.value
  }
  const right = arithmeticNumber(rightValue)
  if (right.kind === 'error') {
    return right.value
  }
  switch (operator) {
    case '+':
      return { tag: ValueTag.Number, value: left.value + right.value }
    case '-':
      return { tag: ValueTag.Number, value: left.value - right.value }
    case '*':
      return { tag: ValueTag.Number, value: left.value * right.value }
    case '/':
      return right.value === 0 ? errorValue(ErrorCode.Div0) : { tag: ValueTag.Number, value: left.value / right.value }
    case '^':
      return { tag: ValueTag.Number, value: left.value ** right.value }
  }
}

function arithmeticNumber(
  value: CellValue,
): { readonly kind: 'number'; readonly value: number } | { readonly kind: 'error'; readonly value: CellValue } {
  switch (value.tag) {
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.Boolean:
      return { kind: 'number', value: value.value ? 1 : 0 }
    case ValueTag.Empty:
      return { kind: 'number', value: 0 }
    case ValueTag.String: {
      const parsed = parseNumericText(value.value)
      return parsed === undefined ? { kind: 'error', value: errorValue(ErrorCode.Value) } : { kind: 'number', value: parsed }
    }
    case ValueTag.Error:
      return { kind: 'error', value }
  }
}

function readCellValueByAddress(state: StaticAggregateExpressionState, sheetName: string, address: string): CellValue {
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

function readCellValueAt(state: StaticAggregateExpressionState, sheetName: string, row: number, col: number): CellValue {
  const sheet = state.workbook.getSheet(sheetName)
  if (!sheet) {
    return errorValue(ErrorCode.Ref)
  }
  const cellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
  if (cellIndex === -1) {
    return emptyValue()
  }
  return state.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : state.strings.get(stringId)))
}
