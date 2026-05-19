import { ErrorCode, ValueTag, formatGeneralNumberValue, type CellValue, type NumberValue } from '@bilig/protocol'
import {
  getBuiltin,
  getDateSystemBuiltin,
  isArrayValue,
  normalizeBuiltinLookupName,
  parseNumericText,
  type CompiledFormula,
  type JsPlanInstruction,
  type ParsedDependencyReference,
} from '@bilig/formula'
import {
  INLINE_SCALAR_FAST_PLAN_ARITHMETIC,
  INLINE_SCALAR_FAST_PLAN_CONCAT,
  INLINE_SCALAR_FAST_PLAN_IF_STRING,
  INLINE_SCALAR_FAST_PLAN_LEN,
  INLINE_SCALAR_FAST_PLAN_MIN_MAX,
  INLINE_SCALAR_FAST_PLAN_PMT,
  INLINE_SCALAR_FAST_PLAN_ROUND_SQRT,
  type EngineRuntimeState,
  type InlineScalarFastPlanKind,
  type RuntimeFormula,
} from '../runtime-state.js'
import {
  compareInlineScalars,
  inlineArithmeticNumber,
  inlineNumber,
  inlineOptionalNumber,
  inlinePaymentType,
  inlineRequiredNumber,
  inlineStringValue,
  inlineTruthy,
  roundInlineToDigits,
} from './formula-leaf-inline-scalar-values.js'

type InlineScalarState = Pick<EngineRuntimeState, 'workbook' | 'strings'>

const MAX_INLINE_PLAN_LENGTH = 48
const MAX_INLINE_STACK_DEPTH = 16
const INVALID_INLINE_CELL_INDEX = 0xffffffff

const numberValue = (value: number): NumberValue => ({ tag: ValueTag.Number, value })
const booleanValue = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const stringValue = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const errorValue = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

export function tryEvaluateFormulaLeafInlineScalar(args: {
  readonly state: InlineScalarState
  readonly formula: RuntimeFormula
}): CellValue | undefined {
  const plan = args.formula.compiled.jsPlan
  if (
    plan.length === 0 ||
    plan.length > MAX_INLINE_PLAN_LENGTH ||
    args.formula.compiled.producesSpill ||
    args.formula.compiled.astMatchesSource === false
  ) {
    return undefined
  }

  const fastPlanResult = tryEvaluateFormulaLeafInlineScalarFastPlan(args)
  if (fastPlanResult !== undefined) {
    return fastPlanResult
  }

  const dependencyResolver = createInlineDependencyResolver(args.state, args.formula)
  if (!dependencyResolver) {
    return undefined
  }

  const stack: CellValue[] = []
  let pc = 0
  while (pc < plan.length) {
    const instruction = plan[pc]
    if (!instruction) {
      return undefined
    }
    switch (instruction.opcode) {
      case 'push-number':
        if (!pushInlineValue(stack, numberValue(instruction.value))) {
          return undefined
        }
        break
      case 'push-boolean':
        if (!pushInlineValue(stack, booleanValue(instruction.value))) {
          return undefined
        }
        break
      case 'push-string':
        if (!pushInlineValue(stack, stringValue(instruction.value))) {
          return undefined
        }
        break
      case 'push-error':
        if (!pushInlineValue(stack, errorValue(instruction.code))) {
          return undefined
        }
        break
      case 'push-cell': {
        const value = dependencyResolver(instruction, pc)
        if (!value || !pushInlineValue(stack, value)) {
          return undefined
        }
        break
      }
      case 'unary': {
        const value = stack.pop()
        if (!value || !pushInlineValue(stack, evaluateInlineUnary(instruction.operator, value))) {
          return undefined
        }
        break
      }
      case 'binary': {
        const right = stack.pop()
        const left = stack.pop()
        if (!left || !right || !pushInlineValue(stack, evaluateInlineBinary(instruction.operator, left, right))) {
          return undefined
        }
        break
      }
      case 'call': {
        if (instruction.argc < 0 || instruction.argc > stack.length) {
          return undefined
        }
        const start = stack.length - instruction.argc
        const result = evaluateInlineCall(args.state, instruction.callee, stack, start, instruction.argc)
        if (!result) {
          return undefined
        }
        stack.length = start
        if (!pushInlineValue(stack, result)) {
          return undefined
        }
        break
      }
      case 'jump-if-false': {
        const value = stack.pop()
        if (!value) {
          return undefined
        }
        if (value.tag === ValueTag.Error) {
          return value
        }
        if (!inlineTruthy(value)) {
          pc = instruction.target
          continue
        }
        break
      }
      case 'jump':
        pc = instruction.target
        continue
      case 'return':
        return stack.length === 1 ? stack[0] : undefined
      case 'begin-scope':
      case 'bind-name':
      case 'end-scope':
      case 'invoke':
      case 'lookup-approximate-match':
      case 'lookup-exact-match':
      case 'make-array':
      case 'push-lambda':
      case 'push-name':
      case 'push-omitted':
      case 'push-range':
        return undefined
    }
    pc += 1
  }

  return undefined
}

function tryEvaluateFormulaLeafInlineScalarFastPlan(args: {
  readonly state: InlineScalarState
  readonly formula: RuntimeFormula
}): CellValue | undefined {
  const cellIndices = args.formula.inlineScalarPlanCellIndices
  if (cellIndices === undefined || cellIndices.length !== args.formula.compiled.jsPlan.length) {
    return undefined
  }
  const plan = args.formula.compiled.jsPlan
  const fastPlanKind = args.formula.inlineScalarFastPlanKind ?? classifyInlineScalarFastPlan(args.formula.compiled)
  switch (fastPlanKind) {
    case INLINE_SCALAR_FAST_PLAN_ARITHMETIC:
      return tryEvaluateFastArithmeticPlan(args.state, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_LEN:
      return tryEvaluateFastLenPlan(args.state, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_CONCAT:
      return tryEvaluateFastConcatPlan(args.state, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_MIN_MAX:
      return tryEvaluateFastMinMaxPlan(args.state, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_ROUND_SQRT:
      return tryEvaluateFastRoundSqrtPlan(args.state, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_PMT:
      return tryEvaluateFastPmtPlan(args.state, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_IF_STRING:
      return tryEvaluateFastIfStringPlan(args.state, plan, cellIndices)
    case undefined:
      return undefined
    default:
      return undefined
  }
}

export function classifyInlineScalarFastPlan(
  compiled: Pick<CompiledFormula, 'astMatchesSource' | 'jsPlan' | 'producesSpill'>,
): InlineScalarFastPlanKind | undefined {
  const plan = compiled.jsPlan
  if (plan.length === 0 || plan.length > MAX_INLINE_PLAN_LENGTH || compiled.producesSpill || compiled.astMatchesSource === false) {
    return undefined
  }
  switch (plan.length) {
    case 5:
      return isFastConcatPlan(plan)
        ? INLINE_SCALAR_FAST_PLAN_CONCAT
        : isFastRoundSqrtPlan(plan)
          ? INLINE_SCALAR_FAST_PLAN_ROUND_SQRT
          : undefined
    case 6:
      return isFastArithmeticPlan(plan) ? INLINE_SCALAR_FAST_PLAN_ARITHMETIC : isFastLenPlan(plan) ? INLINE_SCALAR_FAST_PLAN_LEN : undefined
    case 7:
      return isFastPmtPlan(plan) ? INLINE_SCALAR_FAST_PLAN_PMT : undefined
    case 8:
      return isFastIfStringPlan(plan) ? INLINE_SCALAR_FAST_PLAN_IF_STRING : undefined
    case 10:
      return isFastMinMaxPlan(plan) ? INLINE_SCALAR_FAST_PLAN_MIN_MAX : undefined
    default:
      return undefined
  }
}

function isFastArithmeticPlan(plan: readonly JsPlanInstruction[]): boolean {
  return (
    plan.length === 6 &&
    plan[0]?.opcode === 'push-cell' &&
    plan[1]?.opcode === 'push-cell' &&
    plan[2]?.opcode === 'push-number' &&
    plan[3]?.opcode === 'binary' &&
    plan[4]?.opcode === 'binary' &&
    plan[5]?.opcode === 'return'
  )
}

function tryEvaluateFastArithmeticPlan(
  state: InlineScalarState,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): CellValue | undefined {
  if (!isFastArithmeticPlan(plan)) {
    return undefined
  }
  const left = inlineArithmeticNumberFromCell(state, cellIndices[0])
  const rightLeft = inlineArithmeticNumberFromCell(state, cellIndices[1])
  if (left === undefined || rightLeft === undefined) {
    return undefined
  }
  const literal = inlinePushNumberValue(plan[2])
  const innerOperator = inlineBinaryOperator(plan[3])
  const outerOperator = inlineBinaryOperator(plan[4])
  if (literal === undefined || innerOperator === undefined || outerOperator === undefined) {
    return undefined
  }
  const right = evaluateFastInlineNumericBinary(innerOperator, rightLeft, literal)
  if (right === undefined) {
    return undefined
  }
  const value = evaluateFastInlineNumericBinary(outerOperator, left, right)
  return value === undefined ? undefined : numberValue(value)
}

function isFastLenPlan(plan: readonly JsPlanInstruction[]): boolean {
  return (
    plan.length === 6 &&
    plan[0]?.opcode === 'push-cell' &&
    plan[1]?.opcode === 'call' &&
    plan[1].callee === 'LEN' &&
    plan[1].argc === 1 &&
    plan[2]?.opcode === 'push-cell' &&
    plan[3]?.opcode === 'call' &&
    plan[3].callee === 'LEN' &&
    plan[3].argc === 1 &&
    plan[4]?.opcode === 'binary' &&
    plan[4].operator === '+' &&
    plan[5]?.opcode === 'return'
  )
}

function tryEvaluateFastLenPlan(
  state: InlineScalarState,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): CellValue | undefined {
  if (!isFastLenPlan(plan)) {
    return undefined
  }
  const leftLength = inlineStringLengthFromCell(state, cellIndices[0])
  const rightLength = inlineStringLengthFromCell(state, cellIndices[2])
  return leftLength === undefined || rightLength === undefined ? undefined : numberValue(leftLength + rightLength)
}

function isFastConcatPlan(plan: readonly JsPlanInstruction[]): boolean {
  return (
    plan.length === 5 &&
    plan[0]?.opcode === 'push-cell' &&
    plan[1]?.opcode === 'push-string' &&
    plan[2]?.opcode === 'push-cell' &&
    plan[3]?.opcode === 'call' &&
    plan[3].callee === 'CONCATENATE' &&
    plan[3].argc === 3 &&
    plan[4]?.opcode === 'return'
  )
}

function tryEvaluateFastConcatPlan(
  state: InlineScalarState,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): CellValue | undefined {
  if (!isFastConcatPlan(plan)) {
    return undefined
  }
  const left = inlineStringFromCell(state, cellIndices[0])
  const right = inlineStringFromCell(state, cellIndices[2])
  const separator = inlinePushStringValue(plan[1])
  return left === undefined || right === undefined || separator === undefined ? undefined : stringValue(`${left}${separator}${right}`)
}

function isFastMinMaxPlan(plan: readonly JsPlanInstruction[]): boolean {
  return (
    plan.length === 10 &&
    plan[0]?.opcode === 'push-cell' &&
    plan[1]?.opcode === 'push-cell' &&
    plan[2]?.opcode === 'push-cell' &&
    plan[3]?.opcode === 'call' &&
    plan[3].callee === 'MIN' &&
    plan[3].argc === 3 &&
    plan[4]?.opcode === 'push-cell' &&
    plan[5]?.opcode === 'push-cell' &&
    plan[6]?.opcode === 'push-cell' &&
    plan[7]?.opcode === 'call' &&
    plan[7].callee === 'MAX' &&
    plan[7].argc === 3 &&
    plan[8]?.opcode === 'binary' &&
    plan[8].operator === '+' &&
    plan[9]?.opcode === 'return'
  )
}

function tryEvaluateFastMinMaxPlan(
  state: InlineScalarState,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): CellValue | undefined {
  if (!isFastMinMaxPlan(plan)) {
    return undefined
  }
  const first = inlineNumberFromCell(state, cellIndices[0])
  const second = inlineNumberFromCell(state, cellIndices[1])
  const third = inlineNumberFromCell(state, cellIndices[2])
  if (first === undefined || second === undefined || third === undefined) {
    return undefined
  }
  return numberValue(Math.min(first, second, third) + Math.max(first, second, third))
}

function isFastRoundSqrtPlan(plan: readonly JsPlanInstruction[]): boolean {
  return (
    plan.length === 5 &&
    plan[0]?.opcode === 'push-cell' &&
    plan[1]?.opcode === 'call' &&
    plan[1].callee === 'SQRT' &&
    plan[1].argc === 1 &&
    plan[2]?.opcode === 'push-number' &&
    plan[3]?.opcode === 'call' &&
    plan[3].callee === 'ROUND' &&
    plan[3].argc === 2 &&
    plan[4]?.opcode === 'return'
  )
}

function tryEvaluateFastRoundSqrtPlan(
  state: InlineScalarState,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): CellValue | undefined {
  if (!isFastRoundSqrtPlan(plan)) {
    return undefined
  }
  const value = inlineNumberFromCell(state, cellIndices[0])
  if (value === undefined) {
    return undefined
  }
  const sqrt = Math.sqrt(value)
  if (!Number.isFinite(sqrt)) {
    return errorValue(ErrorCode.Value)
  }
  const digits = inlinePushNumberValue(plan[2])
  return digits === undefined ? undefined : numberValue(roundInlineToDigits(sqrt, Math.trunc(digits)))
}

function isFastPmtPlan(plan: readonly JsPlanInstruction[]): boolean {
  return (
    plan.length === 7 &&
    plan[0]?.opcode === 'push-cell' &&
    plan[1]?.opcode === 'push-number' &&
    plan[2]?.opcode === 'binary' &&
    plan[2].operator === '/' &&
    plan[3]?.opcode === 'push-cell' &&
    plan[4]?.opcode === 'push-cell' &&
    plan[5]?.opcode === 'call' &&
    plan[5].callee === 'PMT' &&
    plan[5].argc === 3 &&
    plan[6]?.opcode === 'return'
  )
}

function tryEvaluateFastPmtPlan(
  state: InlineScalarState,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): CellValue | undefined {
  if (!isFastPmtPlan(plan)) {
    return undefined
  }
  const divisor = inlinePushNumberValue(plan[1])
  const annualRate = inlineNumberFromCell(state, cellIndices[0])
  const periods = inlineNumberFromCell(state, cellIndices[3])
  const present = inlineNumberFromCell(state, cellIndices[4])
  if (
    annualRate === undefined ||
    periods === undefined ||
    present === undefined ||
    divisor === undefined ||
    divisor === 0 ||
    periods <= 0
  ) {
    return undefined
  }
  const rate = annualRate / divisor
  if (rate === 0) {
    return numberValue(-present / periods)
  }
  const growth = (1 + rate) ** periods
  const denominator = growth - 1
  return denominator === 0 ? errorValue(ErrorCode.Value) : numberValue((-rate * present * growth) / denominator)
}

function isFastIfStringPlan(plan: readonly JsPlanInstruction[]): boolean {
  return (
    plan.length === 8 &&
    plan[0]?.opcode === 'push-cell' &&
    plan[1]?.opcode === 'push-number' &&
    plan[2]?.opcode === 'binary' &&
    plan[3]?.opcode === 'jump-if-false' &&
    plan[4]?.opcode === 'push-string' &&
    plan[5]?.opcode === 'jump' &&
    plan[6]?.opcode === 'push-string' &&
    plan[7]?.opcode === 'return'
  )
}

function tryEvaluateFastIfStringPlan(
  state: InlineScalarState,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): CellValue | undefined {
  if (!isFastIfStringPlan(plan)) {
    return undefined
  }
  const left = inlineNumberFromCell(state, cellIndices[0])
  if (left === undefined) {
    return undefined
  }
  const operator = inlineBinaryOperator(plan[2])
  const right = inlinePushNumberValue(plan[1])
  const trueValue = inlinePushStringValue(plan[4])
  const falseValue = inlinePushStringValue(plan[6])
  if (operator === undefined || right === undefined || trueValue === undefined || falseValue === undefined) {
    return undefined
  }
  const comparison = compareFastInlineNumbers(operator, left, right)
  return comparison === undefined ? undefined : stringValue(comparison ? trueValue : falseValue)
}

function inlineBinaryOperator(
  instruction: JsPlanInstruction | undefined,
): Extract<JsPlanInstruction, { opcode: 'binary' }>['operator'] | undefined {
  return instruction?.opcode === 'binary' ? instruction.operator : undefined
}

function inlinePushNumberValue(instruction: JsPlanInstruction | undefined): number | undefined {
  return instruction?.opcode === 'push-number' ? instruction.value : undefined
}

function inlinePushStringValue(instruction: JsPlanInstruction | undefined): string | undefined {
  return instruction?.opcode === 'push-string' ? instruction.value : undefined
}

function inlineNumberFromCell(state: InlineScalarState, cellIndex: number | undefined): number | undefined {
  if (cellIndex === undefined || cellIndex === INVALID_INLINE_CELL_INDEX) {
    return undefined
  }
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] ?? ValueTag.Empty) as ValueTag
  switch (tag) {
    case ValueTag.Number:
      return cellStore.numbers[cellIndex] ?? 0
    case ValueTag.Boolean:
      return (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
    default:
      return undefined
  }
}

function inlineArithmeticNumberFromCell(state: InlineScalarState, cellIndex: number | undefined): number | undefined {
  if (cellIndex === undefined || cellIndex === INVALID_INLINE_CELL_INDEX) {
    return undefined
  }
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] ?? ValueTag.Empty) as ValueTag
  if (tag !== ValueTag.String) {
    return inlineNumberFromCell(state, cellIndex)
  }
  const trimmed = state.strings.get(cellStore.stringIds[cellIndex] ?? 0).trim()
  return trimmed === '' ? 0 : parseNumericText(trimmed)
}

function inlineStringLengthFromCell(state: InlineScalarState, cellIndex: number | undefined): number | undefined {
  const value = inlineStringFromCell(state, cellIndex)
  return value === undefined ? undefined : value.length
}

function inlineStringFromCell(state: InlineScalarState, cellIndex: number | undefined): string | undefined {
  if (cellIndex === undefined || cellIndex === INVALID_INLINE_CELL_INDEX) {
    return undefined
  }
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] ?? ValueTag.Empty) as ValueTag
  switch (tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return formatGeneralNumberValue(cellStore.numbers[cellIndex] ?? 0)
    case ValueTag.Boolean:
      return (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return state.strings.get(cellStore.stringIds[cellIndex] ?? 0)
    case ValueTag.Error:
    default:
      return undefined
  }
}

function evaluateFastInlineNumericBinary(
  operator: Extract<JsPlanInstruction, { opcode: 'binary' }>['operator'],
  left: number,
  right: number,
): number | undefined {
  if (operator === '/' && right === 0) {
    return undefined
  }
  const value =
    operator === '+'
      ? left + right
      : operator === '-'
        ? left - right
        : operator === '*'
          ? left * right
          : operator === '/'
            ? left / right
            : operator === '^'
              ? Math.pow(left, right)
              : undefined
  return value !== undefined && Number.isFinite(value) ? value : undefined
}

function compareFastInlineNumbers(
  operator: Extract<JsPlanInstruction, { opcode: 'binary' }>['operator'],
  left: number,
  right: number,
): boolean | undefined {
  switch (operator) {
    case '=':
      return left === right
    case '<>':
      return left !== right
    case '>':
      return left > right
    case '>=':
      return left >= right
    case '<':
      return left < right
    case '<=':
      return left <= right
    case '&':
    case '*':
    case '+':
    case '-':
    case '/':
    case ':':
    case '^':
      return undefined
    default:
      return undefined
  }
}

function createInlineDependencyResolver(
  state: InlineScalarState,
  formula: RuntimeFormula,
): ((instruction: Extract<JsPlanInstruction, { opcode: 'push-cell' }>, planIndex: number) => CellValue | undefined) | undefined {
  const inlineCellIndices = formula.inlineScalarPlanCellIndices
  if (inlineCellIndices !== undefined && inlineCellIndices.length === formula.compiled.jsPlan.length) {
    return (_instruction, planIndex) => {
      const cellIndex = inlineCellIndices[planIndex]
      if (cellIndex === undefined || cellIndex === INVALID_INLINE_CELL_INDEX) {
        return undefined
      }
      return state.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : state.strings.get(stringId)))
    }
  }

  const parsedDeps = formula.compiled.parsedDeps
  if (parsedDeps === undefined) {
    return undefined
  }
  if (parsedDeps.length !== formula.dependencyIndices.length || parsedDeps.length === 0) {
    return undefined
  }
  for (let index = 0; index < parsedDeps.length; index += 1) {
    if (parsedDeps[index]?.kind !== 'cell') {
      return undefined
    }
  }
  return (instruction) => {
    const cellIndex = findInlineDependencyIndex(parsedDeps, formula.dependencyIndices, instruction)
    if (cellIndex === undefined) {
      return undefined
    }
    return state.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : state.strings.get(stringId)))
  }
}

export function buildInlineScalarPlanCellIndices(
  compiled: Pick<CompiledFormula, 'astMatchesSource' | 'jsPlan' | 'parsedDeps' | 'producesSpill'>,
  dependencyIndices: Uint32Array,
): Uint32Array | undefined {
  const plan = compiled.jsPlan
  if (plan.length === 0 || plan.length > MAX_INLINE_PLAN_LENGTH || compiled.producesSpill || compiled.astMatchesSource === false) {
    return undefined
  }
  const parsedDeps = compiled.parsedDeps
  if (parsedDeps === undefined || parsedDeps.length !== dependencyIndices.length || parsedDeps.length === 0) {
    return undefined
  }
  for (let index = 0; index < parsedDeps.length; index += 1) {
    if (parsedDeps[index]?.kind !== 'cell') {
      return undefined
    }
  }

  let hasPushCell = false
  const cellIndicesByPlanIndex = new Uint32Array(plan.length)
  cellIndicesByPlanIndex.fill(INVALID_INLINE_CELL_INDEX)
  for (let index = 0; index < plan.length; index += 1) {
    const instruction = plan[index]
    if (!instruction || !isInlineScalarPlanInstruction(instruction)) {
      return undefined
    }
    if (instruction.opcode !== 'push-cell') {
      continue
    }
    const cellIndex = findInlineDependencyIndex(parsedDeps, dependencyIndices, instruction)
    if (cellIndex === undefined) {
      return undefined
    }
    cellIndicesByPlanIndex[index] = cellIndex
    hasPushCell = true
  }
  return hasPushCell ? cellIndicesByPlanIndex : undefined
}

function isInlineScalarPlanInstruction(instruction: JsPlanInstruction): boolean {
  switch (instruction.opcode) {
    case 'binary':
    case 'call':
    case 'jump':
    case 'jump-if-false':
    case 'push-boolean':
    case 'push-cell':
    case 'push-error':
    case 'push-number':
    case 'push-string':
    case 'return':
    case 'unary':
      return true
    case 'begin-scope':
    case 'bind-name':
    case 'end-scope':
    case 'invoke':
    case 'lookup-approximate-match':
    case 'lookup-exact-match':
    case 'make-array':
    case 'push-lambda':
    case 'push-name':
    case 'push-omitted':
    case 'push-range':
      return false
  }
}

function findInlineDependencyIndex(
  parsedDeps: readonly ParsedDependencyReference[],
  dependencyIndices: Uint32Array,
  instruction: Extract<JsPlanInstruction, { opcode: 'push-cell' }>,
): number | undefined {
  for (let index = 0; index < parsedDeps.length; index += 1) {
    const dep = parsedDeps[index]!
    if (
      dep.kind === 'cell' &&
      dep.address === instruction.address &&
      (dep.sheetName ?? undefined) === (instruction.sheetName ?? undefined)
    ) {
      return dependencyIndices[index]
    }
  }
  return undefined
}

function pushInlineValue(stack: CellValue[], value: CellValue): boolean {
  if (stack.length >= MAX_INLINE_STACK_DEPTH) {
    return false
  }
  stack.push(value)
  return true
}

function evaluateInlineUnary(operator: '+' | '-', value: CellValue): CellValue {
  if (value.tag === ValueTag.Error) {
    return value
  }
  const numeric = inlineArithmeticNumber(value)
  return numeric === undefined ? errorValue(ErrorCode.Value) : numberValue(operator === '-' ? -numeric : numeric)
}

function evaluateInlineBinary(
  operator: Extract<JsPlanInstruction, { opcode: 'binary' }>['operator'],
  left: CellValue,
  right: CellValue,
): CellValue {
  if (operator === ':') {
    return errorValue(ErrorCode.Value)
  }
  if (left.tag === ValueTag.Error) {
    return left
  }
  if (right.tag === ValueTag.Error) {
    return right
  }
  if (operator === '&') {
    return stringValue(`${inlineStringValue(left)}${inlineStringValue(right)}`)
  }
  if (operator === '+' || operator === '-' || operator === '*' || operator === '/' || operator === '^') {
    const leftNumber = inlineArithmeticNumber(left)
    const rightNumber = inlineArithmeticNumber(right)
    if (leftNumber === undefined || rightNumber === undefined) {
      return errorValue(ErrorCode.Value)
    }
    if (operator === '/' && rightNumber === 0) {
      return errorValue(ErrorCode.Div0)
    }
    const value =
      operator === '+'
        ? leftNumber + rightNumber
        : operator === '-'
          ? leftNumber - rightNumber
          : operator === '*'
            ? leftNumber * rightNumber
            : operator === '/'
              ? leftNumber / rightNumber
              : Math.pow(leftNumber, rightNumber)
    return Number.isFinite(value) ? numberValue(value) : errorValue(ErrorCode.Value)
  }
  const comparison = compareInlineScalars(left, right)
  if (comparison === undefined) {
    return errorValue(ErrorCode.Value)
  }
  return booleanValue(
    operator === '='
      ? comparison === 0
      : operator === '<>'
        ? comparison !== 0
        : operator === '>'
          ? comparison > 0
          : operator === '>='
            ? comparison >= 0
            : operator === '<'
              ? comparison < 0
              : comparison <= 0,
  )
}

function evaluateInlineCall(
  state: InlineScalarState,
  rawCallee: string,
  stack: readonly CellValue[],
  start: number,
  argc: number,
): CellValue | undefined {
  const direct = evaluateDirectInlineCall(rawCallee, stack, start, argc)
  if (direct !== undefined) {
    return direct
  }
  const callee = normalizeBuiltinLookupName(rawCallee)
  if (callee !== rawCallee) {
    const normalizedDirect = evaluateDirectInlineCall(callee, stack, start, argc)
    if (normalizedDirect !== undefined) {
      return normalizedDirect
    }
  }
  const dateSystem = state.workbook.getCalculationSettings().dateSystem
  const builtin = dateSystem ? getDateSystemBuiltin(callee, dateSystem) : getBuiltin(callee)
  if (!builtin) {
    return undefined
  }
  const args = stack.slice(start, start + argc)
  const result = builtin(...args)
  return isArrayValue(result) ? undefined : result
}

function evaluateDirectInlineCall(callee: string, stack: readonly CellValue[], start: number, argc: number): CellValue | undefined {
  switch (callee) {
    case 'CONCATENATE':
      return inlineConcatenate(stack, start, argc)
    case 'LEN':
      return argc === 1 ? inlineLen(stack[start]!) : errorValue(ErrorCode.Value)
    case 'MAX':
      return inlineMinMax(stack, start, argc, 'max')
    case 'MIN':
      return inlineMinMax(stack, start, argc, 'min')
    case 'PMT':
      return inlinePmt(stack, start, argc)
    case 'POWER':
      return argc === 2
        ? inlineBinaryNumericResult(stack[start]!, stack[start + 1]!, (left, right) => Math.pow(left, right))
        : errorValue(ErrorCode.Value)
    case 'ROUND':
      return inlineRound(stack, start, argc)
    case 'SQRT':
      return argc === 1 ? inlineUnaryNumericResult(stack[start]!, (value) => Math.sqrt(value)) : errorValue(ErrorCode.Value)
    default:
      return undefined
  }
}

function inlineConcatenate(stack: readonly CellValue[], start: number, count: number): CellValue {
  if (count === 0) {
    return errorValue(ErrorCode.Value)
  }
  let output = ''
  const end = start + count
  for (let index = start; index < end; index += 1) {
    const arg = stack[index]!
    if (arg.tag === ValueTag.Error) {
      return arg
    }
    output += inlineStringValue(arg)
  }
  return stringValue(output)
}

function inlineLen(value: CellValue): CellValue {
  return value.tag === ValueTag.Error ? value : numberValue(inlineStringValue(value).length)
}

function inlineMinMax(stack: readonly CellValue[], start: number, count: number, mode: 'max' | 'min'): CellValue {
  let result = mode === 'max' ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY
  const end = start + count
  for (let index = start; index < end; index += 1) {
    const numeric = inlineNumber(stack[index]!)
    if (numeric === undefined) {
      continue
    }
    result = mode === 'max' ? Math.max(result, numeric) : Math.min(result, numeric)
  }
  return numberValue(result)
}

function inlineUnaryNumericResult(value: CellValue, evaluate: (value: number) => number): CellValue {
  if (value.tag === ValueTag.Error) {
    return value
  }
  const numeric = inlineNumber(value)
  if (numeric === undefined) {
    return errorValue(ErrorCode.Value)
  }
  const result = evaluate(numeric)
  return Number.isFinite(result) ? numberValue(result) : errorValue(ErrorCode.Value)
}

function inlineBinaryNumericResult(left: CellValue, right: CellValue, evaluate: (left: number, right: number) => number): CellValue {
  if (left.tag === ValueTag.Error) {
    return left
  }
  if (right.tag === ValueTag.Error) {
    return right
  }
  const leftNumber = inlineNumber(left)
  const rightNumber = inlineNumber(right)
  if (leftNumber === undefined || rightNumber === undefined) {
    return errorValue(ErrorCode.Value)
  }
  const result = evaluate(leftNumber, rightNumber)
  return Number.isFinite(result) ? numberValue(result) : errorValue(ErrorCode.Value)
}

function inlineRound(stack: readonly CellValue[], start: number, count: number): CellValue {
  if (count !== 2) {
    return errorValue(ErrorCode.Value)
  }
  return inlineBinaryNumericResult(stack[start]!, stack[start + 1]!, (value, digits) => roundInlineToDigits(value, Math.trunc(digits)))
}

function inlinePmt(stack: readonly CellValue[], start: number, count: number): CellValue {
  if (count < 3 || count > 5) {
    return errorValue(ErrorCode.Value)
  }
  const rate = inlineRequiredNumber(stack[start]!)
  const periods = inlineRequiredNumber(stack[start + 1]!)
  const present = inlineRequiredNumber(stack[start + 2]!)
  const future = inlineOptionalNumber(count > 3 ? stack[start + 3] : undefined, 0)
  const type = inlinePaymentType(count > 4 ? stack[start + 4] : undefined, 0)
  if (rate === undefined || periods === undefined || present === undefined || future === undefined || type === undefined || periods <= 0) {
    return errorValue(ErrorCode.Value)
  }
  if (rate === 0) {
    return numberValue(-(future + present) / periods)
  }
  const growth = (1 + rate) ** periods
  const denominator = (1 + rate * type) * (growth - 1)
  if (denominator === 0) {
    return errorValue(ErrorCode.Value)
  }
  return numberValue((-rate * (future + present * growth)) / denominator)
}
