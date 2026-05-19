import { parseNumericText, type JsPlanInstruction } from '@bilig/formula'
import {
  ErrorCode,
  FormulaMode,
  ValueTag,
  formatGeneralNumberValue,
  type CellValue,
  type LiteralInput,
  type RecalcMetrics,
} from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { SheetRecord } from '../../workbook-store.js'
import {
  INLINE_SCALAR_FAST_PLAN_ARITHMETIC,
  INLINE_SCALAR_FAST_PLAN_CONCAT,
  INLINE_SCALAR_FAST_PLAN_IF_STRING,
  INLINE_SCALAR_FAST_PLAN_LEN,
  INLINE_SCALAR_FAST_PLAN_MIN_MAX,
  INLINE_SCALAR_FAST_PLAN_PMT,
  INLINE_SCALAR_FAST_PLAN_ROUND_SQRT,
  type RuntimeFormula,
} from '../runtime-state.js'
import { cellValuesEqual } from './formula-evaluation-helpers.js'
import { tryEvaluateFormulaLeafInlineScalar } from './formula-leaf-inline-scalar-evaluator.js'
import { tryEvaluateFormulaLeafStaticAggregateExpression } from './formula-leaf-static-aggregate-expression.js'
import { tryEvaluateFormulaLeafStaticTableLookup } from './formula-leaf-static-table-lookup.js'
import { makeCompactExistingNumericMutationResult } from './operation-change-helpers.js'
import { emitOperationTrackedCellsBatch } from './operation-tracked-event-helpers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

const NUMBER_VALUE_TAG = ValueTag.Number
const WASM_FAST_PATH_FORMULA_MODE = FormulaMode.WasmFastPath
const INVALID_INLINE_CELL_INDEX = 0xffffffff
const INVALID_INLINE_STRING_ID = 0xffffffff

interface FormulaLeafEvaluationResult {
  readonly mode: 'inline' | 'js' | 'wasm'
  readonly changed: boolean
  readonly value?: CellValue
  readonly numericValue?: number
}

export interface OperationTrustedFormulaLeafExistingNumericMutationRequest {
  readonly existingIndex: number
  readonly formulaCellIndex: number
  readonly sheet: SheetRecord
  readonly col: number
  readonly value: number
  readonly oldNumber: number
  readonly hasTrackedEventListeners: boolean
}

export interface OperationFormulaLeafExistingLiteralMutationRequest {
  readonly existingIndex: number
  readonly formulaCellIndex: number
  readonly value: LiteralInput
  readonly hasTrackedEventListeners: boolean
}

export interface OperationFormulaLeafExistingNumericFastPathArgs {
  readonly state: Pick<
    CreateEngineOperationServiceArgs['state'],
    'workbook' | 'strings' | 'wasm' | 'formulas' | 'counters' | 'events' | 'setLastMetrics'
  >
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly writeTrustedExistingNumericLiteralToCell: (existingIndex: number, sheet: SheetRecord, col: number, value: number) => void
  readonly evaluateFormulaCell: (formulaCellIndex: number) => readonly number[]
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export interface OperationFormulaLeafExistingLiteralFastPathArgs {
  readonly state: OperationFormulaLeafExistingNumericFastPathArgs['state']
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly writeFastPathLiteralToExistingCell: (existingIndex: number, value: LiteralInput) => void
  readonly evaluateFormulaCell: (formulaCellIndex: number) => readonly number[]
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export function tryApplyTrustedFormulaLeafExistingNumericMutation(
  args: OperationFormulaLeafExistingNumericFastPathArgs,
  request: OperationTrustedFormulaLeafExistingNumericMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const formula = getApplicableFormulaLeaf(args, request.formulaCellIndex)
  if (!formula) {
    return null
  }
  const inlineDeltaResult = tryApplyTrustedFormulaLeafInlineNumericDelta(args, request, formula)
  if (inlineDeltaResult !== undefined) {
    return inlineDeltaResult
  }

  const cellStore = args.state.workbook.cellStore
  args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
  const evaluation = evaluateFormulaLeafAfterInputWrite(args, request.existingIndex, request.formulaCellIndex, formula)
  const afterFormulaNumber = formulaLeafEvaluationNumericValue(evaluation)
  const afterFormulaValue = formulaLeafEvaluationPublicValue(evaluation)
  const formulaChanged = evaluation.changed
  addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics =
    evaluation.mode === 'inline'
      ? args.makeSingleLiteralSkipMetrics()
      : {
          ...args.makeSingleLiteralSkipMetrics(),
          wasmFormulaCount: evaluation.mode === 'wasm' ? 1 : 0,
          jsFormulaCount: evaluation.mode === 'js' ? 1 : 0,
        }
  args.state.setLastMetrics(lastMetrics)

  const result = formulaChanged
    ? makeCompactExistingNumericMutationResult(
        request.existingIndex,
        request.formulaCellIndex,
        1,
        afterFormulaNumber,
        cellStore.rows[request.formulaCellIndex] ?? 0,
        cellStore.cols[request.formulaCellIndex] ?? 0,
        afterFormulaValue,
      )
    : makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
  if (request.hasTrackedEventListeners) {
    const changedCellIndices = formulaChanged
      ? Uint32Array.of(request.existingIndex, request.formulaCellIndex)
      : Uint32Array.of(request.existingIndex)
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices,
      metrics: lastMetrics,
    })
  }
  return result
}

function tryApplyTrustedFormulaLeafInlineNumericDelta(
  args: OperationFormulaLeafExistingNumericFastPathArgs,
  request: OperationTrustedFormulaLeafExistingNumericMutationRequest,
  formula: RuntimeFormula,
): EngineExistingNumericCellMutationResult | undefined {
  const coefficient = inlineNumericDeltaCoefficient(formula, request.existingIndex)
  if (coefficient === undefined) {
    return undefined
  }
  const cellStore = args.state.workbook.cellStore
  if (cellStore.tags[request.formulaCellIndex] !== NUMBER_VALUE_TAG) {
    return undefined
  }
  const oldFormulaNumber = cellStore.numbers[request.formulaCellIndex] ?? 0
  const inputDelta = request.value - request.oldNumber
  const formulaDelta = inputDelta * coefficient
  const nextFormulaNumber = oldFormulaNumber + formulaDelta
  if (!Number.isFinite(nextFormulaNumber)) {
    return undefined
  }

  args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
  const formulaChanged = writeInlineFormulaLeafNumber(args, request.formulaCellIndex, nextFormulaNumber)
  addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = args.makeSingleLiteralSkipMetrics()
  args.state.setLastMetrics(lastMetrics)
  if (request.hasTrackedEventListeners) {
    const changedCellIndices = formulaChanged
      ? Uint32Array.of(request.existingIndex, request.formulaCellIndex)
      : Uint32Array.of(request.existingIndex)
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices,
      metrics: lastMetrics,
    })
  }
  return formulaChanged
    ? makeCompactExistingNumericMutationResult(
        request.existingIndex,
        request.formulaCellIndex,
        1,
        nextFormulaNumber,
        cellStore.rows[request.formulaCellIndex] ?? 0,
        cellStore.cols[request.formulaCellIndex] ?? 0,
      )
    : makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
}

export function tryApplyFormulaLeafExistingLiteralMutation(
  args: OperationFormulaLeafExistingLiteralFastPathArgs,
  request: OperationFormulaLeafExistingLiteralMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const formula = getApplicableFormulaLeaf(args, request.formulaCellIndex)
  if (!formula) {
    return null
  }

  args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
  const evaluation = evaluateFormulaLeafAfterInputWrite(args, request.existingIndex, request.formulaCellIndex, formula)
  const afterFormulaNumber = formulaLeafEvaluationNumericValue(evaluation)
  const afterFormulaValue = formulaLeafEvaluationPublicValue(evaluation)
  const formulaChanged = evaluation.changed
  addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics =
    evaluation.mode === 'inline'
      ? args.makeSingleLiteralSkipMetrics()
      : {
          ...args.makeSingleLiteralSkipMetrics(),
          wasmFormulaCount: evaluation.mode === 'wasm' ? 1 : 0,
          jsFormulaCount: evaluation.mode === 'js' ? 1 : 0,
        }
  args.state.setLastMetrics(lastMetrics)
  if (request.hasTrackedEventListeners) {
    const changedCellIndices = formulaChanged
      ? Uint32Array.of(request.existingIndex, request.formulaCellIndex)
      : Uint32Array.of(request.existingIndex)
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices,
      metrics: lastMetrics,
    })
  }
  const cellStore = args.state.workbook.cellStore
  return formulaChanged
    ? makeCompactExistingNumericMutationResult(
        request.existingIndex,
        request.formulaCellIndex,
        1,
        afterFormulaNumber,
        cellStore.rows[request.formulaCellIndex] ?? 0,
        cellStore.cols[request.formulaCellIndex] ?? 0,
        afterFormulaValue,
      )
    : makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
}

function getApplicableFormulaLeaf(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state' | 'getSingleEntityDependent'>,
  formulaCellIndex: number,
): RuntimeFormula | undefined {
  if (formulaCellIndex < 0 || args.getSingleEntityDependent(makeCellEntity(formulaCellIndex)) !== -1) {
    return undefined
  }
  const formula = args.state.formulas.get(formulaCellIndex)
  if (
    !formula ||
    formula.directLookup !== undefined ||
    formula.directAggregate !== undefined ||
    formula.directCriteria !== undefined ||
    formula.directScalar !== undefined ||
    formula.compiled.volatile ||
    formula.compiled.producesSpill ||
    ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
  ) {
    return undefined
  }
  return formula
}

function evaluateFormulaLeafAfterInputWrite(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state' | 'evaluateFormulaCell'>,
  existingIndex: number,
  formulaCellIndex: number,
  formula: RuntimeFormula,
): FormulaLeafEvaluationResult {
  const storedFastPlan = tryStoreFormulaLeafInlineFastPlan(args, formulaCellIndex, formula)
  if (storedFastPlan !== undefined) {
    return storedFastPlan
  }

  const staticTableLookupResult = tryEvaluateFormulaLeafStaticTableLookup({
    state: args.state,
    formula,
  })
  if (staticTableLookupResult !== undefined) {
    const stored = storeInlineFormulaLeafValue(args, formulaCellIndex, staticTableLookupResult)
    return { mode: 'inline', value: stored.value, changed: stored.changed }
  }

  const staticAggregateExpressionResult = tryEvaluateFormulaLeafStaticAggregateExpression({
    state: args.state,
    formula,
  })
  if (staticAggregateExpressionResult !== undefined) {
    const stored = storeInlineFormulaLeafValue(args, formulaCellIndex, staticAggregateExpressionResult)
    return { mode: 'inline', value: stored.value, changed: stored.changed }
  }

  const inlineResult = tryEvaluateFormulaLeafInlineScalar({
    state: args.state,
    formula,
  })
  if (inlineResult !== undefined) {
    if (inlineResult.tag === NUMBER_VALUE_TAG) {
      const changed = writeInlineFormulaLeafNumber(args, formulaCellIndex, inlineResult.value)
      return { mode: 'inline', numericValue: inlineResult.value, changed }
    }
    const stored = storeInlineFormulaLeafValue(args, formulaCellIndex, inlineResult)
    return { mode: 'inline', value: stored.value, changed: stored.changed }
  }

  const beforeFormulaValue = readFormulaCellValue(args, formulaCellIndex)
  const evaluatedByWasm = tryEvaluateLeafFormulaWithWasm(args, existingIndex, formulaCellIndex, formula)
  if (evaluatedByWasm) {
    const afterFormulaValue = readFormulaCellValue(args, formulaCellIndex)
    return { mode: 'wasm', value: afterFormulaValue, changed: !cellValuesEqual(beforeFormulaValue, afterFormulaValue) }
  }
  args.evaluateFormulaCell(formulaCellIndex)
  const afterFormulaValue = readFormulaCellValue(args, formulaCellIndex)
  return { mode: 'js', value: afterFormulaValue, changed: !cellValuesEqual(beforeFormulaValue, afterFormulaValue) }
}

function formulaLeafEvaluationNumericValue(evaluation: FormulaLeafEvaluationResult): number | undefined {
  if (evaluation.numericValue !== undefined) {
    return evaluation.numericValue
  }
  return evaluation.value?.tag === NUMBER_VALUE_TAG ? evaluation.value.value : undefined
}

function formulaLeafEvaluationPublicValue(evaluation: FormulaLeafEvaluationResult): CellValue | undefined {
  const value = evaluation.value
  return value?.tag === NUMBER_VALUE_TAG ? undefined : value
}

function tryStoreFormulaLeafInlineFastPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  formula: RuntimeFormula,
): FormulaLeafEvaluationResult | undefined {
  const fastPlanKind = formula.inlineScalarFastPlanKind
  if (fastPlanKind === undefined) {
    return undefined
  }
  const plan = formula.compiled.jsPlan
  const cellIndices = formula.inlineScalarPlanCellIndices
  if (cellIndices === undefined || cellIndices.length !== plan.length) {
    return undefined
  }

  switch (fastPlanKind) {
    case INLINE_SCALAR_FAST_PLAN_ARITHMETIC:
      return storeFastArithmeticPlan(args, formulaCellIndex, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_LEN:
      return storeFastLenPlan(args, formulaCellIndex, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_CONCAT:
      return storeFastConcatPlan(args, formulaCellIndex, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_MIN_MAX:
      return storeFastMinMaxPlan(args, formulaCellIndex, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_ROUND_SQRT:
      return storeFastRoundSqrtPlan(args, formulaCellIndex, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_PMT:
      return storeFastPmtPlan(args, formulaCellIndex, plan, cellIndices)
    case INLINE_SCALAR_FAST_PLAN_IF_STRING:
      return storeFastIfStringPlan(args, formulaCellIndex, plan, cellIndices, formula.inlineScalarFastPlanStringIds)
    default:
      return undefined
  }
}

function storeFastArithmeticPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): FormulaLeafEvaluationResult | undefined {
  const literal = inlinePushNumberValue(plan[2])
  const innerOperator = inlineBinaryOperator(plan[3])
  const outerOperator = inlineBinaryOperator(plan[4])
  if (literal === undefined || innerOperator === undefined || outerOperator === undefined) {
    return undefined
  }
  const left = inlineArithmeticNumberFromStoredCell(args.state, cellIndices[0])
  const rightLeft = inlineArithmeticNumberFromStoredCell(args.state, cellIndices[1])
  if (left === undefined || rightLeft === undefined) {
    return undefined
  }
  const right = evaluateFastInlineNumericBinary(innerOperator, rightLeft, literal)
  if (right === undefined) {
    return undefined
  }
  const value = evaluateFastInlineNumericBinary(outerOperator, left, right)
  return value === undefined ? undefined : storeInlineFormulaLeafNumberResult(args, formulaCellIndex, value)
}

function storeFastLenPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  cellIndices: Uint32Array,
): FormulaLeafEvaluationResult | undefined {
  const leftLength = inlineStringLengthFromStoredCell(args.state, cellIndices[0])
  const rightLength = inlineStringLengthFromStoredCell(args.state, cellIndices[2])
  return leftLength === undefined || rightLength === undefined
    ? undefined
    : storeInlineFormulaLeafNumberResult(args, formulaCellIndex, leftLength + rightLength)
}

function storeFastConcatPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): FormulaLeafEvaluationResult | undefined {
  const separator = inlinePushStringValue(plan[1])
  const left = inlineStringFromStoredCell(args.state, cellIndices[0])
  const right = inlineStringFromStoredCell(args.state, cellIndices[2])
  if (left === undefined || right === undefined || separator === undefined) {
    return undefined
  }
  return storeInlineFormulaLeafStringResult(args, formulaCellIndex, `${left}${separator}${right}`)
}

function storeFastMinMaxPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  cellIndices: Uint32Array,
): FormulaLeafEvaluationResult | undefined {
  const first = inlineNumberFromStoredCell(args.state, cellIndices[0])
  const second = inlineNumberFromStoredCell(args.state, cellIndices[1])
  const third = inlineNumberFromStoredCell(args.state, cellIndices[2])
  if (first === undefined || second === undefined || third === undefined) {
    return undefined
  }
  return storeInlineFormulaLeafNumberResult(args, formulaCellIndex, Math.min(first, second, third) + Math.max(first, second, third))
}

function storeFastRoundSqrtPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): FormulaLeafEvaluationResult | undefined {
  const digits = inlinePushNumberValue(plan[2])
  const value = inlineNumberFromStoredCell(args.state, cellIndices[0])
  if (value === undefined || digits === undefined) {
    return undefined
  }
  const sqrt = Math.sqrt(value)
  return Number.isFinite(sqrt)
    ? storeInlineFormulaLeafNumberResult(args, formulaCellIndex, roundInlineToDigits(sqrt, Math.trunc(digits)))
    : storeInlineFormulaLeafErrorResult(args, formulaCellIndex, ErrorCode.Value)
}

function storeFastPmtPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
): FormulaLeafEvaluationResult | undefined {
  const divisor = inlinePushNumberValue(plan[1])
  const annualRate = inlineNumberFromStoredCell(args.state, cellIndices[0])
  const periods = inlineNumberFromStoredCell(args.state, cellIndices[3])
  const present = inlineNumberFromStoredCell(args.state, cellIndices[4])
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
    return storeInlineFormulaLeafNumberResult(args, formulaCellIndex, -present / periods)
  }
  const growth = (1 + rate) ** periods
  const denominator = growth - 1
  return denominator === 0
    ? storeInlineFormulaLeafErrorResult(args, formulaCellIndex, ErrorCode.Value)
    : storeInlineFormulaLeafNumberResult(args, formulaCellIndex, (-rate * present * growth) / denominator)
}

function storeFastIfStringPlan(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  plan: readonly JsPlanInstruction[],
  cellIndices: Uint32Array,
  stringIds: Uint32Array | undefined,
): FormulaLeafEvaluationResult | undefined {
  const right = inlinePushNumberValue(plan[1])
  const operator = inlineBinaryOperator(plan[2])
  const trueValue = inlinePushStringValue(plan[4])
  const falseValue = inlinePushStringValue(plan[6])
  const left = inlineNumberFromStoredCell(args.state, cellIndices[0])
  if (left === undefined || right === undefined || operator === undefined || trueValue === undefined || falseValue === undefined) {
    return undefined
  }
  const comparison = compareFastInlineNumbers(operator, left, right)
  if (comparison === undefined) {
    return undefined
  }
  const selectedValue = comparison ? trueValue : falseValue
  const selectedStringId = stringIds?.[comparison ? 4 : 6]
  return selectedStringId === undefined || selectedStringId === INVALID_INLINE_STRING_ID
    ? storeInlineFormulaLeafStringResult(args, formulaCellIndex, selectedValue)
    : storeInlineFormulaLeafKnownStringResult(args, formulaCellIndex, selectedValue, selectedStringId)
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

function inlineNumberFromStoredCell(
  state: OperationFormulaLeafExistingNumericFastPathArgs['state'],
  cellIndex: number | undefined,
): number | undefined {
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

function inlineArithmeticNumberFromStoredCell(
  state: OperationFormulaLeafExistingNumericFastPathArgs['state'],
  cellIndex: number | undefined,
): number | undefined {
  if (cellIndex === undefined || cellIndex === INVALID_INLINE_CELL_INDEX) {
    return undefined
  }
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] ?? ValueTag.Empty) as ValueTag
  if (tag !== ValueTag.String) {
    return inlineNumberFromStoredCell(state, cellIndex)
  }
  const trimmed = state.strings.get(cellStore.stringIds[cellIndex] ?? 0).trim()
  return trimmed === '' ? 0 : parseNumericText(trimmed)
}

function inlineStringLengthFromStoredCell(
  state: OperationFormulaLeafExistingNumericFastPathArgs['state'],
  cellIndex: number | undefined,
): number | undefined {
  const value = inlineStringFromStoredCell(state, cellIndex)
  return value === undefined ? undefined : value.length
}

function inlineStringFromStoredCell(
  state: OperationFormulaLeafExistingNumericFastPathArgs['state'],
  cellIndex: number | undefined,
): string | undefined {
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

function roundInlineToDigits(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits
    return Math.round(value * factor) / factor
  }
  const factor = 10 ** -digits
  return Math.round(value / factor) * factor
}

function storeInlineFormulaLeafNumberResult(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: number,
): FormulaLeafEvaluationResult {
  const changed = writeInlineFormulaLeafNumber(args, formulaCellIndex, value)
  return { mode: 'inline', numericValue: value, changed }
}

function storeInlineFormulaLeafStringResult(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: string,
): FormulaLeafEvaluationResult {
  const stored = writeInlineFormulaLeafString(args, formulaCellIndex, value)
  return { mode: 'inline', value: stored.value, changed: stored.changed }
}

function storeInlineFormulaLeafKnownStringResult(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: string,
  stringId: number,
): FormulaLeafEvaluationResult {
  const stored = writeInlineFormulaLeafKnownString(args, formulaCellIndex, value, stringId)
  return { mode: 'inline', value: stored.value, changed: stored.changed }
}

function storeInlineFormulaLeafErrorResult(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  code: ErrorCode,
): FormulaLeafEvaluationResult {
  const stored = writeInlineFormulaLeafError(args, formulaCellIndex, code)
  return { mode: 'inline', value: stored.value, changed: stored.changed }
}

function tryEvaluateLeafFormulaWithWasm(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  existingIndex: number,
  formulaCellIndex: number,
  formula: RuntimeFormula,
): boolean {
  if (formula.compiled.mode !== WASM_FAST_PATH_FORMULA_MODE || !args.state.wasm.ready) {
    return false
  }
  args.state.wasm.syncStringPool(args.state.strings.exportLayout())
  args.state.wasm.syncFromStore(
    args.state.workbook.cellStore,
    formula.dependencyIndices.length > 0 ? formula.dependencyIndices : Uint32Array.of(existingIndex),
  )
  const formulaIndices = Uint32Array.of(formulaCellIndex)
  args.state.wasm.evalBatch(formulaIndices)
  args.state.wasm.syncToStore(args.state.workbook.cellStore, formulaIndices, args.state.strings, (changedCellIndex: number) =>
    args.state.workbook.notifyCellValueWritten(changedCellIndex),
  )
  return true
}

function readFormulaCellValue(args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>, formulaCellIndex: number): CellValue {
  return args.state.workbook.cellStore.getValue(formulaCellIndex, (stringId) => (stringId === 0 ? '' : args.state.strings.get(stringId)))
}

function inlineNumericDeltaCoefficient(formula: RuntimeFormula, changedCellIndex: number): number | undefined {
  if (formula.inlineScalarFastPlanKind !== INLINE_SCALAR_FAST_PLAN_ARITHMETIC) {
    return undefined
  }
  const plan = formula.compiled.jsPlan
  const cellIndices = formula.inlineScalarPlanCellIndices
  if (cellIndices === undefined) {
    return undefined
  }
  const coefficients = formula.inlineScalarArithmeticDeltaCoefficients
  if (coefficients !== undefined) {
    for (let index = 0; index < coefficients.length; index += 1) {
      if (cellIndices[index] === changedCellIndex) {
        return coefficients[index]
      }
    }
    return undefined
  }
  const outerOperator = inlineBinaryOperator(plan[4])
  if (outerOperator === undefined) {
    return undefined
  }
  if (outerOperator !== '+' && outerOperator !== '-') {
    return undefined
  }
  if (cellIndices[0] === changedCellIndex) {
    return 1
  }
  if (cellIndices[1] !== changedCellIndex) {
    return undefined
  }
  const literal = inlinePushNumberValue(plan[2])
  const innerOperator = inlineBinaryOperator(plan[3])
  if (literal === undefined || innerOperator === undefined) {
    return undefined
  }
  const innerCoefficient = innerOperator === '*' ? literal : innerOperator === '/' && literal !== 0 ? 1 / literal : undefined
  if (innerCoefficient === undefined || !Number.isFinite(innerCoefficient)) {
    return undefined
  }
  return outerOperator === '-' ? -innerCoefficient : innerCoefficient
}

function writeInlineFormulaLeafNumber(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: number,
): boolean {
  const cellStore = args.state.workbook.cellStore
  const changed = cellStore.tags[formulaCellIndex] !== NUMBER_VALUE_TAG || cellStore.numbers[formulaCellIndex] !== value
  cellStore.flags[formulaCellIndex] = (cellStore.flags[formulaCellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.tags[formulaCellIndex] = NUMBER_VALUE_TAG
  cellStore.errors[formulaCellIndex] = ErrorCode.None
  cellStore.stringIds[formulaCellIndex] = 0
  cellStore.numbers[formulaCellIndex] = value
  cellStore.versions[formulaCellIndex] = (cellStore.versions[formulaCellIndex] ?? 0) + 1
  notifyInlineFormulaValueStored(args, formulaCellIndex, changed)
  return changed
}

function storeInlineFormulaLeafValue(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: CellValue,
): { readonly value: CellValue; readonly changed: boolean } {
  if (value.tag === NUMBER_VALUE_TAG) {
    return { value, changed: writeInlineFormulaLeafNumber(args, formulaCellIndex, value.value) }
  }
  const cellStore = args.state.workbook.cellStore
  const existingTag = (cellStore.tags[formulaCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
  let stringId = 0
  let number = 0
  let error = ErrorCode.None
  let storedValue = value
  let changed = existingTag !== value.tag
  switch (value.tag) {
    case ValueTag.Empty:
      break
    case ValueTag.Boolean:
      number = value.value ? 1 : 0
      changed ||= cellStore.numbers[formulaCellIndex] !== number
      break
    case ValueTag.String:
      stringId = args.state.strings.intern(value.value)
      storedValue = { tag: ValueTag.String, value: value.value, stringId }
      changed ||= cellStore.stringIds[formulaCellIndex] !== stringId
      break
    case ValueTag.Error:
      error = value.code
      changed ||= cellStore.errors[formulaCellIndex] !== error
      break
  }
  cellStore.flags[formulaCellIndex] = (cellStore.flags[formulaCellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.tags[formulaCellIndex] = storedValue.tag
  cellStore.errors[formulaCellIndex] = error
  cellStore.stringIds[formulaCellIndex] = stringId
  cellStore.numbers[formulaCellIndex] = number
  cellStore.versions[formulaCellIndex] = (cellStore.versions[formulaCellIndex] ?? 0) + 1
  notifyInlineFormulaValueStored(args, formulaCellIndex, changed)
  return { value: storedValue, changed }
}

function writeInlineFormulaLeafString(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: string,
): { readonly value: CellValue; readonly changed: boolean } {
  const stringId = args.state.strings.intern(value)
  return writeInlineFormulaLeafKnownString(args, formulaCellIndex, value, stringId)
}

function writeInlineFormulaLeafKnownString(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  value: string,
  stringId: number,
): { readonly value: CellValue; readonly changed: boolean } {
  const cellStore = args.state.workbook.cellStore
  const changed = cellStore.tags[formulaCellIndex] !== ValueTag.String || cellStore.stringIds[formulaCellIndex] !== stringId
  cellStore.flags[formulaCellIndex] = (cellStore.flags[formulaCellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.tags[formulaCellIndex] = ValueTag.String
  cellStore.errors[formulaCellIndex] = ErrorCode.None
  cellStore.stringIds[formulaCellIndex] = stringId
  cellStore.numbers[formulaCellIndex] = 0
  cellStore.versions[formulaCellIndex] = (cellStore.versions[formulaCellIndex] ?? 0) + 1
  notifyInlineFormulaValueStored(args, formulaCellIndex, changed)
  return { value: { tag: ValueTag.String, value, stringId }, changed }
}

function writeInlineFormulaLeafError(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  code: ErrorCode,
): { readonly value: CellValue; readonly changed: boolean } {
  const cellStore = args.state.workbook.cellStore
  const changed = cellStore.tags[formulaCellIndex] !== ValueTag.Error || cellStore.errors[formulaCellIndex] !== code
  cellStore.flags[formulaCellIndex] = (cellStore.flags[formulaCellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.tags[formulaCellIndex] = ValueTag.Error
  cellStore.errors[formulaCellIndex] = code
  cellStore.stringIds[formulaCellIndex] = 0
  cellStore.numbers[formulaCellIndex] = 0
  cellStore.versions[formulaCellIndex] = (cellStore.versions[formulaCellIndex] ?? 0) + 1
  notifyInlineFormulaValueStored(args, formulaCellIndex, changed)
  return { value: { tag: ValueTag.Error, code }, changed }
}

function notifyInlineFormulaValueStored(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
  changed: boolean,
): void {
  const onSetValue = args.state.workbook.cellStore.onSetValue
  if (onSetValue) {
    onSetValue(formulaCellIndex)
    return
  }
  if (changed) {
    notifyPhysicalInlineFormulaValueWritten(args, formulaCellIndex)
  }
}

function notifyPhysicalInlineFormulaValueWritten(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  formulaCellIndex: number,
): void {
  const cellStore = args.state.workbook.cellStore
  const sheet = args.state.workbook.getSheetById(cellStore.sheetIds[formulaCellIndex]!)
  const col = cellStore.cols[formulaCellIndex]!
  if (sheet?.structureVersion === 1 && col < sheet.columnVersions.length) {
    sheet.columnVersions[col] = (sheet.columnVersions[col] ?? 0) + 1
    return
  }
  args.state.workbook.notifyCellValueWritten(formulaCellIndex)
}
