import { ErrorCode, FormulaMode, ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'
import type { RuntimeDirectCriteriaDescriptor, RuntimeDirectCriteriaOperand, RuntimeDirectScalarDescriptor, U32 } from '../runtime-state.js'
import type { DirectFormulaIndexCollection, DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import { mergeChangedCellIndices } from './operation-change-helpers.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

export interface DirectFormulaMetricCounts {
  wasmFormulaCount: number
  jsFormulaCount: number
}

export interface OperationPostRecalcFormula {
  readonly compiled: {
    readonly producesSpill: boolean
    readonly mode?: FormulaMode
  }
  readonly directAggregate: object | undefined
  readonly directCriteria: RuntimeDirectCriteriaDescriptor | undefined
  readonly directScalar: RuntimeDirectScalarDescriptor | undefined
}

export interface OperationPostRecalcDirectFormulaState {
  readonly workbook: {
    readonly cellStore: {
      readonly flags: {
        readonly [index: number]: number | undefined
      }
      readonly tags: ArrayLike<ValueTag | undefined>
      readonly numbers: ArrayLike<number | undefined>
      readonly errors: ArrayLike<ErrorCode | undefined>
    }
    readonly withBatchedColumnVersionUpdates: (apply: () => void) => void
  }
  readonly formulas: {
    readonly get: (cellIndex: number) => OperationPostRecalcFormula | undefined
  }
  readonly counters: EngineCounters
}

export interface ApplyPostRecalcDirectFormulaChangesArgs {
  readonly state: OperationPostRecalcDirectFormulaState
  readonly collection: DirectFormulaIndexCollection
  readonly recalculated: U32
  readonly didRunRecalc: boolean
  readonly captureChanged?: boolean
  readonly metrics: DirectFormulaMetricCounts
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, result: DirectScalarCurrentOperand) => boolean
  readonly applyDirectFormulaNumericDelta: (cellIndex: number, delta: number) => boolean
  readonly applyDirectScalarCurrentValue: (cellIndex: number, directScalar: RuntimeDirectScalarDescriptor) => boolean
  readonly tryApplyDirectScalarDeltas: (collection: DirectFormulaIndexCollection, captureChanged?: boolean) => U32 | undefined
  readonly tryApplyDirectFormulaDeltas: (collection: DirectFormulaIndexCollection, captureChanged?: boolean) => U32 | undefined
  readonly countPostRecalcDirectFormulaMetric: (cellIndex: number, counts: DirectFormulaMetricCounts) => void
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
}

export function applyPostRecalcDirectFormulaChanges(args: ApplyPostRecalcDirectFormulaChangesArgs): U32 {
  if (args.collection.size === 0) {
    return args.recalculated
  }
  const captureChanged = args.captureChanged ?? true
  const constantScalarChanged = !args.didRunRecalc ? args.tryApplyDirectScalarDeltas(args.collection, captureChanged) : undefined
  if (constantScalarChanged !== undefined) {
    return mergeChangedCellIndices(args.recalculated, constantScalarChanged)
  }
  const directDeltaChanged = !args.didRunRecalc ? args.tryApplyDirectFormulaDeltas(args.collection, captureChanged) : undefined
  if (directDeltaChanged !== undefined) {
    return mergeChangedCellIndices(args.recalculated, directDeltaChanged)
  }
  const singleDirectChanged = tryApplySinglePostRecalcDirectFormula(args, captureChanged)
  if (singleDirectChanged !== undefined) {
    return mergeChangedCellIndices(args.recalculated, singleDirectChanged)
  }
  return applyDirectFormulaFallbacks(args, captureChanged)
}

export function countOperationPostRecalcDirectFormulaMetric(input: {
  readonly formulas: {
    readonly get: (cellIndex: number) => OperationPostRecalcFormula | undefined
  }
  readonly cellIndex: number
  readonly counts: DirectFormulaMetricCounts
}): void {
  const formula = input.formulas.get(input.cellIndex)
  if (!formula || (formula.directScalar === undefined && formula.directAggregate === undefined && formula.directCriteria === undefined)) {
    return
  }
  if (formula.compiled.mode === FormulaMode.WasmFastPath) {
    input.counts.wasmFormulaCount += 1
    return
  }
  input.counts.jsFormulaCount += 1
}

export function tryApplySinglePostRecalcDirectFormula(
  args: ApplyPostRecalcDirectFormulaChangesArgs,
  captureChanged = true,
): U32 | undefined {
  if (args.didRunRecalc || args.collection.size !== 1) {
    return undefined
  }
  const cellIndex = args.collection.getCellIndexAt(0)
  const formula = args.state.formulas.get(cellIndex)
  if (!formula) {
    return undefined
  }
  if (isCycleFormulaCell(args, cellIndex)) {
    return undefined
  }
  const currentResult = args.collection.getCurrentResultAt(0)
  if (currentResult !== undefined) {
    return args.applyDirectFormulaCurrentResult(cellIndex, currentResult)
      ? captureChanged
        ? Uint32Array.of(cellIndex)
        : EMPTY_CHANGED_CELLS
      : undefined
  }
  const delta = args.collection.getDeltaAt(0)
  if (delta !== undefined) {
    if (!args.applyDirectFormulaNumericDelta(cellIndex, delta)) {
      return undefined
    }
    countDirectDeltaApplication(args, args.state.formulas.get(cellIndex))
    return captureChanged ? Uint32Array.of(cellIndex) : EMPTY_CHANGED_CELLS
  }
  if (
    formula?.directScalar !== undefined &&
    !formula.compiled.producesSpill &&
    args.applyDirectScalarCurrentValue(cellIndex, formula.directScalar)
  ) {
    args.countPostRecalcDirectFormulaMetric(cellIndex, args.metrics)
    return captureChanged ? Uint32Array.of(cellIndex) : EMPTY_CHANGED_CELLS
  }
  return undefined
}

function applyDirectFormulaFallbacks(args: ApplyPostRecalcDirectFormulaChangesArgs, captureChanged: boolean): U32 {
  const postRecalcChanged = captureChanged ? new Uint32Array(args.collection.size) : EMPTY_CHANGED_CELLS
  let postRecalcChangedCount = 0
  let postRecalcExtraChanged: number[] | undefined
  let directAggregateDeltaApplicationCount = 0
  let directScalarDeltaApplicationCount = 0
  const directCriteriaCurrentResults = !args.didRunRecalc ? new Map<string, DirectScalarCurrentOperand>() : undefined
  args.state.workbook.withBatchedColumnVersionUpdates(() => {
    args.collection.forEachIndexed((cellIndex, directIndex) => {
      if (isCycleFormulaCell(args, cellIndex)) {
        return
      }
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      const currentResult = args.collection.getCurrentResultAt(directIndex)
      if (!args.didRunRecalc && currentResult !== undefined && args.applyDirectFormulaCurrentResult(cellIndex, currentResult)) {
        if (captureChanged) {
          postRecalcChanged[postRecalcChangedCount++] = cellIndex
        }
        return
      }
      const delta = args.collection.getDeltaAt(directIndex)
      if (!args.didRunRecalc && delta !== undefined && args.applyDirectFormulaNumericDelta(cellIndex, delta)) {
        if (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) {
          directAggregateDeltaApplicationCount += 1
        }
        if (formula?.directScalar !== undefined) {
          directScalarDeltaApplicationCount += 1
        }
        if (captureChanged) {
          postRecalcChanged[postRecalcChangedCount++] = cellIndex
        }
        return
      }
      if (
        !args.didRunRecalc &&
        formula?.directScalar !== undefined &&
        !formula.compiled.producesSpill &&
        args.applyDirectScalarCurrentValue(cellIndex, formula.directScalar)
      ) {
        args.countPostRecalcDirectFormulaMetric(cellIndex, args.metrics)
        if (captureChanged) {
          postRecalcChanged[postRecalcChangedCount++] = cellIndex
        }
        return
      }
      const directCriteriaCurrentResultKey =
        directCriteriaCurrentResults === undefined || formula === undefined ? undefined : directCriteriaCurrentResultCacheKey(formula)
      const cachedDirectCriteriaResult =
        directCriteriaCurrentResultKey === undefined || directCriteriaCurrentResults === undefined
          ? undefined
          : directCriteriaCurrentResults.get(directCriteriaCurrentResultKey)
      if (cachedDirectCriteriaResult !== undefined && args.applyDirectFormulaCurrentResult(cellIndex, cachedDirectCriteriaResult)) {
        if (captureChanged) {
          postRecalcChanged[postRecalcChangedCount++] = cellIndex
        }
        return
      }
      args.countPostRecalcDirectFormulaMetric(cellIndex, args.metrics)
      const changedCellIndices = args.evaluateDirectFormula(cellIndex)
      if (directCriteriaCurrentResultKey !== undefined) {
        const evaluatedResult = readDirectScalarCurrentOperand(args, cellIndex)
        if (evaluatedResult !== undefined) {
          directCriteriaCurrentResults?.set(directCriteriaCurrentResultKey, evaluatedResult)
        }
      }
      if (captureChanged) {
        postRecalcChanged[postRecalcChangedCount++] = cellIndex
      }
      if (captureChanged && changedCellIndices) {
        postRecalcExtraChanged ??= []
        for (let index = 0; index < changedCellIndices.length; index += 1) {
          postRecalcExtraChanged.push(changedCellIndices[index]!)
        }
      }
    })
  })
  if (directAggregateDeltaApplicationCount > 0) {
    addEngineCounter(args.state.counters, 'directAggregateDeltaApplications', directAggregateDeltaApplicationCount)
  }
  if (directScalarDeltaApplicationCount > 0) {
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', directScalarDeltaApplicationCount)
  }
  if (!captureChanged) {
    return EMPTY_CHANGED_CELLS
  }
  const directChanged = postRecalcChanged.subarray(0, postRecalcChangedCount)
  return postRecalcExtraChanged && postRecalcExtraChanged.length > 0
    ? mergeChangedCellIndices(args.recalculated, mergeChangedCellIndices(directChanged, postRecalcExtraChanged))
    : mergeChangedCellIndices(args.recalculated, directChanged)
}

function directCriteriaCurrentResultCacheKey(formula: OperationPostRecalcFormula): string | undefined {
  const directCriteria = formula.directCriteria
  if (
    directCriteria === undefined ||
    formula.compiled.producesSpill ||
    directCriteria.aggregateKind === 'first' ||
    directCriteria.offsetOperand !== undefined ||
    directCriteria.firstMatchMode !== undefined ||
    directCriteria.resultTransforms !== undefined
  ) {
    return undefined
  }
  const aggregateRangeKey = directCriteria.aggregateRange === undefined ? 'none' : directCriteriaRangeKey(directCriteria.aggregateRange)
  const criteriaPairsKey = directCriteria.criteriaPairs
    .map((pair) => `${directCriteriaRangeKey(pair.range)}=${directCriteriaOperandKey(pair.criterion)}`)
    .join('\u0002')
  return `${directCriteria.aggregateKind}\u0001${aggregateRangeKey}\u0001${criteriaPairsKey}`
}

function directCriteriaRangeKey(range: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['range']): string {
  return `${range.sheetName}:${range.rowStart}:${range.rowEnd}:${range.col}:${range.length}`
}

function directCriteriaOperandKey(operand: RuntimeDirectCriteriaOperand): string {
  switch (operand.kind) {
    case 'literal':
      return `literal:${cellValueKey(operand.value)}`
    case 'cell':
      return `cell:${operand.cellIndex}`
    case 'cell-string-concat':
      return `cell-string-concat:${operand.cellIndex}:${operand.prefix}:${operand.suffix}`
    case 'cell-month-boundary-string-concat':
      return `cell-month-boundary-string-concat:${operand.cellIndex}:${operand.prefix}:${operand.suffix}:${operand.offsetMonths}`
  }
}

function cellValueKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'empty'
    case ValueTag.Number:
      return `number:${Object.is(value.value, -0) ? 0 : value.value}`
    case ValueTag.Boolean:
      return `boolean:${value.value ? 1 : 0}`
    case ValueTag.String:
      return `string:${value.stringId ?? 0}:${value.value}`
    case ValueTag.Error:
      return `error:${value.code}`
  }
}

function readDirectScalarCurrentOperand(
  args: Pick<ApplyPostRecalcDirectFormulaChangesArgs, 'state'>,
  cellIndex: number,
): DirectScalarCurrentOperand | undefined {
  const cellStore = args.state.workbook.cellStore
  const tag = cellStore.tags[cellIndex]
  if (tag === ValueTag.Number) {
    return { kind: 'number', value: cellStore.numbers[cellIndex] ?? 0 }
  }
  if (tag === ValueTag.Error) {
    return { kind: 'error', code: cellStore.errors[cellIndex] ?? ErrorCode.None }
  }
  return undefined
}

function isCycleFormulaCell(args: Pick<ApplyPostRecalcDirectFormulaChangesArgs, 'state'>, cellIndex: number): boolean {
  return ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0
}

function countDirectDeltaApplication(
  args: Pick<ApplyPostRecalcDirectFormulaChangesArgs, 'state'>,
  formula: OperationPostRecalcFormula | undefined,
): void {
  if (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) {
    addEngineCounter(args.state.counters, 'directAggregateDeltaApplications')
  }
  if (formula?.directScalar !== undefined) {
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications')
  }
}
