import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { U32 } from '../runtime-state.js'
import { evaluateInitialDirectScalar, evaluateInitialDirectScalarNumber } from './formula-initialization-direct-scalar.js'
import { evaluateInitialPrefixAggregateGroups } from './formula-initialization-prefix-aggregates.js'
import { canEvaluateInitialDirectRuntimeFormula } from './formula-initialization-predicates.js'
import type { EngineFormulaInitializationServiceArgs } from './formula-initialization-service-types.js'
import type { InitialFormulaCellIndexList } from './formula-initialization-refs.js'
import { createInitialFormulaValueWriter } from './formula-initialization-value-writer.js'

export const INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT = 16_384

export interface InitialDirectFormulaEvaluationOptions {
  readonly alreadyValidated?: boolean
  readonly hasPrefixAggregateCandidates?: boolean
  readonly preEvaluatedCellIndices?: InitialFormulaCellIndexList
  readonly preEvaluatedCellCount?: number
}

export function evaluateInitialDirectFormulas(
  args: EngineFormulaInitializationServiceArgs,
  orderedCellIndices: InitialFormulaCellIndexList,
  options?: InitialDirectFormulaEvaluationOptions,
): U32 | undefined {
  if (
    orderedCellIndices.length === 0 ||
    orderedCellIndices.length > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT ||
    (options?.alreadyValidated !== true &&
      !orderedCellIndices.every((cellIndex) => canEvaluateInitialDirectRuntimeFormula(args.state.formulas.get(cellIndex))))
  ) {
    return undefined
  }
  let changedCellBuffer = new Uint32Array(Math.max(orderedCellIndices.length, 1))
  let changedCellCount = 0
  const preEvaluatedCellIndices = options?.preEvaluatedCellIndices
  const preEvaluatedCellCount = Math.min(
    options?.preEvaluatedCellCount ?? preEvaluatedCellIndices?.length ?? 0,
    preEvaluatedCellIndices?.length ?? 0,
  )
  let preEvaluatedCellIndex = 0
  let preEvaluatedCells: Uint8Array | undefined
  const pushChangedCellIndex = (cellIndex: number): void => {
    if (changedCellCount === changedCellBuffer.length) {
      const next = new Uint32Array(changedCellBuffer.length * 2)
      next.set(changedCellBuffer)
      changedCellBuffer = next
    }
    changedCellBuffer[changedCellCount] = cellIndex
    changedCellCount += 1
  }
  const isPreEvaluatedCell = (cellIndex: number): boolean => {
    if (!preEvaluatedCellIndices || preEvaluatedCellCount === 0) {
      return false
    }
    if (!preEvaluatedCells) {
      preEvaluatedCells = new Uint8Array(args.state.workbook.cellStore.size + 1)
      for (let index = 0; index < preEvaluatedCellCount; index += 1) {
        preEvaluatedCells[preEvaluatedCellIndices[index]!] = 1
      }
    }
    return preEvaluatedCells[cellIndex] === 1
  }
  const canReusePreEvaluatedFormula = (cellIndex: number): boolean => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return false
    }
    for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
      const dependencyCellIndex = formula.dependencyIndices[index]!
      if (
        ((args.state.workbook.cellStore.flags[dependencyCellIndex] ?? 0) & CellFlags.HasFormula) !== 0 &&
        !isPreEvaluatedCell(dependencyCellIndex)
      ) {
        return false
      }
    }
    return true
  }
  const shouldReusePreEvaluatedCell = (cellIndex: number): boolean => {
    if (!preEvaluatedCellIndices || preEvaluatedCellIndex >= preEvaluatedCellCount) {
      return false
    }
    if (preEvaluatedCellIndices[preEvaluatedCellIndex] !== cellIndex) {
      return false
    }
    preEvaluatedCellIndex += 1
    return canReusePreEvaluatedFormula(cellIndex)
  }
  const valueWriter = createInitialFormulaValueWriter(args)
  args.state.workbook.withBatchedColumnVersionUpdates(() => {
    const prefixAggregateHandled =
      options?.hasPrefixAggregateCandidates === true
        ? evaluateInitialPrefixAggregateGroups(args, orderedCellIndices, pushChangedCellIndex, valueWriter.writeValue)
        : undefined
    for (let index = 0; index < orderedCellIndices.length; index += 1) {
      args.checkEvaluationBudget()
      const cellIndex = orderedCellIndices[index]!
      if (prefixAggregateHandled?.has(cellIndex)) {
        continue
      }
      if (shouldReusePreEvaluatedCell(cellIndex)) {
        pushChangedCellIndex(cellIndex)
        continue
      }
      const formula = args.state.formulas.get(cellIndex)
      if (formula?.directScalar !== undefined) {
        const numericValue = evaluateInitialDirectScalarNumber(args.state, formula.directScalar)
        if (numericValue !== undefined) {
          valueWriter.writeNumber(cellIndex, numericValue)
          pushChangedCellIndex(cellIndex)
          continue
        }
        const fallbackValue = evaluateInitialDirectScalar(args.state, formula.directScalar)
        if (fallbackValue !== undefined) {
          valueWriter.writeValue(cellIndex, fallbackValue)
          pushChangedCellIndex(cellIndex)
          continue
        }
      }
      const changedSpillIndices = args.evaluateDirectFormula(cellIndex)
      pushChangedCellIndex(cellIndex)
      if (changedSpillIndices) {
        for (let spillIndex = 0; spillIndex < changedSpillIndices.length; spillIndex += 1) {
          pushChangedCellIndex(changedSpillIndices[spillIndex]!)
        }
      }
    }
    valueWriter.flush()
  })
  const changedCellIndices = changedCellBuffer.subarray(0, changedCellCount)
  args.deferKernelSync(changedCellIndices)
  addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedCellIndices.length)
  return changedCellIndices
}
