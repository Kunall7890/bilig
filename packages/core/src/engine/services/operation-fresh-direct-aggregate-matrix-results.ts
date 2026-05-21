import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import {
  evaluateFreshDirectAggregateMatrixNumericRow,
  evaluateFreshDirectAggregateMatrixRow,
} from './operation-fresh-direct-aggregate-matrix-helpers.js'
import type { FreshDirectAggregateFormulaEntrySeed } from './operation-fresh-direct-aggregate-formula-batch-records.js'
import type { OperationFreshDirectAggregateFormulaBatchFastPathArgs } from './operation-fresh-direct-aggregate-formula-batch-fast-path.js'
import { tryEvaluateNativeFreshDirectAggregateMatrixResults } from './operation-fresh-direct-aggregate-native-batch.js'

export function materializeFreshDirectAggregateFormulaResults(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  input: {
    readonly inputColCount: number
    readonly matrixColStart: number
    readonly seeds: readonly FreshDirectAggregateFormulaEntrySeed[]
    readonly values: Float64Array
  },
): Float64Array | DirectScalarCurrentOperand[] {
  args.checkEvaluationBudget(input.seeds.length * Math.max(1, input.inputColCount))
  const nativeResults = tryEvaluateNativeFreshDirectAggregateMatrixResults(args, input)
  if (nativeResults !== undefined) {
    return nativeResults
  }
  const numericResults = new Float64Array(input.seeds.length)
  for (let rowOffset = 0; rowOffset < input.seeds.length; rowOffset += 1) {
    const seed = input.seeds[rowOffset]!
    args.checkEvaluationBudget(input.inputColCount)
    const result = evaluateFreshDirectAggregateMatrixNumericRow({
      aggregateKind: seed.aggregateKind,
      colEnd: seed.aggregateColEnd,
      colStart: seed.aggregateColStart,
      inputColCount: input.inputColCount,
      matrixColStart: input.matrixColStart,
      resultOffset: seed.resultOffset,
      rowOffset,
      values: input.values,
    })
    if (result === undefined) {
      return materializeFreshDirectAggregateFormulaObjectResults(args, input, rowOffset)
    }
    numericResults[rowOffset] = result
  }
  return numericResults
}

function materializeFreshDirectAggregateFormulaObjectResults(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  input: {
    readonly inputColCount: number
    readonly matrixColStart: number
    readonly seeds: readonly FreshDirectAggregateFormulaEntrySeed[]
    readonly values: Float64Array
  },
  startRowOffset: number,
): DirectScalarCurrentOperand[] {
  const results: DirectScalarCurrentOperand[] = []
  results.length = input.seeds.length
  for (let rowOffset = 0; rowOffset < input.seeds.length; rowOffset += 1) {
    const seed = input.seeds[rowOffset]!
    if (rowOffset > startRowOffset) {
      args.checkEvaluationBudget(input.inputColCount)
    }
    results[rowOffset] = evaluateFreshDirectAggregateMatrixRow({
      aggregateKind: seed.aggregateKind,
      colEnd: seed.aggregateColEnd,
      colStart: seed.aggregateColStart,
      inputColCount: input.inputColCount,
      matrixColStart: input.matrixColStart,
      resultOffset: seed.resultOffset,
      rowOffset,
      values: input.values,
    })
  }
  return results
}
