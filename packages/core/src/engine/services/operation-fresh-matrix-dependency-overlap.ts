import type {
  EngineRuntimeState,
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeFormula,
} from '../runtime-state.js'

export interface FreshMatrixDependencyOverlapArgs {
  readonly getRegionFormulaSubscriptionCount?: () => number
  readonly hasRegionFormulaSubscriptionsIntersectingRect:
    | ((sheetId: number, rowStart: number, rowEnd: number, colStart: number, colEnd: number) => boolean)
    | undefined
  readonly hasRegionFormulaSubscriptionsOverlappingRange?: (
    sheetId: number,
    rowStart: number,
    rowEnd: number,
    colStart: number,
    colEnd: number,
  ) => boolean
  readonly state: Pick<EngineRuntimeState, 'formulas' | 'ranges'>
}

export interface FreshMatrixDependencyOverlapBatch {
  readonly colStart: number
  readonly formulaCol: number
  readonly rowCount: number
  readonly rowStart: number
  readonly sheet: { readonly name: string }
  readonly sheetId: number
}

export function freshMatrixOverlapsFormulaDependencies(
  args: FreshMatrixDependencyOverlapArgs,
  matrix: FreshMatrixDependencyOverlapBatch,
): boolean {
  const rowEnd = matrix.rowStart + matrix.rowCount - 1
  const colEnd = matrix.formulaCol
  if (
    args.getRegionFormulaSubscriptionCount?.() === args.state.formulas.size &&
    args.hasRegionFormulaSubscriptionsOverlappingRange?.(matrix.sheetId, matrix.rowStart, rowEnd, matrix.colStart, colEnd) === false
  ) {
    return false
  }
  if (
    args.state.ranges.size === 0 &&
    args.hasRegionFormulaSubscriptionsIntersectingRect?.(matrix.sheetId, matrix.rowStart, rowEnd, matrix.colStart, colEnd) === false
  ) {
    return false
  }
  let overlaps = false
  args.state.formulas.forEach((formula) => {
    if (overlaps) {
      return
    }
    overlaps =
      directAggregateOverlapsFreshMatrix(formula.directAggregate, matrix.sheet.name, matrix.rowStart, rowEnd, matrix.colStart, colEnd) ||
      directCriteriaOverlapsFreshMatrix(formula.directCriteria, matrix.sheet.name, matrix.rowStart, rowEnd, matrix.colStart, colEnd) ||
      rangeDependenciesOverlapFreshMatrix(args, formula, matrix.sheetId, matrix.rowStart, rowEnd, matrix.colStart, colEnd)
  })
  return overlaps
}

function directAggregateOverlapsFreshMatrix(
  aggregate: RuntimeDirectAggregateDescriptor | undefined,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  return (
    aggregate !== undefined &&
    aggregate.sheetName === sheetName &&
    aggregate.rowEnd >= rowStart &&
    aggregate.rowStart <= rowEnd &&
    aggregate.colEnd >= colStart &&
    aggregate.col <= colEnd
  )
}

function directCriteriaOverlapsFreshMatrix(
  criteria: RuntimeDirectCriteriaDescriptor | undefined,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  if (criteria === undefined) {
    return false
  }
  const rangeOverlaps = (range: {
    readonly col: number
    readonly rowEnd: number
    readonly rowStart: number
    readonly sheetName: string
  }): boolean =>
    range.sheetName === sheetName && range.rowEnd >= rowStart && range.rowStart <= rowEnd && range.col >= colStart && range.col <= colEnd
  return (
    (criteria.aggregateRange !== undefined && rangeOverlaps(criteria.aggregateRange)) ||
    criteria.criteriaPairs.some((pair) => rangeOverlaps(pair.range))
  )
}

function rangeDependenciesOverlapFreshMatrix(
  args: FreshMatrixDependencyOverlapArgs,
  formula: RuntimeFormula,
  sheetId: number,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  for (let index = 0; index < formula.rangeDependencies.length; index += 1) {
    if (rangeDescriptorOverlapsFreshMatrix(args, formula.rangeDependencies[index]!, sheetId, rowStart, rowEnd, colStart, colEnd)) {
      return true
    }
  }
  for (let index = 0; index < formula.graphRangeDependencies.length; index += 1) {
    if (rangeDescriptorOverlapsFreshMatrix(args, formula.graphRangeDependencies[index]!, sheetId, rowStart, rowEnd, colStart, colEnd)) {
      return true
    }
  }
  return false
}

function rangeDescriptorOverlapsFreshMatrix(
  args: FreshMatrixDependencyOverlapArgs,
  rangeIndex: number,
  sheetId: number,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  const descriptor = args.state.ranges.getDescriptor(rangeIndex)
  return (
    descriptor.sheetId === sheetId &&
    descriptor.row2 >= rowStart &&
    descriptor.row1 <= rowEnd &&
    descriptor.col2 >= colStart &&
    descriptor.col1 <= colEnd
  )
}
