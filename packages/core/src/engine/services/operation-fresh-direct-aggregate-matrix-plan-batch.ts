import type { CompiledFormula } from '@bilig/formula'
import type { EngineCellMutationRef, EngineFreshDirectAggregateMatrixPlan } from '../../cell-mutations-at.js'
import type {
  FreshDirectAggregateMatrixBatch,
  OperationFreshDirectAggregateFormulaBatchFastPathArgs,
} from './operation-fresh-direct-aggregate-formula-batch-fast-path.js'
import {
  createFreshMatrixDirectAggregateTemplate,
  normalizeFreshMatrixDirectAggregateOffset,
  tryTranslateFreshMatrixDirectAggregateTemplate,
  type FreshMatrixDirectAggregateTemplate,
} from './operation-fresh-direct-aggregate-matrix-helpers.js'
import type { FreshDirectAggregateFormulaEntrySeed } from './operation-fresh-direct-aggregate-formula-batch-records.js'
import { rethrowFatalFormulaBindingError } from './formula-binding-error-policy.js'
import { materializeFreshDirectAggregateFormulaResults } from './operation-fresh-direct-aggregate-matrix-results.js'

const FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT = 4096

export function collectFreshDirectAggregateMatrixPlanBatch(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
  plan: EngineFreshDirectAggregateMatrixPlan,
): FreshDirectAggregateMatrixBatch | null {
  const valueCount = plan.rowCount * plan.inputColCount
  const formulaCol = plan.colStart + plan.inputColCount
  if (
    plan.sheetId !== firstRef.sheetId ||
    plan.rowCount < 2 ||
    plan.inputColCount < 2 ||
    plan.values.length !== valueCount ||
    plan.formulaSources.length !== plan.rowCount ||
    refs.length !== valueCount + plan.rowCount
  ) {
    return null
  }
  const firstMutation = firstRef.mutation
  if (
    firstRef.cellIndex !== undefined ||
    firstMutation.kind !== 'setCellValue' ||
    firstMutation.row !== plan.rowStart ||
    firstMutation.col !== plan.colStart
  ) {
    return null
  }
  const firstFormulaRef = refs[valueCount]
  const lastFormulaRef = refs[refs.length - 1]
  if (
    !isMatrixBoundaryFormulaRef(firstFormulaRef, plan.sheetId, plan.rowStart, formulaCol, plan.formulaSources[0]) ||
    !isMatrixBoundaryFormulaRef(
      lastFormulaRef,
      plan.sheetId,
      plan.rowStart + plan.rowCount - 1,
      formulaCol,
      plan.formulaSources[plan.rowCount - 1],
    )
  ) {
    return null
  }
  const sheet = args.state.workbook.getSheetById(plan.sheetId)
  if (!sheet) {
    return null
  }
  const totalColCount = plan.inputColCount + 1
  let valueIndex = 0
  for (let rowOffset = 0; rowOffset < plan.rowCount; rowOffset += 1) {
    args.checkEvaluationBudget(totalColCount)
    const row = plan.rowStart + rowOffset
    for (let colOffset = 0; colOffset < plan.inputColCount; colOffset += 1) {
      const col = plan.colStart + colOffset
      if (!isMatrixValueRef(refs[valueIndex], plan.sheetId, row, col, plan.values[valueIndex]!)) {
        return null
      }
      if (sheet.grid.getPhysical(row, col) !== -1 || sheet.logical.getVisibleCell(row, col) !== undefined) {
        return null
      }
      valueIndex += 1
    }
    if (sheet.grid.getPhysical(row, formulaCol) !== -1 || sheet.logical.getVisibleCell(row, formulaCol) !== undefined) {
      return null
    }
  }
  for (let col = plan.colStart; col < formulaCol; col += 1) {
    if (args.hasTrackedExactLookupDependents(plan.sheetId, col) || args.hasTrackedSortedLookupDependents(plan.sheetId, col)) {
      return null
    }
  }

  const formulaEntrySeeds: FreshDirectAggregateFormulaEntrySeed[] = []
  let directAggregateTemplate: FreshMatrixDirectAggregateTemplate | undefined
  for (let rowOffset = 0; rowOffset < plan.rowCount; rowOffset += 1) {
    args.checkEvaluationBudget()
    const source = plan.formulaSources[rowOffset]
    const formulaRef = refs[valueCount + rowOffset]
    const row = plan.rowStart + rowOffset
    if (
      typeof source !== 'string' ||
      !isMatrixBoundaryFormulaRef(formulaRef, plan.sheetId, row, formulaCol, source) ||
      args.hasTrackedExactLookupDependents(plan.sheetId, formulaCol) ||
      args.hasTrackedSortedLookupDependents(plan.sheetId, formulaCol) ||
      args.hasTrackedDirectRangeDependents(plan.sheetId, formulaCol)
    ) {
      return null
    }
    let compiled: CompiledFormula
    let templateId: number
    const translated = directAggregateTemplate
      ? tryTranslateFreshMatrixDirectAggregateTemplate(directAggregateTemplate, source, row, formulaCol)
      : undefined
    if (translated) {
      compiled = translated
      templateId = directAggregateTemplate!.templateId
    } else {
      let template: ReturnType<OperationFreshDirectAggregateFormulaBatchFastPathArgs['compileTemplateFormula']>
      try {
        template = args.compileTemplateFormula(source, row, formulaCol)
      } catch (error) {
        rethrowFatalFormulaBindingError(error)
        return null
      }
      compiled = template.compiled
      templateId = template.templateId
    }
    if (
      compiled.volatile ||
      compiled.producesSpill ||
      compiled.symbolicNames.length !== 0 ||
      compiled.symbolicTables.length !== 0 ||
      compiled.symbolicSpills.length !== 0
    ) {
      return null
    }
    const aggregate = compiled.directAggregateCandidate
    const range = aggregate === undefined ? undefined : compiled.parsedSymbolicRanges?.[aggregate.symbolicRangeIndex]
    if (
      aggregate === undefined ||
      range === undefined ||
      range.refKind !== 'cells' ||
      (range.sheetName ?? sheet.name) !== sheet.name ||
      range.startRow !== range.endRow ||
      range.startRow !== row ||
      range.startCol < plan.colStart ||
      range.endCol >= formulaCol ||
      range.startCol > range.endCol ||
      range.endCol - range.startCol + 1 > FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT
    ) {
      return null
    }
    if (directAggregateTemplate === undefined) {
      directAggregateTemplate = createFreshMatrixDirectAggregateTemplate({
        aggregate,
        compiled,
        formulaCol,
        range,
        row,
        templateId,
      })
    }
    formulaEntrySeeds.push({
      row,
      col: formulaCol,
      source,
      compiled,
      templateId,
      aggregateKind: aggregate.aggregateKind,
      aggregateRowStart: range.startRow,
      aggregateRowEnd: range.endRow,
      aggregateColStart: range.startCol,
      aggregateColEnd: range.endCol,
      resultOffset: normalizeFreshMatrixDirectAggregateOffset(aggregate.resultOffset),
    })
  }
  const formulaResults = materializeFreshDirectAggregateFormulaResults(args, {
    inputColCount: plan.inputColCount,
    matrixColStart: plan.colStart,
    seeds: formulaEntrySeeds,
    values: plan.values,
  })

  return {
    sheet,
    sheetId: plan.sheetId,
    rowStart: plan.rowStart,
    rowCount: plan.rowCount,
    colStart: plan.colStart,
    inputColCount: plan.inputColCount,
    formulaCol,
    values: plan.values,
    formulaEntries: formulaEntrySeeds,
    formulaResults,
  }
}

function isMatrixValueRef(ref: EngineCellMutationRef | undefined, sheetId: number, row: number, col: number, value: number): boolean {
  return (
    ref !== undefined &&
    ref.sheetId === sheetId &&
    ref.cellIndex === undefined &&
    ref.mutation.kind === 'setCellValue' &&
    ref.mutation.row === row &&
    ref.mutation.col === col &&
    typeof ref.mutation.value === 'number' &&
    !Object.is(ref.mutation.value, -0) &&
    Object.is(ref.mutation.value, value)
  )
}

function isMatrixBoundaryFormulaRef(
  ref: EngineCellMutationRef | undefined,
  sheetId: number,
  row: number,
  col: number,
  formula: string | undefined,
): boolean {
  return (
    ref !== undefined &&
    ref.sheetId === sheetId &&
    ref.cellIndex === undefined &&
    ref.mutation.kind === 'setCellFormula' &&
    ref.mutation.row === row &&
    ref.mutation.col === col &&
    ref.mutation.formula === formula
  )
}
