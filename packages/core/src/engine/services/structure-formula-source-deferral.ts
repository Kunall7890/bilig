import type { StructuralAxisTransform } from '@bilig/formula'
import { mapStructuralAxisIndex } from '../../engine-structural-utils.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { isStructurallyStableSimpleFormulaNode } from './structure-formula-rewrite-guards.js'
import { shouldCaptureStoredCell } from './structure-cell-state.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'

export function isSimpleStructuralFormulaSourceDeferrable(args: CreateEngineStructureServiceArgs, formula: RuntimeFormula): boolean {
  return (
    formula.rangeDependencies.length === 0 &&
    formula.dependencyIndices.every((dependencyCellIndex) => shouldCaptureStoredCell(args, dependencyCellIndex)) &&
    !formula.compiled.volatile &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    formula.directLookup === undefined &&
    formula.directAggregate === undefined &&
    formula.directCriteria === undefined &&
    isStructurallyStableSimpleFormulaNode(formula.compiled.ast)
  )
}

export function canDeferSimpleStructuralFormulaSource(
  args: CreateEngineStructureServiceArgs,
  formula: RuntimeFormula,
  transform: StructuralAxisTransform,
): boolean {
  return transform.kind !== 'delete' && transform.axis === 'column' && isSimpleStructuralFormulaSourceDeferrable(args, formula)
}

export function canDeferSimpleDeleteStructuralFormulaSource(
  args: CreateEngineStructureServiceArgs,
  formula: RuntimeFormula,
  targetSheetId: number | undefined,
  transform: StructuralAxisTransform,
): boolean {
  if (
    transform.kind !== 'delete' ||
    transform.axis !== 'column' ||
    targetSheetId === undefined ||
    !isSimpleStructuralFormulaSourceDeferrable(args, formula)
  ) {
    return false
  }
  return formula.dependencyIndices.every((dependencyCellIndex) => {
    if (args.state.workbook.cellStore.sheetIds[dependencyCellIndex] !== targetSheetId) {
      return true
    }
    const dependencyColumn = args.state.workbook.getCellAxisIndex(dependencyCellIndex, 'column')
    return dependencyColumn !== undefined && mapStructuralAxisIndex(dependencyColumn, transform) !== undefined
  })
}

export function canDeferSimpleDeleteRefErrorFormulaSource(
  args: CreateEngineStructureServiceArgs,
  formula: RuntimeFormula,
  targetSheetId: number | undefined,
  transform: StructuralAxisTransform,
): boolean {
  if (
    transform.kind !== 'delete' ||
    transform.axis !== 'column' ||
    targetSheetId === undefined ||
    !isSimpleStructuralFormulaSourceDeferrable(args, formula)
  ) {
    return false
  }
  return formula.dependencyIndices.some((dependencyCellIndex) => {
    if (args.state.workbook.cellStore.sheetIds[dependencyCellIndex] !== targetSheetId) {
      return false
    }
    const dependencyColumn = args.state.workbook.getCellAxisIndex(dependencyCellIndex, 'column')
    return dependencyColumn !== undefined && mapStructuralAxisIndex(dependencyColumn, transform) === undefined
  })
}
