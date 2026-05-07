import type { RuntimeFormula } from '../runtime-state.js'
import { getRuntimeFormulaSource, getRuntimeFormulaStructuralCompiled } from '../runtime-formula-source.js'
import { isCellIndexMapped } from './structure-runtime-cleanup.js'
import type { CreateEngineStructureServiceArgs, StructuralFormulaRebindInput } from './structure-service-types.js'

export function materializeDeferredStructuralFormulaSources(
  args: CreateEngineStructureServiceArgs,
  hasDeferredStructuralFormulaSources: boolean,
): boolean {
  if (!hasDeferredStructuralFormulaSources) {
    return false
  }
  const inputs: StructuralFormulaRebindInput[] = []
  const enqueueMaterializedFormulaSource = (
    formula: RuntimeFormula,
    cellIndex: number,
    structuralSourceTransform: NonNullable<RuntimeFormula['structuralSourceTransform']>,
  ): void => {
    if (
      formula.directLookup !== undefined ||
      formula.directCriteria !== undefined ||
      (formula.directAggregate === undefined && formula.rangeDependencies.length !== 0) ||
      !isCellIndexMapped(args, cellIndex)
    ) {
      return
    }
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
    if (sheetId === undefined || !ownerPosition) {
      return
    }
    const ownerSheetName = args.state.workbook.getSheetNameById(sheetId)
    if (!ownerSheetName) {
      return
    }
    const source = getRuntimeFormulaSource(formula, structuralSourceTransform)
    const compiled = getRuntimeFormulaStructuralCompiled(formula, structuralSourceTransform)
    const preservesValue = structuralSourceTransform.preservesValue
    inputs.push({
      cellIndex,
      ownerSheetName,
      ownerRow: ownerPosition.row,
      ownerCol: ownerPosition.col,
      source,
      ...(compiled
        ? {
            compiled,
            preservesBinding: true,
            preservesValue,
          }
        : {}),
    })
  }
  args.consumeFormulaFamilyStructuralSourceTransforms().forEach((entry) => {
    entry.cellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula || formula.structuralSourceTransform !== undefined) {
        return
      }
      enqueueMaterializedFormulaSource(formula, cellIndex, entry.transform)
    })
  })
  args.state.formulas.forEach((formula, cellIndex) => {
    if (formula.structuralSourceTransform === undefined) {
      return
    }
    enqueueMaterializedFormulaSource(formula, cellIndex, formula.structuralSourceTransform)
  })
  if (inputs.length > 0) {
    args.rebindFormulaCells(inputs)
  }
  inputs.forEach(({ cellIndex }) => {
    const formula = args.state.formulas.get(cellIndex)
    if (formula) {
      formula.structuralSourceTransform = undefined
    }
  })
  return false
}
