import { markFormulaCellBound } from './formula-binding-cell-flags.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'
import { updateFormulaBindingVolatileIndex } from './formula-binding-volatile-index.js'
import type { RuntimeFormula } from '../runtime-state.js'

type FormulaBindingTrackedMetadataArgs = Pick<CreateEngineFormulaBindingServiceArgs, 'state' | 'volatileFormulaCells'>

export function refreshFormulaBindingTrackedMetadata(
  args: FormulaBindingTrackedMetadataArgs,
  cellIndex: number,
  formula: RuntimeFormula,
): void {
  args.state.formulas.refreshTrackedMetadata(cellIndex)
  updateFormulaBindingVolatileIndex(args.volatileFormulaCells, cellIndex, formula)
  markFormulaCellBound(args.state.workbook.cellStore, cellIndex, formula.compiled.mode)
}
