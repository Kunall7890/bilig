import type { FormulaInstanceTable } from '../../formula/formula-instance-table.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import type { StructuralFormulaRebindInput } from './structure-service-types.js'

export function rewriteCachedFormulaRuntimeSource(
  state: EngineRuntimeState,
  formulaInstances: FormulaInstanceTable,
  input: StructuralFormulaRebindInput,
): boolean {
  const existing = state.formulas.get(input.cellIndex)
  if (input.preservesBinding !== true || input.preservesValue !== true || existing?.preserveCachedValueOnFullRecalc !== true) {
    return false
  }

  existing.source = input.source
  existing.plan = { ...existing.plan, source: input.source }
  formulaInstances.upsert({
    cellIndex: input.cellIndex,
    sheetName: input.ownerSheetName,
    row: input.ownerRow,
    col: input.ownerCol,
    source: input.source,
  })
  return true
}
