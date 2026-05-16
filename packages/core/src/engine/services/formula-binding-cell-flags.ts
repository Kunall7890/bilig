import { FormulaMode } from '@bilig/protocol'
import { CellFlags, type CellStore } from '../../cell-store.js'

const DERIVED_OUTPUT_FLAGS = CellFlags.SpillChild | CellFlags.PivotOutput
const FORMULA_RUNTIME_FLAGS = CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput

export function markFormulaCellBound(cellStore: Pick<CellStore, 'flags'>, cellIndex: number, mode: FormulaMode): void {
  let nextFlags = ((cellStore.flags[cellIndex] ?? 0) & ~DERIVED_OUTPUT_FLAGS) | CellFlags.HasFormula
  nextFlags = mode === FormulaMode.JsOnly ? nextFlags | CellFlags.JsOnly : nextFlags & ~CellFlags.JsOnly
  cellStore.flags[cellIndex] = nextFlags
}

export function clearFormulaRuntimeFlags(cellStore: Pick<CellStore, 'flags'>, cellIndex: number): void {
  cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~FORMULA_RUNTIME_FLAGS
}
