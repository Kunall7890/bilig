import type { FormulaFamilyMember, FormulaFamilyRunUpsertArgs } from '../../formula/formula-family-store.js'

export type DeferredInitialFormulaFamilyRun = Omit<FormulaFamilyRunUpsertArgs, 'members'> & {
  axis: 'row'
  fixedIndex: number
  start: number
  step: number
  lastIndex: number
  ordered: boolean
  cellIndices: number[]
  rows?: number[]
}

export function materializeDeferredFormulaFamilyRunMembers(run: DeferredInitialFormulaFamilyRun): FormulaFamilyMember[] {
  const step = run.cellIndices.length <= 1 ? 1 : run.step
  return run.cellIndices.map((cellIndex, index) => ({
    cellIndex,
    row: run.ordered ? run.start + step * index : run.rows![index]!,
    col: run.fixedIndex,
  }))
}
