import type {
  FormulaFamilyMember,
  FormulaFamilyRunUpsertArgs,
  FormulaFamilyStore,
  FormulaFamilyStructuralSourceTransform,
} from '../../formula/formula-family-store.js'
import type { FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'

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

export function registerDeferredFormulaFamilyIndexRunsNow(args: {
  readonly formulaFamilies: FormulaFamilyStore
  readonly formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache
  readonly runs: readonly DeferredInitialFormulaFamilyRun[]
  readonly structuralSourceTransforms?: ReadonlyMap<number, FormulaFamilyStructuralSourceTransform>
}): void {
  args.formulaFamilies.clear()
  args.formulaFamilyShapeKeyCache.clear()
  args.runs.forEach((run, runIndex) => {
    const step = run.cellIndices.length <= 1 ? 1 : run.step
    if (
      run.ordered &&
      step > 0 &&
      args.formulaFamilies.registerFreshUniformRun({
        sheetId: run.sheetId,
        templateId: run.templateId,
        shapeKey: run.shapeKey,
        axis: run.axis,
        fixedIndex: run.fixedIndex,
        start: run.start,
        step,
        cellIndices: run.cellIndices,
      })
    ) {
      const transform = args.structuralSourceTransforms?.get(runIndex)
      if (transform !== undefined) {
        const membership = args.formulaFamilies.getMembership(run.cellIndices[0]!)
        if (membership) {
          args.formulaFamilies.setStructuralSourceTransform(membership.familyId, transform)
        }
      }
      return
    }
    args.formulaFamilies.registerFormulaRun({
      sheetId: run.sheetId,
      templateId: run.templateId,
      shapeKey: run.shapeKey,
      members: materializeDeferredFormulaFamilyRunMembers(run),
    })
    const transform = args.structuralSourceTransforms?.get(runIndex)
    if (transform !== undefined) {
      const membership = args.formulaFamilies.getMembership(run.cellIndices[0]!)
      if (membership) {
        args.formulaFamilies.setStructuralSourceTransform(membership.familyId, transform)
      }
    }
  })
}
