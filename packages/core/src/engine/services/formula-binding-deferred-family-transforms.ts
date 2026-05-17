import type { FormulaFamilyStructuralSourceTransform } from '../../formula/formula-family-store.js'
import type { DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'

export function queueDeferredFormulaFamilyStructuralSourceTransforms(args: {
  readonly runs: readonly DeferredInitialFormulaFamilyRun[] | undefined
  readonly existingTransforms: ReadonlyMap<number, FormulaFamilyStructuralSourceTransform> | undefined
  readonly sheetId: number
  readonly transform: FormulaFamilyStructuralSourceTransform
  readonly ownedFormulaCount: number
  readonly canDeferCellIndex: (cellIndex: number) => boolean
}): { readonly memberCount: number; readonly transforms: Map<number, FormulaFamilyStructuralSourceTransform> } | undefined {
  if (args.runs === undefined) {
    return undefined
  }
  let memberCount = 0
  const runIndexes: number[] = []
  for (let runIndex = 0; runIndex < args.runs.length; runIndex += 1) {
    const run = args.runs[runIndex]!
    if (run.sheetId !== args.sheetId) {
      continue
    }
    const representativeCellIndex = run.cellIndices[0]
    if (representativeCellIndex === undefined || !args.canDeferCellIndex(representativeCellIndex)) {
      return undefined
    }
    memberCount += run.cellIndices.length
    runIndexes.push(runIndex)
  }
  if (runIndexes.length === 0 || memberCount !== args.ownedFormulaCount) {
    return undefined
  }
  const transforms = new Map(args.existingTransforms)
  runIndexes.forEach((runIndex) => {
    transforms.set(runIndex, args.transform)
  })
  return { memberCount, transforms }
}
