import {
  composeFormulaFamilyStructuralSourceTransform,
  type FormulaFamilyStructuralSourceTransform,
} from '../../formula/formula-family-store.js'
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
  let transforms: Map<number, FormulaFamilyStructuralSourceTransform> | undefined
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
    const existingTransform = args.existingTransforms?.get(runIndex)
    const nextTransform = existingTransform
      ? composeFormulaFamilyStructuralSourceTransform(existingTransform, args.transform)
      : args.transform
    if (!nextTransform) {
      return undefined
    }
    transforms ??= new Map(args.existingTransforms)
    transforms.set(runIndex, nextTransform)
  }
  if (transforms === undefined || memberCount !== args.ownedFormulaCount) {
    return undefined
  }
  return { memberCount, transforms }
}
