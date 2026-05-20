import { initialFormulaEntryRefAt, type InitialFormulaEntryRefSource } from './formula-initialization-refs.js'

const EMPTY_U32 = new Uint32Array(0)

export interface InitialFormulaCellIndexPlan {
  readonly targetCellIndices: Uint32Array
  readonly pendingInitialFormulaCellIndices: Uint32Array
  readonly pendingFormulaCells: Uint8Array | undefined
  readonly maxTargetCellIndex: number
}

export function createInitialFormulaCellIndexPlan<Entry>(args: {
  readonly refs: InitialFormulaEntryRefSource<Entry>
  readonly hadExistingFormulas: boolean
  readonly resolveCellIndex: (ref: Entry) => number
  readonly checkEvaluationBudget: () => void
}): InitialFormulaCellIndexPlan {
  const targetCellIndices = args.hadExistingFormulas ? EMPTY_U32 : new Uint32Array(args.refs.length)
  const pendingInitialFormulaCellIndices = args.hadExistingFormulas ? new Uint32Array(args.refs.length) : targetCellIndices
  let maxTargetCellIndex = 0

  for (let index = 0; index < args.refs.length; index += 1) {
    args.checkEvaluationBudget()
    const cellIndex = args.resolveCellIndex(initialFormulaEntryRefAt(args.refs, index))
    if (args.hadExistingFormulas) {
      pendingInitialFormulaCellIndices[index] = cellIndex
    } else {
      targetCellIndices[index] = cellIndex
      if (cellIndex > maxTargetCellIndex) {
        maxTargetCellIndex = cellIndex
      }
    }
  }

  const pendingFormulaCells = args.hadExistingFormulas ? undefined : new Uint8Array(maxTargetCellIndex + 1)
  if (pendingFormulaCells) {
    for (let index = 0; index < targetCellIndices.length; index += 1) {
      args.checkEvaluationBudget()
      pendingFormulaCells[targetCellIndices[index]!] = 1
    }
  }

  return {
    targetCellIndices,
    pendingInitialFormulaCellIndices,
    pendingFormulaCells,
    maxTargetCellIndex,
  }
}
