import type { Effect } from 'effect'
import type { CellSnapshot } from '@bilig/protocol'
import type { U32 } from '../runtime-state.js'
import type { EngineRecalcError } from '../errors.js'

export interface DirtyRegion {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export interface EngineRecalcService {
  readonly recalculateNow: () => Effect.Effect<number[], EngineRecalcError>
  readonly recalculateDirty: (dirtyRegions: ReadonlyArray<DirtyRegion>) => Effect.Effect<number[], EngineRecalcError>
  readonly recalculateDifferential: () => Effect.Effect<{ js: CellSnapshot[]; wasm: CellSnapshot[]; drift: string[] }, EngineRecalcError>
  readonly recalculatePreordered: (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots?: readonly number[] | U32,
  ) => Effect.Effect<U32, EngineRecalcError>
  readonly recalculate: (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots?: readonly number[] | U32,
  ) => Effect.Effect<U32, EngineRecalcError>
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => Effect.Effect<U32, EngineRecalcError>
  readonly recalculatePreorderedNowSync: (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32
  readonly recalculateNowSync: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
  readonly reconcilePivotOutputsNow: (baseChanged: U32, forceAllPivots?: boolean) => U32
}
