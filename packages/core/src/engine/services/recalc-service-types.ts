import type { Effect } from 'effect'
import type { CellSnapshot, CellValue, EngineChangedCell, WorkbookSnapshot } from '@bilig/protocol'
import type { EnginePatch } from '../../patches/patch-types.js'
import type { WorkbookPivotRecord } from '../../workbook-store.js'
import type { EngineRecalcError } from '../errors.js'
import type { EngineRuntimeState, SpillMaterialization, U32 } from '../runtime-state.js'
import type { EngineDirtyFrontierSchedulerService } from './dirty-frontier-scheduler-service.js'

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
  readonly recalculateAllNowSync: () => number[]
  readonly recalculateChangedValuesNowForRebuildSync: () => number[]
  readonly recalculateNowSync: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
  readonly reconcilePivotOutputsNow: (baseChanged: U32, forceAllPivots?: boolean) => U32
}

export interface EngineRecalcServiceArgs {
  readonly state: Pick<
    EngineRuntimeState,
    'workbook' | 'strings' | 'wasm' | 'formulas' | 'ranges' | 'events' | 'counters' | 'getLastMetrics' | 'setLastMetrics'
  >
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly exportSnapshot: () => WorkbookSnapshot
  readonly importSnapshot: (snapshot: WorkbookSnapshot) => void
  readonly beginMutationCollection: () => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly captureChangedPatches: (
    changedCellIndices: readonly number[] | U32,
    request?: {
      invalidation?: 'cells' | 'full'
      invalidatedRanges?: readonly {
        sheetName: string
        startAddress: string
        endAddress: string
      }[]
      invalidatedRows?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
      invalidatedColumns?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    },
  ) => readonly EnginePatch[]
  readonly unionChangedSets: (...sets: Array<readonly number[] | U32>) => U32
  readonly composeChangedRootsAndOrdered: (changedRoots: readonly number[] | U32, ordered: U32, orderedCount: number) => U32
  readonly emptyChangedSet: () => U32
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly getPendingKernelSync: () => U32
  readonly getDeferredKernelSyncCount: () => number
  readonly setDeferredKernelSyncCount: (next: number) => void
  readonly getDeferredKernelSyncEpoch: () => number
  readonly setDeferredKernelSyncEpoch: (next: number) => void
  readonly getDeferredKernelSyncSeen: () => U32
  readonly getWasmBatch: () => U32
  readonly getChangedInputBuffer: () => U32
  readonly flushWasmProgramSync: () => void
  readonly beginEvaluationBudget: (startedAtMs: number) => void
  readonly endEvaluationBudget: () => void
  readonly checkEvaluationBudget: (stepCost?: number) => void
  readonly now: () => Date
  readonly random: () => number
  readonly performanceNow: () => number
  readonly dirtyScheduler: EngineDirtyFrontierSchedulerService
  readonly materializeSpill: (cellIndex: number, arrayValue: { values: CellValue[]; rows: number; cols: number }) => SpillMaterialization
  readonly clearOwnedSpill: (cellIndex: number) => number[]
  readonly evaluateDirectLookupFormula: (cellIndex: number) => number[] | undefined
  readonly evaluateUnsupportedFormula: (cellIndex: number) => number[]
  readonly materializePivot: (pivot: WorkbookPivotRecord) => number[]
  readonly forEachFormulaDependencyCell: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void
}
