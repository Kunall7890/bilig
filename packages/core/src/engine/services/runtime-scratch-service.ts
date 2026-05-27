import { Effect } from 'effect'
import { growUint32 } from '../../engine-buffer-utils.js'
import type { U32 } from '../runtime-state.js'
import { EngineRuntimeScratchError } from '../errors.js'

export const INITIAL_RUNTIME_SCRATCH_CAPACITY = 16

function createScratchBuffer(): U32 {
  return new Uint32Array(INITIAL_RUNTIME_SCRATCH_CAPACITY)
}

function growScratchBuffer(buffer: U32 | undefined, size: number): U32 | undefined {
  if (size <= (buffer?.length ?? INITIAL_RUNTIME_SCRATCH_CAPACITY)) {
    return buffer
  }
  return growUint32(buffer ?? createScratchBuffer(), size)
}

function scratchErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export interface EngineRuntimeScratchService {
  readonly ensureRecalcCapacity: (size: number) => Effect.Effect<void, EngineRuntimeScratchError>
  readonly ensureRecalcCapacityNow: (size: number) => void
  readonly getPendingKernelSyncNow: () => U32
  readonly setPendingKernelSyncNow: (next: U32) => void
  readonly getDeferredKernelSyncCountNow: () => number
  readonly setDeferredKernelSyncCountNow: (next: number) => void
  readonly getDeferredKernelSyncEpochNow: () => number
  readonly setDeferredKernelSyncEpochNow: (next: number) => void
  readonly getDeferredKernelSyncSeenNow: () => U32
  readonly setDeferredKernelSyncSeenNow: (next: U32) => void
  readonly getWasmBatchNow: () => U32
  readonly setWasmBatchNow: (next: U32) => void
  readonly getMutationRootsNow: () => U32
  readonly setMutationRootsNow: (next: U32) => void
  readonly getChangedInputEpochNow: () => number
  readonly setChangedInputEpochNow: (next: number) => void
  readonly getChangedInputSeenNow: () => U32
  readonly setChangedInputSeenNow: (next: U32) => void
  readonly getChangedInputBufferNow: () => U32
  readonly setChangedInputBufferNow: (next: U32) => void
  readonly getChangedFormulaEpochNow: () => number
  readonly setChangedFormulaEpochNow: (next: number) => void
  readonly getChangedFormulaSeenNow: () => U32
  readonly setChangedFormulaSeenNow: (next: U32) => void
  readonly getChangedFormulaBufferNow: () => U32
  readonly setChangedFormulaBufferNow: (next: U32) => void
  readonly getChangedUnionEpochNow: () => number
  readonly setChangedUnionEpochNow: (next: number) => void
  readonly getChangedUnionSeenNow: () => U32
  readonly setChangedUnionSeenNow: (next: U32) => void
  readonly getChangedUnionNow: () => U32
  readonly setChangedUnionNow: (next: U32) => void
  readonly getMaterializedCellCountNow: () => number
  readonly setMaterializedCellCountNow: (next: number) => void
  readonly getMaterializedCellsNow: () => U32
  readonly setMaterializedCellsNow: (next: U32) => void
  readonly getExplicitChangedEpochNow: () => number
  readonly setExplicitChangedEpochNow: (next: number) => void
  readonly getExplicitChangedSeenNow: () => U32
  readonly setExplicitChangedSeenNow: (next: U32) => void
  readonly getExplicitChangedBufferNow: () => U32
  readonly setExplicitChangedBufferNow: (next: U32) => void
  readonly getImpactedFormulaEpochNow: () => number
  readonly setImpactedFormulaEpochNow: (next: number) => void
  readonly getImpactedFormulaSeenNow: () => U32
  readonly setImpactedFormulaSeenNow: (next: U32) => void
  readonly getImpactedFormulaBufferNow: () => U32
  readonly setImpactedFormulaBufferNow: (next: U32) => void
}

export function createEngineRuntimeScratchService(): EngineRuntimeScratchService {
  let pendingKernelSync: U32 | undefined
  let deferredKernelSyncCount = 0
  let deferredKernelSyncEpoch = 1
  let deferredKernelSyncSeen: U32 | undefined
  let wasmBatch: U32 | undefined
  let mutationRoots: U32 | undefined
  let changedInputEpoch = 1
  let changedInputSeen: U32 | undefined
  let changedInputBuffer: U32 | undefined
  let changedFormulaEpoch = 1
  let changedFormulaSeen: U32 | undefined
  let changedFormulaBuffer: U32 | undefined
  let changedUnionEpoch = 1
  let changedUnionSeen: U32 | undefined
  let changedUnion: U32 | undefined
  let materializedCellCount = 0
  let materializedCells: U32 | undefined
  let explicitChangedEpoch = 1
  let explicitChangedSeen: U32 | undefined
  let explicitChangedBuffer: U32 | undefined
  let impactedFormulaEpoch = 1
  let impactedFormulaSeen: U32 | undefined
  let impactedFormulaBuffer: U32 | undefined

  const ensureRecalcCapacityNow = (size: number): void => {
    mutationRoots = growScratchBuffer(mutationRoots, size)
    changedInputSeen = growScratchBuffer(changedInputSeen, size)
    changedInputBuffer = growScratchBuffer(changedInputBuffer, size)
    changedFormulaSeen = growScratchBuffer(changedFormulaSeen, size)
    changedFormulaBuffer = growScratchBuffer(changedFormulaBuffer, size)
    pendingKernelSync = growScratchBuffer(pendingKernelSync, size)
    deferredKernelSyncSeen = growScratchBuffer(deferredKernelSyncSeen, size)
    wasmBatch = growScratchBuffer(wasmBatch, size)
    changedUnion = growScratchBuffer(changedUnion, size)
    changedUnionSeen = growScratchBuffer(changedUnionSeen, size)
    materializedCells = growScratchBuffer(materializedCells, size)
    explicitChangedSeen = growScratchBuffer(explicitChangedSeen, size)
    explicitChangedBuffer = growScratchBuffer(explicitChangedBuffer, size)
    impactedFormulaSeen = growScratchBuffer(impactedFormulaSeen, size)
    impactedFormulaBuffer = growScratchBuffer(impactedFormulaBuffer, size)
  }

  return {
    ensureRecalcCapacity(size) {
      return Effect.try({
        try: () => {
          ensureRecalcCapacityNow(size)
        },
        catch: (cause) =>
          new EngineRuntimeScratchError({
            message: scratchErrorMessage('Failed to ensure recalc scratch capacity', cause),
            cause,
          }),
      })
    },
    ensureRecalcCapacityNow,
    getPendingKernelSyncNow: () => (pendingKernelSync ??= createScratchBuffer()),
    setPendingKernelSyncNow: (next) => {
      pendingKernelSync = next
    },
    getDeferredKernelSyncCountNow: () => deferredKernelSyncCount,
    setDeferredKernelSyncCountNow: (next) => {
      deferredKernelSyncCount = next
    },
    getDeferredKernelSyncEpochNow: () => deferredKernelSyncEpoch,
    setDeferredKernelSyncEpochNow: (next) => {
      deferredKernelSyncEpoch = next
    },
    getDeferredKernelSyncSeenNow: () => (deferredKernelSyncSeen ??= createScratchBuffer()),
    setDeferredKernelSyncSeenNow: (next) => {
      deferredKernelSyncSeen = next
    },
    getWasmBatchNow: () => (wasmBatch ??= createScratchBuffer()),
    setWasmBatchNow: (next) => {
      wasmBatch = next
    },
    getMutationRootsNow: () => (mutationRoots ??= createScratchBuffer()),
    setMutationRootsNow: (next) => {
      mutationRoots = next
    },
    getChangedInputEpochNow: () => changedInputEpoch,
    setChangedInputEpochNow: (next) => {
      changedInputEpoch = next
    },
    getChangedInputSeenNow: () => (changedInputSeen ??= createScratchBuffer()),
    setChangedInputSeenNow: (next) => {
      changedInputSeen = next
    },
    getChangedInputBufferNow: () => (changedInputBuffer ??= createScratchBuffer()),
    setChangedInputBufferNow: (next) => {
      changedInputBuffer = next
    },
    getChangedFormulaEpochNow: () => changedFormulaEpoch,
    setChangedFormulaEpochNow: (next) => {
      changedFormulaEpoch = next
    },
    getChangedFormulaSeenNow: () => (changedFormulaSeen ??= createScratchBuffer()),
    setChangedFormulaSeenNow: (next) => {
      changedFormulaSeen = next
    },
    getChangedFormulaBufferNow: () => (changedFormulaBuffer ??= createScratchBuffer()),
    setChangedFormulaBufferNow: (next) => {
      changedFormulaBuffer = next
    },
    getChangedUnionEpochNow: () => changedUnionEpoch,
    setChangedUnionEpochNow: (next) => {
      changedUnionEpoch = next
    },
    getChangedUnionSeenNow: () => (changedUnionSeen ??= createScratchBuffer()),
    setChangedUnionSeenNow: (next) => {
      changedUnionSeen = next
    },
    getChangedUnionNow: () => (changedUnion ??= createScratchBuffer()),
    setChangedUnionNow: (next) => {
      changedUnion = next
    },
    getMaterializedCellCountNow: () => materializedCellCount,
    setMaterializedCellCountNow: (next) => {
      materializedCellCount = next
    },
    getMaterializedCellsNow: () => (materializedCells ??= createScratchBuffer()),
    setMaterializedCellsNow: (next) => {
      materializedCells = next
    },
    getExplicitChangedEpochNow: () => explicitChangedEpoch,
    setExplicitChangedEpochNow: (next) => {
      explicitChangedEpoch = next
    },
    getExplicitChangedSeenNow: () => (explicitChangedSeen ??= createScratchBuffer()),
    setExplicitChangedSeenNow: (next) => {
      explicitChangedSeen = next
    },
    getExplicitChangedBufferNow: () => (explicitChangedBuffer ??= createScratchBuffer()),
    setExplicitChangedBufferNow: (next) => {
      explicitChangedBuffer = next
    },
    getImpactedFormulaEpochNow: () => impactedFormulaEpoch,
    setImpactedFormulaEpochNow: (next) => {
      impactedFormulaEpoch = next
    },
    getImpactedFormulaSeenNow: () => (impactedFormulaSeen ??= createScratchBuffer()),
    setImpactedFormulaSeenNow: (next) => {
      impactedFormulaSeen = next
    },
    getImpactedFormulaBufferNow: () => (impactedFormulaBuffer ??= createScratchBuffer()),
    setImpactedFormulaBufferNow: (next) => {
      impactedFormulaBuffer = next
    },
  }
}
