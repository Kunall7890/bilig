import { Effect } from 'effect'
import { growUint32 } from '../../engine-buffer-utils.js'
import type { U32 } from '../runtime-state.js'
import { EngineRuntimeScratchError } from '../errors.js'

export const INITIAL_RUNTIME_SCRATCH_CAPACITY = 16

function scratchCapacityFor(required: number): number {
  let capacity = INITIAL_RUNTIME_SCRATCH_CAPACITY
  while (capacity < required) {
    capacity *= 2
  }
  return capacity
}

function createScratchBuffer(required = INITIAL_RUNTIME_SCRATCH_CAPACITY): U32 {
  return new Uint32Array(scratchCapacityFor(required))
}

function growScratchBuffer(buffer: U32 | undefined, size: number): U32 | undefined {
  if (size <= (buffer?.length ?? INITIAL_RUNTIME_SCRATCH_CAPACITY)) {
    return buffer
  }
  if (buffer === undefined) {
    return createScratchBuffer(size)
  }
  return growUint32(buffer, size)
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

class EngineRuntimeScratchServiceImpl implements EngineRuntimeScratchService {
  private pendingKernelSync: U32 | undefined
  private deferredKernelSyncCount = 0
  private deferredKernelSyncEpoch = 1
  private deferredKernelSyncSeen: U32 | undefined
  private wasmBatch: U32 | undefined
  private mutationRoots: U32 | undefined
  private changedInputEpoch = 1
  private changedInputSeen: U32 | undefined
  private changedInputBuffer: U32 | undefined
  private changedFormulaEpoch = 1
  private changedFormulaSeen: U32 | undefined
  private changedFormulaBuffer: U32 | undefined
  private changedUnionEpoch = 1
  private changedUnionSeen: U32 | undefined
  private changedUnion: U32 | undefined
  private materializedCellCount = 0
  private materializedCells: U32 | undefined
  private explicitChangedEpoch = 1
  private explicitChangedSeen: U32 | undefined
  private explicitChangedBuffer: U32 | undefined
  private impactedFormulaEpoch = 1
  private impactedFormulaSeen: U32 | undefined
  private impactedFormulaBuffer: U32 | undefined

  ensureRecalcCapacity(size: number): Effect.Effect<void, EngineRuntimeScratchError> {
    return Effect.try({
      try: () => {
        this.ensureRecalcCapacityNow(size)
      },
      catch: (cause) =>
        new EngineRuntimeScratchError({
          message: scratchErrorMessage('Failed to ensure recalc scratch capacity', cause),
          cause,
        }),
    })
  }

  ensureRecalcCapacityNow(size: number): void {
    this.mutationRoots = growScratchBuffer(this.mutationRoots, size)
    this.changedInputSeen = growScratchBuffer(this.changedInputSeen, size)
    this.changedInputBuffer = growScratchBuffer(this.changedInputBuffer, size)
    this.changedFormulaSeen = growScratchBuffer(this.changedFormulaSeen, size)
    this.changedFormulaBuffer = growScratchBuffer(this.changedFormulaBuffer, size)
    this.pendingKernelSync = growScratchBuffer(this.pendingKernelSync, size)
    this.deferredKernelSyncSeen = growScratchBuffer(this.deferredKernelSyncSeen, size)
    this.wasmBatch = growScratchBuffer(this.wasmBatch, size)
    this.changedUnion = growScratchBuffer(this.changedUnion, size)
    this.changedUnionSeen = growScratchBuffer(this.changedUnionSeen, size)
    this.materializedCells = growScratchBuffer(this.materializedCells, size)
    this.explicitChangedSeen = growScratchBuffer(this.explicitChangedSeen, size)
    this.explicitChangedBuffer = growScratchBuffer(this.explicitChangedBuffer, size)
    this.impactedFormulaSeen = growScratchBuffer(this.impactedFormulaSeen, size)
    this.impactedFormulaBuffer = growScratchBuffer(this.impactedFormulaBuffer, size)
  }

  getPendingKernelSyncNow(): U32 {
    return (this.pendingKernelSync ??= createScratchBuffer())
  }

  setPendingKernelSyncNow(next: U32): void {
    this.pendingKernelSync = next
  }

  getDeferredKernelSyncCountNow(): number {
    return this.deferredKernelSyncCount
  }

  setDeferredKernelSyncCountNow(next: number): void {
    this.deferredKernelSyncCount = next
  }

  getDeferredKernelSyncEpochNow(): number {
    return this.deferredKernelSyncEpoch
  }

  setDeferredKernelSyncEpochNow(next: number): void {
    this.deferredKernelSyncEpoch = next
  }

  getDeferredKernelSyncSeenNow(): U32 {
    return (this.deferredKernelSyncSeen ??= createScratchBuffer())
  }

  setDeferredKernelSyncSeenNow(next: U32): void {
    this.deferredKernelSyncSeen = next
  }

  getWasmBatchNow(): U32 {
    return (this.wasmBatch ??= createScratchBuffer())
  }

  setWasmBatchNow(next: U32): void {
    this.wasmBatch = next
  }

  getMutationRootsNow(): U32 {
    return (this.mutationRoots ??= createScratchBuffer())
  }

  setMutationRootsNow(next: U32): void {
    this.mutationRoots = next
  }

  getChangedInputEpochNow(): number {
    return this.changedInputEpoch
  }

  setChangedInputEpochNow(next: number): void {
    this.changedInputEpoch = next
  }

  getChangedInputSeenNow(): U32 {
    return (this.changedInputSeen ??= createScratchBuffer())
  }

  setChangedInputSeenNow(next: U32): void {
    this.changedInputSeen = next
  }

  getChangedInputBufferNow(): U32 {
    return (this.changedInputBuffer ??= createScratchBuffer())
  }

  setChangedInputBufferNow(next: U32): void {
    this.changedInputBuffer = next
  }

  getChangedFormulaEpochNow(): number {
    return this.changedFormulaEpoch
  }

  setChangedFormulaEpochNow(next: number): void {
    this.changedFormulaEpoch = next
  }

  getChangedFormulaSeenNow(): U32 {
    return (this.changedFormulaSeen ??= createScratchBuffer())
  }

  setChangedFormulaSeenNow(next: U32): void {
    this.changedFormulaSeen = next
  }

  getChangedFormulaBufferNow(): U32 {
    return (this.changedFormulaBuffer ??= createScratchBuffer())
  }

  setChangedFormulaBufferNow(next: U32): void {
    this.changedFormulaBuffer = next
  }

  getChangedUnionEpochNow(): number {
    return this.changedUnionEpoch
  }

  setChangedUnionEpochNow(next: number): void {
    this.changedUnionEpoch = next
  }

  getChangedUnionSeenNow(): U32 {
    return (this.changedUnionSeen ??= createScratchBuffer())
  }

  setChangedUnionSeenNow(next: U32): void {
    this.changedUnionSeen = next
  }

  getChangedUnionNow(): U32 {
    return (this.changedUnion ??= createScratchBuffer())
  }

  setChangedUnionNow(next: U32): void {
    this.changedUnion = next
  }

  getMaterializedCellCountNow(): number {
    return this.materializedCellCount
  }

  setMaterializedCellCountNow(next: number): void {
    this.materializedCellCount = next
  }

  getMaterializedCellsNow(): U32 {
    return (this.materializedCells ??= createScratchBuffer())
  }

  setMaterializedCellsNow(next: U32): void {
    this.materializedCells = next
  }

  getExplicitChangedEpochNow(): number {
    return this.explicitChangedEpoch
  }

  setExplicitChangedEpochNow(next: number): void {
    this.explicitChangedEpoch = next
  }

  getExplicitChangedSeenNow(): U32 {
    return (this.explicitChangedSeen ??= createScratchBuffer())
  }

  setExplicitChangedSeenNow(next: U32): void {
    this.explicitChangedSeen = next
  }

  getExplicitChangedBufferNow(): U32 {
    return (this.explicitChangedBuffer ??= createScratchBuffer())
  }

  setExplicitChangedBufferNow(next: U32): void {
    this.explicitChangedBuffer = next
  }

  getImpactedFormulaEpochNow(): number {
    return this.impactedFormulaEpoch
  }

  setImpactedFormulaEpochNow(next: number): void {
    this.impactedFormulaEpoch = next
  }

  getImpactedFormulaSeenNow(): U32 {
    return (this.impactedFormulaSeen ??= createScratchBuffer())
  }

  setImpactedFormulaSeenNow(next: U32): void {
    this.impactedFormulaSeen = next
  }

  getImpactedFormulaBufferNow(): U32 {
    return (this.impactedFormulaBuffer ??= createScratchBuffer())
  }

  setImpactedFormulaBufferNow(next: U32): void {
    this.impactedFormulaBuffer = next
  }
}

export function createEngineRuntimeScratchService(): EngineRuntimeScratchService {
  return new EngineRuntimeScratchServiceImpl()
}
