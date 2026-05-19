import { describe, expect, it } from 'vitest'
import { deferKernelSyncNow } from '../engine/services/live-kernel-sync-state.js'
import { INITIAL_RUNTIME_SCRATCH_CAPACITY, createEngineRuntimeScratchService } from '../engine/services/runtime-scratch-service.js'

describe('live kernel sync state', () => {
  it('defers unique kernel sync cells and preserves existing pending cells', () => {
    const scratch = createEngineRuntimeScratchService()
    scratch.getPendingKernelSyncNow()[0] = 2
    scratch.setDeferredKernelSyncCountNow(1)

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 8,
      cellIndices: Uint32Array.of(2, 5, 5, 7),
      wasmReady: true,
    })

    expect(scratch.getDeferredKernelSyncCountNow()).toBe(3)
    expect(Array.from(scratch.getPendingKernelSyncNow().slice(0, 3))).toEqual([2, 5, 7])
  })

  it('does not advance the epoch for an empty defer request', () => {
    const scratch = createEngineRuntimeScratchService()
    const beforeEpoch = scratch.getDeferredKernelSyncEpochNow()

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 8,
      cellIndices: [],
      wasmReady: true,
    })

    expect(scratch.getDeferredKernelSyncEpochNow()).toBe(beforeEpoch)
    expect(scratch.getDeferredKernelSyncCountNow()).toBe(0)
  })

  it('wraps the deferred seen epoch and clears stale markers', () => {
    const scratch = createEngineRuntimeScratchService()
    scratch.setDeferredKernelSyncEpochNow(0xffff_fffe)
    scratch.getDeferredKernelSyncSeenNow()[3] = 0xffff_fffe

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 8,
      cellIndices: [3],
      wasmReady: true,
    })

    expect(scratch.getDeferredKernelSyncEpochNow()).toBe(1)
    expect(scratch.getDeferredKernelSyncSeenNow()[2]).toBe(0)
    expect(scratch.getDeferredKernelSyncSeenNow()[3]).toBe(1)
    expect(Array.from(scratch.getPendingKernelSyncNow().slice(0, 1))).toEqual([3])
  })

  it('does not grow bitmap or unrelated recalc scratch buffers for single-cell kernel sync defers', () => {
    const scratch = createEngineRuntimeScratchService()
    const deferredSeen = scratch.getDeferredKernelSyncSeenNow()
    const changedInputSeen = scratch.getChangedInputSeenNow()
    const changedFormulaSeen = scratch.getChangedFormulaSeenNow()
    const changedUnionSeen = scratch.getChangedUnionSeenNow()
    const materializedCells = scratch.getMaterializedCellsNow()

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 5001,
      cellIndices: [5000],
      wasmReady: true,
    })

    expect(scratch.getDeferredKernelSyncSeenNow()).toBe(deferredSeen)
    expect(scratch.getPendingKernelSyncNow()).toHaveLength(INITIAL_RUNTIME_SCRATCH_CAPACITY)
    expect(scratch.getDeferredKernelSyncCountNow()).toBe(1)
    expect(scratch.getPendingKernelSyncNow()[0]).toBe(5000)
    expect(scratch.getChangedInputSeenNow()).toBe(changedInputSeen)
    expect(scratch.getChangedFormulaSeenNow()).toBe(changedFormulaSeen)
    expect(scratch.getChangedUnionSeenNow()).toBe(changedUnionSeen)
    expect(scratch.getMaterializedCellsNow()).toBe(materializedCells)
  })

  it('grows only kernel sync state for larger deferred batches', () => {
    const scratch = createEngineRuntimeScratchService()
    const changedInputSeen = scratch.getChangedInputSeenNow()
    const deferredCells = Array.from({ length: 200 }, (_, index) => index + 1000)

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 1200,
      cellIndices: deferredCells,
      wasmReady: true,
    })

    expect(scratch.getPendingKernelSyncNow().length).toBeGreaterThanOrEqual(200)
    expect(scratch.getDeferredKernelSyncSeenNow().length).toBeGreaterThan(1200)
    expect(scratch.getChangedInputSeenNow()).toBe(changedInputSeen)
    expect(scratch.getDeferredKernelSyncCountNow()).toBe(200)
  })

  it('skips deferred kernel sync state entirely before the wasm kernel is ready', () => {
    const scratch = createEngineRuntimeScratchService()
    const pendingKernelSync = scratch.getPendingKernelSyncNow()
    const deferredSeen = scratch.getDeferredKernelSyncSeenNow()

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 5001,
      cellIndices: Array.from({ length: 200 }, (_, index) => index + 1000),
      wasmReady: false,
    })

    expect(scratch.getPendingKernelSyncNow()).toBe(pendingKernelSync)
    expect(scratch.getDeferredKernelSyncSeenNow()).toBe(deferredSeen)
    expect(scratch.getDeferredKernelSyncCountNow()).toBe(0)
  })
})
