import type { EngineRuntimeScratchService } from './runtime-scratch-service.js'
import type { U32 } from '../runtime-state.js'
import { growUint32 } from '../../engine-buffer-utils.js'

const LINEAR_KERNEL_SYNC_DEDUPE_LIMIT = 64

export function deferKernelSyncNow(args: {
  readonly scratch: EngineRuntimeScratchService
  readonly cellStoreSize: number
  readonly cellIndices: readonly number[] | U32
  readonly wasmReady: boolean
}): void {
  if (!args.wasmReady || args.cellIndices.length === 0) {
    return
  }
  let pendingKernelSync = args.scratch.getPendingKernelSyncNow()
  let deferredCount = args.scratch.getDeferredKernelSyncCountNow()
  const requiredPendingLength = deferredCount + args.cellIndices.length
  if (requiredPendingLength > pendingKernelSync.length) {
    pendingKernelSync = growUint32(pendingKernelSync, requiredPendingLength)
    args.scratch.setPendingKernelSyncNow(pendingKernelSync)
  }
  if (requiredPendingLength <= LINEAR_KERNEL_SYNC_DEDUPE_LIMIT && args.scratch.getDeferredKernelSyncEpochNow() < 0xffff_fffe) {
    for (let index = 0; index < args.cellIndices.length; index += 1) {
      const cellIndex = args.cellIndices[index]!
      let alreadyPending = false
      for (let pendingIndex = 0; pendingIndex < deferredCount; pendingIndex += 1) {
        if (pendingKernelSync[pendingIndex] === cellIndex) {
          alreadyPending = true
          break
        }
      }
      if (!alreadyPending) {
        pendingKernelSync[deferredCount] = cellIndex
        deferredCount += 1
      }
    }
    args.scratch.setDeferredKernelSyncCountNow(deferredCount)
    return
  }
  let deferredEpoch = args.scratch.getDeferredKernelSyncEpochNow() + 1
  let deferredSeen = args.scratch.getDeferredKernelSyncSeenNow()
  const requiredSeenLength = args.cellStoreSize + 1
  if (requiredSeenLength > deferredSeen.length) {
    deferredSeen = growUint32(deferredSeen, requiredSeenLength)
    args.scratch.setDeferredKernelSyncSeenNow(deferredSeen)
  }
  if (deferredEpoch === 0xffff_ffff) {
    deferredEpoch = 1
    deferredSeen.fill(0)
  }
  args.scratch.setDeferredKernelSyncEpochNow(deferredEpoch)
  for (let index = 0; index < deferredCount; index += 1) {
    const cellIndex = pendingKernelSync[index]
    if (cellIndex !== undefined) {
      deferredSeen[cellIndex] = deferredEpoch
    }
  }
  for (let index = 0; index < args.cellIndices.length; index += 1) {
    const cellIndex = args.cellIndices[index]!
    if (deferredSeen[cellIndex] === deferredEpoch) {
      continue
    }
    deferredSeen[cellIndex] = deferredEpoch
    pendingKernelSync[deferredCount] = cellIndex
    deferredCount += 1
  }
  args.scratch.setDeferredKernelSyncCountNow(deferredCount)
}
