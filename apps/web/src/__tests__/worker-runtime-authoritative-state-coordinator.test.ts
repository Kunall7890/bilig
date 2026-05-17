import { SpreadsheetEngine } from '@bilig/core'
import type { EngineReplicaSnapshot } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { WorkerRuntimeAuthoritativeStateCoordinator } from '../worker-runtime-authoritative-state-coordinator.js'
import { WorkerRuntimeSnapshotCaches } from '../worker-runtime-snapshot-caches.js'

function createAuthoritativeEngine(value: string): SpreadsheetEngine {
  const engine = new SpreadsheetEngine({ workbookName: 'authoritative-state-doc' })
  engine.setCellValue('Sheet1', 'A1', value)
  return engine
}

describe('WorkerRuntimeAuthoritativeStateCoordinator', () => {
  it('keeps bootstrap revisions monotonic but lets reconcile install the server revision', () => {
    const coordinator = new WorkerRuntimeAuthoritativeStateCoordinator({
      snapshotCaches: new WorkerRuntimeSnapshotCaches(),
      getDocumentId: () => 'authoritative-state-doc',
      getReplicaId: () => 'replica-1',
      listPendingMutations: () => [],
    })

    coordinator.acceptSnapshotRevision(12, 'bootstrap')
    coordinator.acceptSnapshotRevision(8, 'bootstrap')
    expect(coordinator.getRevision()).toBe(12)

    coordinator.acceptSnapshotRevision(9, 'reconcile')
    expect(coordinator.getRevision()).toBe(9)
  })

  it('installs restored authoritative snapshot state without creating an eager authoritative engine', async () => {
    const caches = new WorkerRuntimeSnapshotCaches()
    const restoredEngine = createAuthoritativeEngine('restored')
    const snapshot = restoredEngine.exportSnapshot()
    const replica: EngineReplicaSnapshot = restoredEngine.exportReplicaSnapshot()
    const coordinator = new WorkerRuntimeAuthoritativeStateCoordinator({
      snapshotCaches: caches,
      getDocumentId: () => 'authoritative-state-doc',
      getReplicaId: () => 'replica-1',
      listPendingMutations: () => [],
    })

    coordinator.installPreparedState({
      authoritativeEngine: null,
      authoritativeSnapshot: snapshot,
      authoritativeReplica: replica,
    })

    expect(coordinator.hasMaterializedEngine()).toBe(false)
    expect(await coordinator.resolveStateInput()).toEqual({ snapshot, replica })

    const engine = await coordinator.getEngine()
    expect(coordinator.hasMaterializedEngine()).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: 'restored',
      stringId: expect.any(Number),
    })
  })

  it('rebuilds projection state from the latest exported authoritative engine when caches are invalidated', async () => {
    const caches = new WorkerRuntimeSnapshotCaches()
    const authoritativeEngine = createAuthoritativeEngine('before')
    const coordinator = new WorkerRuntimeAuthoritativeStateCoordinator({
      snapshotCaches: caches,
      getDocumentId: () => 'authoritative-state-doc',
      getReplicaId: () => 'replica-1',
      listPendingMutations: () => [],
    })

    coordinator.installPreparedState({
      authoritativeEngine,
      authoritativeSnapshot: authoritativeEngine.exportSnapshot(),
      authoritativeReplica: authoritativeEngine.exportReplicaSnapshot(),
    })

    authoritativeEngine.setCellValue('Sheet1', 'A1', 'after')
    coordinator.invalidateCachedState()

    const { engine } = await coordinator.rebuildProjectionEngine()

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: 'after',
      stringId: expect.any(Number),
    })
  })
})
