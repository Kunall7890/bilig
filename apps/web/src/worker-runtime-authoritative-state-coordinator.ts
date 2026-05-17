import type { EngineReplicaSnapshot, SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import type { PendingWorkbookMutation } from './workbook-sync.js'
import type { ProjectionOverlayScope } from './worker-local-overlay.js'
import {
  ensureAuthoritativeEngine,
  installAuthoritativeEngineState,
  installRestoredAuthoritativeState,
  rebuildProjectionEngine,
  resolveAuthoritativeStateInput,
  type WorkerRuntimeAuthoritativeStateInput,
} from './worker-runtime-engine-access.js'
import type { InstallAuthoritativeSnapshotInput } from './worker-runtime-state.js'
import type { WorkerRuntimeSnapshotCaches } from './worker-runtime-snapshot-caches.js'

export interface PreparedWorkerRuntimeAuthoritativeState {
  readonly authoritativeEngine: SpreadsheetEngine | null
  readonly authoritativeSnapshot: WorkbookSnapshot | null
  readonly authoritativeReplica: EngineReplicaSnapshot | null
}

export class WorkerRuntimeAuthoritativeStateCoordinator {
  private authoritativeEngine: SpreadsheetEngine | null = null
  private authoritativeStateSource: 'none' | 'memory' = 'none'
  private authoritativeRevision = 0

  constructor(
    private readonly options: {
      readonly snapshotCaches: WorkerRuntimeSnapshotCaches
      readonly getDocumentId: () => string
      readonly getReplicaId: () => string
      readonly listPendingMutations: () => readonly PendingWorkbookMutation[]
    },
  ) {}

  reset(): void {
    this.authoritativeEngine = null
    this.authoritativeStateSource = 'none'
    this.authoritativeRevision = 0
  }

  getRevision(): number {
    return this.authoritativeRevision
  }

  acceptSnapshotRevision(authoritativeRevision: number, mode: InstallAuthoritativeSnapshotInput['mode']): number {
    this.authoritativeRevision = mode === 'bootstrap' ? Math.max(this.authoritativeRevision, authoritativeRevision) : authoritativeRevision
    return this.authoritativeRevision
  }

  acceptEventRevision(authoritativeRevision: number): number {
    this.authoritativeRevision = Math.max(this.authoritativeRevision, authoritativeRevision)
    return this.authoritativeRevision
  }

  hasMaterializedEngine(): boolean {
    return this.authoritativeEngine !== null
  }

  installPreparedState(state: PreparedWorkerRuntimeAuthoritativeState): void {
    this.authoritativeStateSource = 'memory'
    if (state.authoritativeEngine) {
      this.authoritativeEngine = state.authoritativeEngine
      installAuthoritativeEngineState(
        this.options.snapshotCaches,
        state.authoritativeEngine,
        state.authoritativeSnapshot,
        state.authoritativeReplica,
      )
      return
    }
    this.authoritativeEngine = null
    installRestoredAuthoritativeState(this.options.snapshotCaches, state.authoritativeSnapshot, state.authoritativeReplica)
  }

  invalidateCachedState(): void {
    this.options.snapshotCaches.invalidateAuthoritativeState()
  }

  async resolveStateInput(): Promise<WorkerRuntimeAuthoritativeStateInput> {
    return await resolveAuthoritativeStateInput({
      authoritativeStateSource: this.authoritativeStateSource,
      snapshotCaches: this.options.snapshotCaches,
      authoritativeEngine: this.authoritativeEngine,
    })
  }

  async getEngine(): Promise<SpreadsheetEngine> {
    const engine = await ensureAuthoritativeEngine({
      authoritativeEngine: this.authoritativeEngine,
      documentId: this.options.getDocumentId(),
      replicaId: this.options.getReplicaId(),
      snapshotCaches: this.options.snapshotCaches,
      resolveAuthoritativeStateInput: () => this.resolveStateInput(),
    })
    this.authoritativeEngine = engine
    return engine
  }

  async rebuildProjectionEngine(): Promise<{
    engine: SpreadsheetEngine
    overlayScope: ProjectionOverlayScope | null
  }> {
    return await rebuildProjectionEngine({
      documentId: this.options.getDocumentId(),
      replicaId: this.options.getReplicaId(),
      pendingMutations: this.options.listPendingMutations(),
      resolveAuthoritativeStateInput: () => this.resolveStateInput(),
    })
  }
}
