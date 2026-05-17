import { afterEach, describe, expect, it, vi } from 'vitest'
import { createWorkerEngineHost } from '@bilig/worker-transport'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { createEmptyWorkbookSnapshot } from '@bilig/zero-sync'
import { createWorkerRuntimeSessionController, type WorkerRuntimeSessionController } from '../runtime-session.js'
import type { WorkbookWorkerStateSnapshot } from '../worker-runtime.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'

const metrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function runtimeState(overrides: Partial<WorkbookWorkerStateSnapshot> = {}): WorkbookWorkerStateSnapshot {
  return {
    workbookName: 'doc-1',
    sheets: [{ id: 1, name: 'Sheet1', order: 0 }],
    sheetNames: ['Sheet1'],
    definedNames: [],
    metrics,
    syncState: 'live',
    localHistoryState: {
      canUndo: false,
      canRedo: false,
    },
    authoritativeRevision: 3,
    pendingMutationSummary: {
      activeCount: 1,
      failedCount: 0,
      firstFailed: null,
    },
    localPersistenceMode: 'ephemeral',
    ...overrides,
  }
}

class FakeRevisionLiveView {
  readonly listeners = new Set<(value: unknown) => void>()
  destroy = vi.fn()

  constructor(readonly data: unknown) {}

  addListener(listener: (value: unknown) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  emit(value: unknown): void {
    this.listeners.forEach((listener) => {
      listener(value)
    })
  }
}

describe('worker runtime session authoritative event loading', () => {
  let channel: MessageChannel | null = null
  let controller: WorkerRuntimeSessionController | null = null
  let host: { dispose(): void } | null = null

  afterEach(() => {
    controller?.dispose()
    host?.dispose()
    channel?.port1.close()
    channel?.port2.close()
    channel = null
    controller = null
    host = null
  })

  it('rejects authoritative event responses whose cursor does not match the requested revision', async () => {
    channel = new MessageChannel()
    const applyAuthoritativeEvents = vi.fn()
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: runtimeState(),
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 3
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        applyAuthoritativeEvents,
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(
      async () =>
        new Response(
          JSON.stringify({
            afterRevision: 0,
            headRevision: 5,
            calculatedRevision: 5,
            events: [
              {
                revision: 1,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A1',
                  value: 1,
                },
              },
              {
                revision: 2,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A2',
                  value: 2,
                },
              },
              {
                revision: 3,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A3',
                  value: 3,
                },
              },
              {
                revision: 4,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A4',
                  value: 4,
                },
              },
              {
                revision: 5,
                clientMutationId: null,
                payload: {
                  kind: 'setCellValue',
                  sheetName: 'Sheet1',
                  address: 'A5',
                  value: 5,
                },
              },
            ],
          }),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        ),
    )
    const onError = vi.fn()

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        fetchImpl,
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError,
      },
    )

    await expect(controller.invoke('refreshAuthoritativeEvents')).rejects.toThrow(
      'Authoritative event payload does not match the expected schema',
    )
    expect(fetchImpl).toHaveBeenCalledWith('/v2/documents/doc-1/events?afterRevision=3', {
      headers: {
        accept: 'application/json',
      },
      cache: 'no-store',
    })
    expect(applyAuthoritativeEvents).not.toHaveBeenCalled()
    expect(onError).toHaveBeenCalledWith('Authoritative event payload does not match the expected schema')
  })

  it('refreshes a same-head authoritative snapshot when calculation catches up', async () => {
    channel = new MessageChannel()
    const revisionView = new FakeRevisionLiveView({
      headRevision: 5,
      calculatedRevision: 4,
    })
    const installAuthoritativeSnapshot = vi.fn((input: { authoritativeRevision: number }) =>
      runtimeState({
        authoritativeRevision: input.authoritativeRevision,
        pendingMutationSummary: {
          activeCount: 0,
          failedCount: 0,
          firstFailed: null,
        },
      }),
    )
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: runtimeState({
              authoritativeRevision: 0,
              pendingMutationSummary: {
                activeCount: 0,
                failedCount: 0,
                firstFailed: null,
              },
            }),
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getCell(sheetName: string, address: string) {
          return {
            sheetName,
            address,
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          }
        },
        installAuthoritativeSnapshot,
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )
    const fetchImpl = vi.fn(async () => {
      return new Response(JSON.stringify(createEmptyWorkbookSnapshot('doc-1')), {
        status: 200,
        headers: {
          'content-type': 'application/json',
          'x-bilig-snapshot-cursor': '5',
        },
      })
    })

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        fetchImpl,
        zero: {
          materialize: vi.fn(() => revisionView),
        },
        initialSelection: { sheetName: 'Sheet1', address: 'A1' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError: vi.fn(),
      },
    )

    await vi.waitFor(() => {
      expect(installAuthoritativeSnapshot).toHaveBeenCalledTimes(1)
    })
    expect(installAuthoritativeSnapshot.mock.calls[0]?.[0]).toMatchObject({
      authoritativeRevision: 5,
      mode: 'bootstrap',
    })

    revisionView.emit({
      headRevision: 5,
      calculatedRevision: 5,
    })

    await vi.waitFor(() => {
      expect(installAuthoritativeSnapshot).toHaveBeenCalledTimes(2)
    })
    expect(installAuthoritativeSnapshot.mock.calls[1]?.[0]).toMatchObject({
      authoritativeRevision: 5,
      mode: 'reconcile',
    })
  })

  it('hydrates the selected cell after local undo supersedes an optimistic clear', async () => {
    channel = new MessageChannel()
    let selectedCell: CellSnapshot = {
      sheetName: 'Sheet1',
      address: 'D12',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    }
    let state = runtimeState({
      syncState: 'local-only',
      localHistoryState: { canUndo: true, canRedo: false },
      pendingMutationSummary: {
        activeCount: 2,
        failedCount: 0,
        firstFailed: null,
      },
    })
    host = createWorkerEngineHost(
      {
        async bootstrap() {
          return {
            runtimeState: state,
            restoredFromPersistence: true,
            requiresAuthoritativeHydrate: false,
            localPersistenceMode: 'ephemeral',
          }
        },
        getAuthoritativeRevision() {
          return 0
        },
        getRuntimeState() {
          return state
        },
        getCell(sheetName: string, address: string) {
          return selectedCell.sheetName === sheetName && selectedCell.address === address
            ? selectedCell
            : {
                sheetName,
                address,
                value: { tag: ValueTag.Empty },
                flags: 0,
                version: 0,
              }
        },
        undoLocalChange() {
          selectedCell = {
            sheetName: 'Sheet1',
            address: 'D12',
            value: { tag: ValueTag.String, value: 'delete-undo-redo', stringId: 1 },
            input: 'delete-undo-redo',
            flags: 0,
            version: 1,
          }
          state = runtimeState({
            syncState: 'local-only',
            localHistoryState: { canUndo: true, canRedo: true },
            pendingMutationSummary: {
              activeCount: 2,
              failedCount: 0,
              firstFailed: null,
            },
          })
          return true
        },
        subscribeViewportPatches() {
          return () => undefined
        },
      },
      channel.port1,
    )

    controller = await createWorkerRuntimeSessionController(
      {
        documentId: 'doc-1',
        replicaId: 'browser:test',
        persistState: false,
        authoritativeSyncEnabled: false,
        initialSelection: { sheetName: 'Sheet1', address: 'D12' },
        createWorker: () => channel!.port2,
      },
      {
        onRuntimeState: vi.fn(),
        onSelection: vi.fn(),
        onError: vi.fn(),
      },
    )
    controller.handle.viewportStore.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D12',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 2,
    })

    await controller.invoke('undoLocalChange')

    expect(controller.handle.viewportStore.getCell('Sheet1', 'D12')).toMatchObject({
      input: 'delete-undo-redo',
      value: { tag: ValueTag.String, value: 'delete-undo-redo', stringId: 1 },
      flags: 0,
      version: 1,
    })
  })
})
