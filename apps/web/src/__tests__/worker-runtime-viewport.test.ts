import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot, type CellStyleRecord, type EngineEvent, type RecalcMetrics } from '@bilig/protocol'
import { decodeViewportPatch, type ViewportPatch } from '@bilig/worker-transport'
import { WorkbookWorkerRuntime } from '../worker-runtime.js'
import { buildViewportPatchFromEngine } from '../worker-runtime-viewport.js'
import { WorkerViewportPatchPublisher } from '../worker-runtime-viewport-publisher.js'
import type { ViewportSubscriptionState, WorkerEngine } from '../worker-runtime-support.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 1,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

const STYLE_ID = 'style-live'
const CELL: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'B2',
  value: { tag: ValueTag.Empty },
  flags: 0,
  styleId: STYLE_ID,
  version: 1,
}

function createStyle(backgroundColor: string): CellStyleRecord {
  return {
    id: STYLE_ID,
    fill: { backgroundColor },
  }
}

function createEvent(): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: new Uint32Array(),
    changedCells: [],
    invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }],
    invalidatedRows: [],
    invalidatedColumns: [],
    metrics: TEST_METRICS,
  }
}

function createSubscriptionState(): ViewportSubscriptionState {
  return {
    subscription: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 5,
      colStart: 0,
      colEnd: 5,
    },
    listener: vi.fn(),
    nextVersion: 1,
    knownStyleIds: new Set(),
    lastStyleSignatures: new Map(),
    lastCellSignatures: new Map(),
    lastColumnSignatures: new Map(),
    lastRowSignatures: new Map(),
    lastMergeSignatures: new Map(),
  }
}

function createTestEngine(
  input: {
    readonly hasSheet?: () => boolean
    readonly getCell?: (sheetName: string, address: string) => CellSnapshot
    readonly getCellStyle?: (styleId: string | undefined) => CellStyleRecord | undefined
  } = {},
): WorkerEngine {
  return {
    workbook: {
      workbookName: 'viewport-style-proof',
      cellStore: {
        sheetIds: new Uint16Array(),
        rows: new Uint32Array(),
        cols: new Uint16Array(),
      },
      sheetsByName: new Map(),
      getSheet: (sheetName) =>
        input.hasSheet?.() === false ? undefined : { name: sheetName, order: 0, grid: { forEachCellEntry: () => undefined } },
      getSheetNameById: () => 'Sheet1',
      getQualifiedAddress: () => 'Sheet1!B2',
    },
    ready: async () => undefined,
    createSheet: () => undefined,
    subscribe: () => () => undefined,
    subscribeBatches: () => () => undefined,
    getLastMetrics: () => TEST_METRICS,
    getSyncState: () => 'local-only',
    getCell: (sheetName, address) => input.getCell?.(sheetName, address) ?? CELL,
    getCellStyle: (styleId) => input.getCellStyle?.(styleId) ?? createStyle('#00ff00'),
    setRangeNumberFormat: () => undefined,
    clearRangeNumberFormat: () => undefined,
    clearRange: () => undefined,
    setCellValue: () => undefined,
    setCellFormula: () => undefined,
    setRangeStyle: () => undefined,
    clearRangeStyle: () => undefined,
    clearCell: () => undefined,
    undo: () => false,
    redo: () => false,
    canUndo: () => false,
    canRedo: () => false,
    renderCommit: () => undefined,
    fillRange: () => undefined,
    copyRange: () => undefined,
    moveRange: () => undefined,
    insertRows: () => undefined,
    deleteRows: () => undefined,
    insertColumns: () => undefined,
    deleteColumns: () => undefined,
    updateRowMetadata: () => undefined,
    updateColumnMetadata: () => undefined,
    setFreezePane: () => undefined,
    getFreezePane: () => undefined,
    mergeCells: () => undefined,
    unmergeCells: () => false,
    getMergeRange: () => undefined,
    listMergeRanges: () => [],
    exportSnapshot: () => ({
      version: 1,
      workbook: { name: 'viewport-style-proof' },
      sheets: [{ name: 'Sheet1', order: 0, cells: [] }],
    }),
    exportReplicaSnapshot: () => ({
      replica: {
        replicaId: 'viewport-style-proof',
        counter: 0,
        appliedBatchIds: [],
      },
      entityVersions: [],
      sheetDeleteVersions: [],
    }),
    importSnapshot: () => undefined,
    importReplicaSnapshot: () => undefined,
    getColumnAxisEntries: () => [],
    getRowAxisEntries: () => [],
  }
}

describe('worker runtime viewport patches', () => {
  it('returns a typed sheet-id view window for authoritative visible cells', async () => {
    const runtime = new WorkbookWorkerRuntime()
    await runtime.bootstrap({
      documentId: 'viewport-window-ready',
      replicaId: 'replica-1',
      persistState: false,
    })
    await runtime.ready()
    await runtime.setCellValue('Sheet1', 'B2', 'ready-window')

    const window = runtime.getWorkbookViewWindow({
      sheetId: 1,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      colStart: 0,
      colEnd: 2,
    })

    expect(window).toMatchObject({
      status: 'ready',
      request: {
        sheetId: 1,
        sheetName: 'Sheet1',
      },
      sheet: {
        sheetId: 1,
        sheetName: 'Sheet1',
      },
      renderAck: {
        status: 'not-requested',
        reason: 'authoritative-window-built-before-browser-render-ack',
      },
    })
    expect(window.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          row: 1,
          col: 1,
          address: 'B2',
          displayText: 'ready-window',
          editorText: 'ready-window',
        }),
      ]),
    )

    runtime.dispose()
  })

  it('returns a typed unavailable view window instead of succeeding with a blank fallback sheet', async () => {
    const runtime = new WorkbookWorkerRuntime()
    await runtime.bootstrap({
      documentId: 'viewport-window-missing-sheet',
      replicaId: 'replica-1',
      persistState: false,
    })
    await runtime.ready()

    expect(
      runtime.getWorkbookViewWindow({
        sheetId: 9999,
        sheetName: 'Prepaid Template',
        rowStart: 10,
        rowEnd: 12,
        colStart: 1,
        colEnd: 3,
      }),
    ).toMatchObject({
      status: 'missing-sheet',
      reason: 'sheet-id-not-found',
      sheet: null,
      cells: [],
      renderAck: {
        status: 'rejected',
        reason: 'sheet-id-not-found',
      },
    })

    expect(
      runtime.getWorkbookViewWindow({
        sheetId: 1,
        sheetName: 'Prepaid Template',
        rowStart: 10,
        rowEnd: 12,
        colStart: 1,
        colEnd: 3,
      }),
    ).toMatchObject({
      status: 'identity-mismatch',
      reason: 'sheet-name-does-not-match-sheet-id',
      sheet: {
        sheetId: 1,
        sheetName: 'Sheet1',
      },
      cells: [],
    })

    runtime.dispose()
  })

  it('carries sheet identity through runtime viewport patches', async () => {
    const runtime = new WorkbookWorkerRuntime()
    await runtime.bootstrap({
      documentId: 'viewport-patch-sheet-id',
      replicaId: 'replica-1',
      persistState: false,
    })
    await runtime.ready()

    const patches: ViewportPatch[] = []
    const unsubscribe = runtime.subscribeViewportPatches(
      {
        sheetId: 1,
        sheetName: 'Sheet1',
        sheetOrdinal: 0,
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      (bytes) => {
        patches.push(decodeViewportPatch(bytes))
      },
    )

    expect(patches.at(0)?.viewport).toMatchObject({
      sheetId: 1,
      sheetName: 'Sheet1',
      sheetOrdinal: 0,
    })

    unsubscribe()
    runtime.dispose()
  })

  it('rejects sheet-id mismatched viewport subscriptions instead of publishing fallback cells', async () => {
    const runtime = new WorkbookWorkerRuntime()
    await runtime.bootstrap({
      documentId: 'viewport-patch-sheet-id-mismatch',
      replicaId: 'replica-1',
      persistState: false,
    })
    await runtime.ready()

    const listener = vi.fn()
    const unsubscribe = runtime.subscribeViewportPatches(
      {
        sheetId: 999,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      listener,
    )

    await runtime.setCellValue('Sheet1', 'A1', 'must-not-leak')

    expect(listener).not.toHaveBeenCalled()

    unsubscribe()
    runtime.dispose()
  })

  it('refuses to fabricate a full empty viewport patch when the sheet is missing', () => {
    const state = {
      ...createSubscriptionState(),
      subscription: {
        ...createSubscriptionState().subscription,
        sheetName: 'Missing',
      },
    }
    const emptyCellSnapshot = vi.fn((sheetName: string, address: string): CellSnapshot => ({ ...CELL, sheetName, address }))
    const engine = createTestEngine({
      hasSheet: () => false,
      getCell() {
        throw new Error('Missing sheet viewport patches must not read cells')
      },
    })

    expect(() =>
      buildViewportPatchFromEngine({
        state,
        event: null,
        metrics: TEST_METRICS,
        authoritativeRevision: 7,
        sheetImpact: null,
        engine,
        emptyCellSnapshot,
        getStyleRecord: createStyle,
        getFormatId: () => 0,
      }),
    ).toThrow('Cannot build viewport patch for missing sheet: Missing')
    expect(emptyCellSnapshot).not.toHaveBeenCalled()
  })

  it('includes affected cells when a referenced style record changes under the same style id', () => {
    let currentStyle = createStyle('#00ff00')
    const engine = createTestEngine({
      getCell: () => CELL,
      getCellStyle: () => currentStyle,
    })
    const publisher = new WorkerViewportPatchPublisher({
      buildPatch: () => {
        throw new Error('test uses publisher.buildPatch directly')
      },
      getAuthoritativeRevision: () => 0,
      getCurrentMetrics: () => TEST_METRICS,
      getProjectionEngine: () => engine,
      hasProjectionEngine: () => true,
    })
    const state = createSubscriptionState()

    const initialPatch = publisher.buildPatch(state, null, TEST_METRICS, 0, null)
    expect(initialPatch.cells.map((cell) => cell.snapshot.address)).toContain('B2')
    expect(initialPatch.styles).toContainEqual(createStyle('#00ff00'))

    currentStyle = createStyle('#0000ff')
    const styleOnlyEventPatch = publisher.buildPatch(state, createEvent(), TEST_METRICS, 0, {
      changedCells: null,
      invalidatedRanges: [{ rowStart: 1, rowEnd: 1, colStart: 1, colEnd: 1 }],
      invalidatedRows: [],
      invalidatedColumns: [],
    })

    expect(styleOnlyEventPatch.styles).toEqual([createStyle('#0000ff')])
    expect(styleOnlyEventPatch.cells).toEqual([
      expect.objectContaining({
        row: 1,
        col: 1,
        snapshot: expect.objectContaining({
          address: 'B2',
          styleId: STYLE_ID,
        }),
        styleId: STYLE_ID,
      }),
    ])
  })

  it('does not publish synthetic empty viewport patches for a missing sheet', () => {
    let hasSheet = false
    const engine = createTestEngine({
      hasSheet: () => hasSheet,
      getCell: (sheetName, address) => ({ ...CELL, sheetName, address }),
    })
    const buildPatch = vi.fn(
      (
        state: ViewportSubscriptionState,
        _event: EngineEvent | null,
        metrics: RecalcMetrics,
        authoritativeRevision: number,
      ): ViewportPatch => ({
        version: state.nextVersion++,
        authoritativeRevision,
        full: true,
        freezeRows: 0,
        freezeCols: 0,
        viewport: state.subscription,
        metrics,
        styles: [],
        cells: [
          {
            row: state.subscription.rowStart,
            col: state.subscription.colStart,
            snapshot: { ...CELL, sheetName: state.subscription.sheetName, address: 'A1' },
            displayText: '',
            copyText: '',
            editorText: '',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      }),
    )
    const listener = vi.fn()
    const publisher = new WorkerViewportPatchPublisher({
      buildPatch,
      getAuthoritativeRevision: () => 7,
      getCurrentMetrics: () => TEST_METRICS,
      getProjectionEngine: () => engine,
      hasProjectionEngine: () => true,
    })

    const unsubscribe = publisher.subscribe(
      {
        sheetName: 'Missing',
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      listener,
    )

    expect(buildPatch).not.toHaveBeenCalled()
    expect(listener).not.toHaveBeenCalled()

    publisher.broadcast({ event: null, impactsBySheet: null, metrics: TEST_METRICS })

    expect(buildPatch).not.toHaveBeenCalled()
    expect(listener).not.toHaveBeenCalled()

    hasSheet = true
    publisher.broadcast({ event: null, impactsBySheet: null, metrics: TEST_METRICS })

    expect(buildPatch).toHaveBeenCalledTimes(1)
    expect(listener).toHaveBeenCalledTimes(1)

    unsubscribe()
  })

  it('does not rebroadcast a full viewport after journaling a pending cell mutation', async () => {
    const runtime = new WorkbookWorkerRuntime()
    await runtime.bootstrap({
      documentId: 'viewport-pending-mutation',
      replicaId: 'replica-1',
      persistState: false,
    })
    await runtime.ready()

    const patches: ViewportPatch[] = []
    const unsubscribe = runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 5,
        colStart: 0,
        colEnd: 5,
        initialPatch: 'none',
      },
      (bytes) => {
        patches.push(decodeViewportPatch(bytes))
      },
    )

    await runtime.enqueuePendingMutation({
      method: 'setRangeStyle',
      args: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, { fill: { backgroundColor: '#34a853' } }],
    })

    expect(patches.length).toBeGreaterThan(0)
    expect(patches).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          full: false,
          cells: [
            expect.objectContaining({
              row: 1,
              col: 1,
            }),
          ],
        }),
      ]),
    )
    expect(patches.filter((patch) => patch.full)).toEqual([])

    unsubscribe()
    runtime.dispose()
  })
})
