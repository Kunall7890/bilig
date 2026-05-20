import { useCallback, useEffect, useLayoutEffect, useRef, useState, type MutableRefObject } from 'react'
import type { GridSelectionSnapshot } from '@bilig/grid'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  stringifyWorkbookAgentUiContextRenderedProofKey,
  stringifyWorkbookAgentUiContextSemanticKey,
  type WorkbookAgentRenderedRange,
  type WorkbookAgentUiContext,
} from '@bilig/contracts'
import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type CellRangeRef, type Viewport } from '@bilig/protocol'
import { buildWorkbookAgentContext } from './workbook-agent-context.js'
import type { ProjectedViewportStore } from './projected-viewport-store.js'
import type { WorkerRuntimeSelection, WorkerRuntimeSessionController } from './runtime-session.js'

function selectionViewport(selection: WorkerRuntimeSelection): Viewport {
  const parsed = parseCellAddress(selection.address, selection.sheetName)
  return {
    rowStart: parsed.row,
    rowEnd: parsed.row,
    colStart: parsed.col,
    colEnd: parsed.col,
  }
}

function viewportContains(outer: Viewport, inner: Viewport): boolean {
  return (
    outer.rowStart <= inner.rowStart && outer.rowEnd >= inner.rowEnd && outer.colStart <= inner.colStart && outer.colEnd >= inner.colEnd
  )
}

function expandProjectionViewport(viewport: Viewport): Viewport {
  const rowPadding = VIEWPORT_TILE_ROW_COUNT * 3
  const colPadding = VIEWPORT_TILE_COLUMN_COUNT
  return {
    rowStart: Math.max(0, viewport.rowStart - rowPadding),
    rowEnd: Math.min(MAX_ROWS - 1, viewport.rowEnd + rowPadding),
    colStart: Math.max(0, viewport.colStart - colPadding),
    colEnd: Math.min(MAX_COLS - 1, viewport.colEnd + colPadding),
  }
}

const MAX_AGENT_RENDERED_CONTEXT_CELLS = 200

function normalizeCellRangeRef(range: CellRangeRef): CellRangeRef & {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

function buildRenderedRangeSnapshot(
  viewportStore: ProjectedViewportStore | null | undefined,
  range: CellRangeRef,
): WorkbookAgentRenderedRange | null {
  if (!viewportStore) {
    return null
  }
  const normalized = normalizeCellRangeRef(range)
  const rowCount = normalized.endRow - normalized.startRow + 1
  const columnCount = normalized.endCol - normalized.startCol + 1
  const cellCount = rowCount * columnCount
  const rows: Array<Array<WorkbookAgentRenderedRange['rows'][number][number]>> = []
  let emittedCells = 0
  for (let row = normalized.startRow; row <= normalized.endRow && emittedCells < MAX_AGENT_RENDERED_CONTEXT_CELLS; row += 1) {
    const rowEntries: Array<WorkbookAgentRenderedRange['rows'][number][number]> = []
    for (let col = normalized.startCol; col <= normalized.endCol && emittedCells < MAX_AGENT_RENDERED_CONTEXT_CELLS; col += 1) {
      const address = formatAddress(row, col)
      const snapshot = viewportStore.peekCell(normalized.sheetName, address)
      if (!snapshot) {
        return null
      }
      rowEntries.push({
        address,
        input: snapshot.input ?? null,
        value: snapshot.value,
        formula: snapshot.formula ? `=${snapshot.formula.replace(/^=/u, '')}` : null,
        displayFormat: snapshot.format ?? null,
        styleId: snapshot.styleId ?? null,
        numberFormatId: snapshot.numberFormatId ?? null,
        style: viewportStore.getCellStyle(snapshot.styleId) ?? null,
      })
      emittedCells += 1
    }
    rows.push(rowEntries)
  }
  return {
    range: {
      sheetName: normalized.sheetName,
      startAddress: normalized.startAddress,
      endAddress: normalized.endAddress,
    },
    rowCount,
    columnCount,
    cellCount,
    truncated: emittedCells < cellCount,
    rows,
  }
}

export function useWorkerWorkbookAgentContext(input: {
  selection: WorkerRuntimeSelection
  selectionRangeRef: MutableRefObject<CellRangeRef>
  selectionSnapshotRef: MutableRefObject<GridSelectionSnapshot>
  selectionRef: MutableRefObject<WorkerRuntimeSelection>
  workerHandleRef: MutableRefObject<{ viewportStore: ProjectedViewportStore } | null>
  runtimeControllerRef: MutableRefObject<Pick<WorkerRuntimeSessionController, 'subscribeViewport'> | null>
}) {
  const { selection, selectionRangeRef, selectionSnapshotRef, selectionRef, workerHandleRef, runtimeControllerRef } = input
  const [renderedAgentContextVersion, setRenderedAgentContextVersion] = useState(0)
  const [renderedAgentContextProofVersion, setRenderedAgentContextProofVersion] = useState(0)
  const visibleViewportRef = useRef<Viewport>(selectionViewport(selection))
  const visibleViewportSubscriptionRef = useRef<{
    readonly cleanup: () => void
    readonly sheetName: string
    readonly viewport: Viewport
  } | null>(null)
  const lastRenderedAgentContextKeyRef = useRef<string | null>(null)
  const lastRenderedAgentContextProofKeyRef = useRef<string | null>(null)
  const selectedAddress = selection.address
  const selectedSheetName = selection.sheetName
  const selectedRangeStartAddress = selectionSnapshotRef.current.range.startAddress
  const selectedRangeEndAddress = selectionSnapshotRef.current.range.endAddress

  const buildCurrentAgentContext = useCallback((): WorkbookAgentUiContext => {
    const activeSelection = selectionSnapshotRef.current
    const activeViewport = visibleViewportRef.current
    const viewportRange = {
      sheetName: activeSelection.sheetName,
      startAddress: formatAddress(activeViewport.rowStart, activeViewport.colStart),
      endAddress: formatAddress(activeViewport.rowEnd, activeViewport.colEnd),
    }
    return buildWorkbookAgentContext({
      selection: activeSelection,
      viewport: activeViewport,
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: workerHandleRef.current?.viewportStore.getLastAuthoritativeRevision() ?? null,
        batchId: workerHandleRef.current?.viewportStore.getLastMetrics().batchId ?? null,
        selection: buildRenderedRangeSnapshot(workerHandleRef.current?.viewportStore, selectionRangeRef.current),
        visibleRange: buildRenderedRangeSnapshot(workerHandleRef.current?.viewportStore, viewportRange),
      },
    })
  }, [selectionRangeRef, selectionSnapshotRef, workerHandleRef])

  const rememberCurrentRenderedAgentContext = useCallback(() => {
    const context = buildCurrentAgentContext()
    lastRenderedAgentContextKeyRef.current = stringifyWorkbookAgentUiContextSemanticKey(context)
    lastRenderedAgentContextProofKeyRef.current = stringifyWorkbookAgentUiContextRenderedProofKey(context)
  }, [buildCurrentAgentContext])

  const notifyRenderedAgentContextChanged = useCallback(() => {
    const context = buildCurrentAgentContext()
    const nextKey = stringifyWorkbookAgentUiContextSemanticKey(context)
    const nextProofKey = stringifyWorkbookAgentUiContextRenderedProofKey(context)
    if (lastRenderedAgentContextKeyRef.current === nextKey) {
      if (lastRenderedAgentContextProofKeyRef.current !== nextProofKey) {
        lastRenderedAgentContextProofKeyRef.current = nextProofKey
        setRenderedAgentContextProofVersion((version) => version + 1)
      }
      return
    }
    lastRenderedAgentContextKeyRef.current = nextKey
    lastRenderedAgentContextProofKeyRef.current = nextProofKey
    setRenderedAgentContextVersion((version) => version + 1)
    setRenderedAgentContextProofVersion((version) => version + 1)
  }, [buildCurrentAgentContext])

  useLayoutEffect(() => {
    rememberCurrentRenderedAgentContext()
  }, [rememberCurrentRenderedAgentContext, selectedAddress, selectedRangeEndAddress, selectedRangeStartAddress, selectedSheetName])

  const syncVisibleViewportProjection = useCallback(
    (sheetName: string, viewport: Viewport): void => {
      visibleViewportRef.current = viewport
      const current = visibleViewportSubscriptionRef.current
      if (current && current.sheetName === sheetName && viewportContains(current.viewport, viewport)) {
        return
      }
      current?.cleanup()
      const controller = runtimeControllerRef.current
      if (!controller) {
        visibleViewportSubscriptionRef.current = null
        return
      }
      const projectionViewport = expandProjectionViewport(viewport)
      visibleViewportSubscriptionRef.current = {
        cleanup: controller.subscribeViewport(
          sheetName,
          projectionViewport,
          () => {
            notifyRenderedAgentContextChanged()
          },
          { initialPatch: 'full', notifyOnProofRevision: true },
        ),
        sheetName,
        viewport: projectionViewport,
      }
    },
    [notifyRenderedAgentContextChanged, runtimeControllerRef],
  )

  useLayoutEffect(() => {
    syncVisibleViewportProjection(selectedSheetName, selectionViewport({ address: selectedAddress, sheetName: selectedSheetName }))
  }, [selectedAddress, selectedSheetName, syncVisibleViewportProjection])

  useEffect(() => {
    syncVisibleViewportProjection(selection.sheetName, visibleViewportRef.current)
    return () => {
      visibleViewportSubscriptionRef.current?.cleanup()
      visibleViewportSubscriptionRef.current = null
    }
  }, [runtimeControllerRef, selection.sheetName, syncVisibleViewportProjection])

  const handleVisibleViewportChange = useCallback(
    (viewport: Viewport) => {
      syncVisibleViewportProjection(selectionRef.current.sheetName, viewport)
    },
    [selectionRef, syncVisibleViewportProjection],
  )

  const getAgentContext = buildCurrentAgentContext

  const resetVisibleViewportForSheet = useCallback((nextSelection: WorkerRuntimeSelection) => {
    visibleViewportRef.current = selectionViewport(nextSelection)
  }, [])

  return {
    agentContextVersion: [
      selectionSnapshotRef.current.sheetName,
      selectionSnapshotRef.current.address,
      selectionSnapshotRef.current.range.startAddress,
      selectionSnapshotRef.current.range.endAddress,
      renderedAgentContextVersion,
    ].join(':'),
    agentContextProofVersion: renderedAgentContextProofVersion,
    getAgentContext,
    handleVisibleViewportChange,
    resetVisibleViewportForSheet,
  }
}
