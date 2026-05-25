import { useCallback, useEffect, useLayoutEffect, useRef, useState, type MutableRefObject } from 'react'
import type { GridSelectionSnapshot } from '@bilig/grid'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  stringifyWorkbookAgentUiContextRenderedProofKey,
  stringifyWorkbookAgentUiContextSemanticKey,
  type WorkbookAgentRenderedRange,
  type WorkbookAgentRenderedVisibleSceneProof,
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
    ...(typeof range.sheetId === 'number' && Number.isSafeInteger(range.sheetId) ? { sheetId: range.sheetId } : {}),
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
  const sheetIdentity = viewportStore.getSheetIdentity(normalized.sheetName)
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
      ...(sheetIdentity === null ? {} : { sheetId: sheetIdentity.sheetId }),
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

function hasText(value: string | null | undefined): value is string {
  return typeof value === 'string' && value.length > 0
}

function readVisibleSceneProofAttribute(element: Element, name: string): string | null {
  const value = element.getAttribute(name)
  return hasText(value) ? value : null
}

function readWorkbookAgentVisibleSceneProof(): WorkbookAgentRenderedVisibleSceneProof | null {
  const grid = document.querySelector('[data-testid="sheet-grid"]')
  const typeGpu = document.querySelector('[data-testid="grid-pane-renderer"]')
  if (!(grid instanceof HTMLElement) || !(typeGpu instanceof HTMLCanvasElement)) {
    return null
  }
  const rendererMode = readVisibleSceneProofAttribute(typeGpu, 'data-renderer-mode')
  const frameProofStatus = readVisibleSceneProofAttribute(typeGpu, 'data-v3-frame-proof-status')
  const frameProofSignature = readVisibleSceneProofAttribute(typeGpu, 'data-v3-frame-proof-signature')
  const presentedFrameProofSignature = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-frame-proof-signature')
  const currentSceneEpochSignature = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-scene-epoch-signature')
  const currentSceneOwnershipSignature = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-scene-ownership-signature')
  const presentedSceneEpochSignature = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-scene-epoch-signature')
  const presentedSceneOwnershipSignature = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-scene-ownership-signature')
  const currentSceneEpoch = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-scene-epoch')
  const presentedSceneEpoch = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-scene-epoch')
  const currentFillHandleRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-fill-handle-revision')
  const presentedFillHandleRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-fill-handle-revision')
  const currentSelectionRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-selection-revision')
  const presentedSelectionRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-selection-revision')
  const currentViewportRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-viewport-revision')
  const presentedViewportRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-viewport-revision')
  const currentSemanticMutationRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-semantic-mutation-revision')
  const presentedSemanticMutationRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-semantic-mutation-revision')
  const currentWorkbookRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-current-workbook-revision')
  const presentedWorkbookRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-presented-workbook-revision')
  const gridAuthoritativeRevision = readVisibleSceneProofAttribute(grid, 'data-render-authoritative-revision')
  const typeGpuAuthoritativeRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-authoritative-render-revision')
  const visibleAuthoritativeRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-visible-authoritative-render-revision')
  const tileSceneRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-tile-scene-revision')
  const visibleRenderRevision = readVisibleSceneProofAttribute(typeGpu, 'data-v3-visible-render-revision')
  return {
    rendererMode,
    frameProofStatus,
    frameProofSignature,
    presentedFrameProofSignature,
    currentSceneEpochSignature,
    currentSceneOwnershipSignature,
    presentedSceneEpochSignature,
    presentedSceneOwnershipSignature,
    currentSceneEpoch,
    presentedSceneEpoch,
    currentFillHandleRevision,
    presentedFillHandleRevision,
    currentSelectionRevision,
    presentedSelectionRevision,
    currentViewportRevision,
    presentedViewportRevision,
    currentSemanticMutationRevision,
    presentedSemanticMutationRevision,
    currentWorkbookRevision,
    presentedWorkbookRevision,
    gridAuthoritativeRevision,
    typeGpuAuthoritativeRevision,
    visibleAuthoritativeRevision,
    tileSceneRevision,
    visibleRenderRevision,
    hasPresentedFrame: typeGpu.getAttribute('data-v3-has-presented-frame') === 'true',
    hasPresentedVisibleFrame: typeGpu.getAttribute('data-v3-has-presented-visible-frame') === 'true',
    frameProofMatchesPresentedFrame:
      hasText(frameProofSignature) && hasText(presentedFrameProofSignature) && frameProofSignature === presentedFrameProofSignature,
    visibleSceneEpochMatchesPresentedFrame:
      hasText(currentSceneEpochSignature) &&
      hasText(presentedSceneEpochSignature) &&
      currentSceneEpochSignature === presentedSceneEpochSignature,
    visibleSceneOwnershipMatchesPresentedFrame:
      hasText(currentSceneOwnershipSignature) &&
      hasText(presentedSceneOwnershipSignature) &&
      currentSceneOwnershipSignature === presentedSceneOwnershipSignature,
    visibleAuthoritativeRevisionMatchesGrid:
      hasText(gridAuthoritativeRevision) &&
      hasText(typeGpuAuthoritativeRevision) &&
      hasText(visibleAuthoritativeRevision) &&
      typeGpuAuthoritativeRevision === gridAuthoritativeRevision &&
      visibleAuthoritativeRevision === gridAuthoritativeRevision,
    visibleRenderRevisionMatchesTileScene:
      hasText(tileSceneRevision) && hasText(visibleRenderRevision) && visibleRenderRevision === tileSceneRevision,
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
    const activeSheetIdentity = workerHandleRef.current?.viewportStore.getSheetIdentity(activeSelection.sheetName) ?? null
    const viewportRange = {
      sheetName: activeSelection.sheetName,
      startAddress: formatAddress(activeViewport.rowStart, activeViewport.colStart),
      endAddress: formatAddress(activeViewport.rowEnd, activeViewport.colEnd),
    }
    return buildWorkbookAgentContext({
      selection: activeSelection,
      ...(activeSheetIdentity === null ? {} : { selectionSheetId: activeSheetIdentity.sheetId }),
      viewport: activeViewport,
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: workerHandleRef.current?.viewportStore.getLastAuthoritativeRevision() ?? null,
        batchId: workerHandleRef.current?.viewportStore.getLastMetrics().batchId ?? null,
        visibleSceneProof: readWorkbookAgentVisibleSceneProof(),
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
