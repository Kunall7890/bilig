import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { GridHoverState } from './gridHover.js'
import { MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, MIN_COLUMN_WIDTH, MIN_ROW_HEIGHT } from './gridMetrics.js'
import type { HeaderSelection, VisibleRegionState } from './gridPointer.js'

export const RESIZE_HANDLE_DOUBLE_CLICK_MS = 700
export const RESIZE_PREVIEW_THROTTLE_MS = 96

interface PointerEventLike {
  clientX: number
  clientY: number
}

interface ResizePointerEventLike extends PointerEventLike {
  readonly detail: number
  preventDefault(): void
  stopPropagation(): void
}

interface PointerListenerTarget {
  addEventListener(
    type: 'pointermove' | 'pointerup' | 'pointercancel',
    listener: (event: PointerEventLike) => void,
    useCapture: boolean,
  ): void
  removeEventListener(
    type: 'pointermove' | 'pointerup' | 'pointercancel',
    listener: (event: PointerEventLike) => void,
    useCapture: boolean,
  ): void
}

type ResizeCleanup = ((event?: PointerEventLike) => void) | null

type SetGridHoverState = (updater: (current: GridHoverState) => GridHoverState) => void

function clampColumnResizeWidth(size: number): number {
  return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(size)))
}

function clampRowResizeHeight(size: number): number {
  return Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, Math.round(size)))
}

function setResizeHoverState(input: { setHoverState: SetGridHoverState; kind: 'column' | 'row'; index: number }): void {
  const { index, kind, setHoverState } = input
  const header = { kind, index }
  const next: GridHoverState = {
    cell: null,
    header,
    cursor: kind === 'column' ? 'col-resize' : 'row-resize',
  }
  setHoverState((current) => {
    if (
      current.cell === null &&
      current.header?.kind === header.kind &&
      current.header.index === header.index &&
      current.cursor === next.cursor
    ) {
      return current
    }
    return next
  })
}

export function applyWorkbookGridColumnAutofit(input: {
  columnIndex: number
  computeAutofitColumnWidth: (columnIndex: number) => number
  finishResize: () => void
  resetPointerInteraction: () => void
  setActiveResizeColumn: (columnIndex: number | null) => void
  applyAutofitWidth: (columnIndex: number, width: number) => void
}): void {
  const { applyAutofitWidth, columnIndex, computeAutofitColumnWidth, finishResize, resetPointerInteraction, setActiveResizeColumn } = input
  const autofitWidth = computeAutofitColumnWidth(columnIndex)
  finishResize()
  resetPointerInteraction()
  setActiveResizeColumn(null)
  applyAutofitWidth(columnIndex, autofitWidth)
}

export function handleWorkbookGridColumnAutofitAtPointer(input: {
  event: ResizePointerEventLike
  visibleRegion: VisibleRegionState
  pointerGeometry: GridGeometrySnapshot | null
  columnWidths: Readonly<Record<number, number>>
  defaultColumnWidth: number
  isEditingCell: boolean
  commitActiveEdit: () => void
  computeAutofitColumnWidth: (columnIndex: number) => number
  applyAutofitWidth: (columnIndex: number, width: number) => void
  finishResize: () => void
  resetPointerInteraction: () => void
  setActiveResizeColumn: (columnIndex: number | null) => void
  resolveColumnResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
    columnWidths?: Readonly<Record<number, number>>,
    defaultWidth?: number,
  ) => number | null
}): boolean {
  const {
    applyAutofitWidth,
    columnWidths,
    commitActiveEdit,
    computeAutofitColumnWidth,
    defaultColumnWidth,
    event,
    finishResize,
    isEditingCell,
    pointerGeometry,
    resetPointerInteraction,
    resolveColumnResizeTargetAtPointer,
    setActiveResizeColumn,
    visibleRegion,
  } = input
  const columnIndex = resolveColumnResizeTargetAtPointer(
    event.clientX,
    event.clientY,
    visibleRegion,
    pointerGeometry,
    columnWidths,
    defaultColumnWidth,
  )
  if (columnIndex === null) {
    return false
  }
  event.preventDefault()
  event.stopPropagation()
  if (isEditingCell) {
    commitActiveEdit()
  }
  applyWorkbookGridColumnAutofit({
    columnIndex,
    computeAutofitColumnWidth,
    finishResize,
    resetPointerInteraction,
    setActiveResizeColumn,
    applyAutofitWidth,
  })
  return true
}

export function handleWorkbookGridResizePointerDown(input: {
  event: ResizePointerEventLike
  visibleRegion: VisibleRegionState
  pointerGeometry: GridGeometrySnapshot | null
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  defaultColumnWidth: number
  defaultRowHeight: number
  isEditingCell: boolean
  commitActiveEdit: () => void
  focusGrid: () => void
  setActiveHeaderDrag: (header: HeaderSelection | null) => void
  setHoverState: SetGridHoverState
  lastResizeHandleActivationRef: { current: { columnIndex: number; at: number } | null }
  now: () => number
  computeAutofitColumnWidth: (columnIndex: number) => number
  applyAutofitWidth: (columnIndex: number, width: number) => void
  finishResize: () => void
  resetPointerInteraction: () => void
  setActiveResizeColumn: (columnIndex: number | null) => void
  beginColumnResize: (columnIndex: number, startClientX: number) => void
  beginRowResize: (rowIndex: number, startClientY: number) => void
  resolveColumnResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
    columnWidths?: Readonly<Record<number, number>>,
    defaultWidth?: number,
  ) => number | null
  resolveRowResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
    rowHeights?: Readonly<Record<number, number>>,
    defaultHeight?: number,
  ) => number | null
}): boolean {
  const {
    applyAutofitWidth,
    beginColumnResize,
    beginRowResize,
    columnWidths,
    commitActiveEdit,
    computeAutofitColumnWidth,
    defaultColumnWidth,
    defaultRowHeight,
    event,
    finishResize,
    focusGrid,
    isEditingCell,
    lastResizeHandleActivationRef,
    now,
    pointerGeometry,
    resetPointerInteraction,
    resolveColumnResizeTargetAtPointer,
    resolveRowResizeTargetAtPointer,
    rowHeights,
    setActiveHeaderDrag,
    setActiveResizeColumn,
    setHoverState,
    visibleRegion,
  } = input
  const columnResizeTarget = resolveColumnResizeTargetAtPointer(
    event.clientX,
    event.clientY,
    visibleRegion,
    pointerGeometry,
    columnWidths,
    defaultColumnWidth,
  )
  if (columnResizeTarget !== null) {
    event.preventDefault()
    event.stopPropagation()
    if (isEditingCell) {
      commitActiveEdit()
    }
    focusGrid()
    setActiveHeaderDrag(null)
    const activationTime = now()
    const lastActivation = lastResizeHandleActivationRef.current
    const isResizeDoubleClick =
      event.detail >= 2 ||
      (lastActivation !== null &&
        lastActivation.columnIndex === columnResizeTarget &&
        activationTime - lastActivation.at <= RESIZE_HANDLE_DOUBLE_CLICK_MS)
    lastResizeHandleActivationRef.current = { columnIndex: columnResizeTarget, at: activationTime }
    if (isResizeDoubleClick) {
      lastResizeHandleActivationRef.current = null
      applyWorkbookGridColumnAutofit({
        columnIndex: columnResizeTarget,
        computeAutofitColumnWidth,
        finishResize,
        resetPointerInteraction,
        setActiveResizeColumn,
        applyAutofitWidth,
      })
      return true
    }
    setResizeHoverState({
      setHoverState,
      kind: 'column',
      index: columnResizeTarget,
    })
    beginColumnResize(columnResizeTarget, event.clientX)
    return true
  }

  const rowResizeTarget = resolveRowResizeTargetAtPointer(
    event.clientX,
    event.clientY,
    visibleRegion,
    pointerGeometry,
    rowHeights,
    defaultRowHeight,
  )
  if (rowResizeTarget === null) {
    return false
  }
  event.preventDefault()
  event.stopPropagation()
  if (isEditingCell) {
    commitActiveEdit()
  }
  focusGrid()
  setActiveHeaderDrag(null)
  setResizeHoverState({
    setHoverState,
    kind: 'row',
    index: rowResizeTarget,
  })
  beginRowResize(rowResizeTarget, event.clientY)
  return true
}

function installGridResizeLifecycle(input: {
  cleanupRef: { current: ResizeCleanup }
  listenerTarget: PointerListenerTarget
  startResize: () => void
  finishResize: () => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  activate: () => void
  deactivate: () => void
  clearPreview: () => void
  preview: (event: PointerEventLike) => void
  commitOrClear: (event: PointerEventLike) => void
}): void {
  const {
    activate,
    cleanupRef,
    clearPreview,
    commitOrClear,
    deactivate,
    finishResize,
    listenerTarget,
    preview,
    refreshHoverState,
    startResize,
  } = input
  cleanupRef.current?.()
  startResize()
  activate()
  let pendingPreview: PointerEventLike | null = null
  let latestPreview: PointerEventLike | null = null
  let previewFrame: number | null = null
  let previewTimer: number | null = null

  const flushPreview = () => {
    previewFrame = null
    previewTimer = null
    const nextPreview = pendingPreview
    pendingPreview = null
    if (nextPreview) {
      preview(nextPreview)
    }
  }

  const schedulePreview = () => {
    if (previewFrame !== null || previewTimer !== null) {
      return
    }
    if (typeof window === 'undefined') {
      flushPreview()
      return
    }
    previewTimer = window.setTimeout(() => {
      previewTimer = null
      if (typeof window.requestAnimationFrame !== 'function') {
        flushPreview()
        return
      }
      previewFrame = window.requestAnimationFrame(flushPreview)
    }, RESIZE_PREVIEW_THROTTLE_MS)
  }

  const handlePointerMove = (nativeEvent: PointerEventLike) => {
    latestPreview = {
      clientX: nativeEvent.clientX,
      clientY: nativeEvent.clientY,
    }
    pendingPreview = latestPreview
    schedulePreview()
  }

  const cleanup = (nativeEvent?: PointerEventLike, options: { clearPreview?: boolean } = {}) => {
    if (previewTimer !== null && typeof window !== 'undefined') {
      window.clearTimeout(previewTimer)
    }
    if (previewFrame !== null && typeof window !== 'undefined' && typeof window.cancelAnimationFrame === 'function') {
      window.cancelAnimationFrame(previewFrame)
    }
    previewTimer = null
    previewFrame = null
    pendingPreview = null
    latestPreview = null
    listenerTarget.removeEventListener('pointermove', handlePointerMove, true)
    listenerTarget.removeEventListener('pointerup', handlePointerUp, true)
    listenerTarget.removeEventListener('pointercancel', handlePointerCancel, true)
    cleanupRef.current = null
    if (options.clearPreview ?? true) {
      clearPreview()
    }
    deactivate()
    finishResize()
    if (nativeEvent) {
      refreshHoverState(nativeEvent.clientX, nativeEvent.clientY, 0)
    }
  }

  const handlePointerUp = (nativeEvent: PointerEventLike) => {
    commitOrClear(latestPreview ?? nativeEvent)
    cleanup(nativeEvent, { clearPreview: false })
  }

  const handlePointerCancel = (nativeEvent: PointerEventLike) => {
    cleanup(nativeEvent)
  }

  cleanupRef.current = cleanup
  listenerTarget.addEventListener('pointermove', handlePointerMove, true)
  listenerTarget.addEventListener('pointerup', handlePointerUp, true)
  listenerTarget.addEventListener('pointercancel', handlePointerCancel, true)
}

export function beginWorkbookGridColumnResize(input: {
  cleanupRef: { current: ResizeCleanup }
  listenerTarget: PointerListenerTarget
  startResize: () => void
  finishResize: () => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  setActiveResizeColumn: (columnIndex: number | null) => void
  previewColumnWidth: (columnIndex: number, width: number) => void
  clearColumnResizePreview: (columnIndex: number) => void
  commitColumnWidth: (columnIndex: number, width: number) => void
  columnIndex: number
  startClientX: number
  columnWidths: Readonly<Record<number, number>>
  defaultColumnWidth: number
}): void {
  const {
    cleanupRef,
    clearColumnResizePreview,
    columnIndex,
    columnWidths,
    commitColumnWidth,
    defaultColumnWidth,
    finishResize,
    listenerTarget,
    previewColumnWidth,
    refreshHoverState,
    setActiveResizeColumn,
    startClientX,
    startResize,
  } = input
  const startWidth = columnWidths[columnIndex] ?? defaultColumnWidth

  installGridResizeLifecycle({
    cleanupRef,
    listenerTarget,
    startResize,
    finishResize,
    refreshHoverState,
    activate: () => setActiveResizeColumn(columnIndex),
    deactivate: () => setActiveResizeColumn(null),
    clearPreview: () => clearColumnResizePreview(columnIndex),
    preview: (nativeEvent) => {
      previewColumnWidth(columnIndex, startWidth + (nativeEvent.clientX - startClientX))
    },
    commitOrClear: (nativeEvent) => {
      const finalWidth = clampColumnResizeWidth(startWidth + (nativeEvent.clientX - startClientX))
      clearColumnResizePreview(columnIndex)
      if (finalWidth === startWidth) {
        return
      }
      commitColumnWidth(columnIndex, finalWidth)
    },
  })
}

export function beginWorkbookGridRowResize(input: {
  cleanupRef: { current: ResizeCleanup }
  listenerTarget: PointerListenerTarget
  startResize: () => void
  finishResize: () => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  setActiveResizeRow: (rowIndex: number | null) => void
  previewRowHeight: (rowIndex: number, height: number) => void
  clearRowResizePreview: (rowIndex: number) => void
  commitRowHeight: (rowIndex: number, height: number) => void
  rowIndex: number
  startClientY: number
  rowHeights: Readonly<Record<number, number>>
  defaultRowHeight: number
}): void {
  const {
    cleanupRef,
    clearRowResizePreview,
    commitRowHeight,
    defaultRowHeight,
    finishResize,
    listenerTarget,
    previewRowHeight,
    refreshHoverState,
    rowHeights,
    rowIndex,
    setActiveResizeRow,
    startClientY,
    startResize,
  } = input
  const startHeight = rowHeights[rowIndex] ?? defaultRowHeight

  installGridResizeLifecycle({
    cleanupRef,
    listenerTarget,
    startResize,
    finishResize,
    refreshHoverState,
    activate: () => setActiveResizeRow(rowIndex),
    deactivate: () => setActiveResizeRow(null),
    clearPreview: () => clearRowResizePreview(rowIndex),
    preview: (nativeEvent) => {
      previewRowHeight(rowIndex, startHeight + (nativeEvent.clientY - startClientY))
    },
    commitOrClear: (nativeEvent) => {
      const finalHeight = clampRowResizeHeight(startHeight + (nativeEvent.clientY - startClientY))
      clearRowResizePreview(rowIndex)
      if (finalHeight === startHeight) {
        return
      }
      commitRowHeight(rowIndex, finalHeight)
    },
  })
}
