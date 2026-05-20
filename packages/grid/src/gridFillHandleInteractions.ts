import { createRectangleSelectionFromRange, rectangleToAddresses } from './gridSelection.js'
import {
  applyGridDragAutoScroll,
  resolveDefaultCancelAnimationFrame,
  resolveDefaultRequestAnimationFrame,
  type CancelGridDragFrame,
  type GridDragScrollViewport,
  type RequestGridDragFrame,
} from './gridDragAutoScroll.js'
import { resolveFillHandlePreviewRange, resolveFillHandleSelectionRange } from './gridFillHandle.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'

interface FillHandlePointerEventLike {
  readonly clientX: number
  readonly clientY: number
  readonly pointerId: number
}

interface FillHandleListenerTarget {
  addEventListener(
    type: 'pointermove' | 'pointerup' | 'pointercancel',
    listener: (event: FillHandlePointerEventLike) => void,
    useCapture: boolean,
  ): void
  removeEventListener(
    type: 'pointermove' | 'pointerup' | 'pointercancel',
    listener: (event: FillHandlePointerEventLike) => void,
    useCapture: boolean,
  ): void
}

export function beginWorkbookGridFillHandleDrag(input: {
  cleanupRef: { current: (() => void) | null }
  listenerTarget: FillHandleListenerTarget
  pointerId: number
  sourceRange: Rectangle
  gridSelection: GridSelection
  resolvePointerCell: (clientX: number, clientY: number) => Item | null
  setGridSelection: (selection: GridSelection) => void
  onSelectionChange: (selection: GridSelection) => void
  onFillRange: (sourceStartAddress: string, sourceEndAddress: string, targetStartAddress: string, targetEndAddress: string) => void
  setFillPreviewRange: (range: Rectangle | null) => void
  setFillPreviewRangeRef: (range: Rectangle | null) => void
  setIsFillHandleDragging: (isDragging: boolean) => void
  resetHoverState: () => void
  scrollViewport?: GridDragScrollViewport | null | undefined
  requestAnimationFrame?: RequestGridDragFrame | undefined
  cancelAnimationFrame?: CancelGridDragFrame | undefined
}): void {
  const {
    cancelAnimationFrame = resolveDefaultCancelAnimationFrame(),
    cleanupRef,
    gridSelection,
    listenerTarget,
    onFillRange,
    onSelectionChange,
    pointerId,
    requestAnimationFrame = resolveDefaultRequestAnimationFrame(),
    resolvePointerCell,
    resetHoverState,
    scrollViewport = null,
    setFillPreviewRange,
    setFillPreviewRangeRef,
    setGridSelection,
    setIsFillHandleDragging,
    sourceRange,
  } = input
  let activePreviewRange: Rectangle | null = null
  let lastPointerEvent: FillHandlePointerEventLike | null = null
  let autoScrollFrame: number | null = null

  cleanupRef.current?.()
  setFillPreviewRangeRef(null)
  setFillPreviewRange(null)
  setIsFillHandleDragging(true)
  resetHoverState()

  const updatePreview = (nativeEvent: FillHandlePointerEventLike) => {
    if (nativeEvent.pointerId !== pointerId) {
      return
    }
    const pointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY)
    activePreviewRange = pointerCell ? resolveFillHandlePreviewRange(sourceRange, pointerCell) : null
    setFillPreviewRangeRef(activePreviewRange)
    setFillPreviewRange(activePreviewRange)
  }

  const cancelAutoScroll = () => {
    if (autoScrollFrame === null) {
      return
    }
    cancelAnimationFrame?.(autoScrollFrame)
    autoScrollFrame = null
  }

  const scheduleAutoScroll = () => {
    if (!scrollViewport || !requestAnimationFrame || autoScrollFrame !== null) {
      return
    }
    autoScrollFrame = requestAnimationFrame(() => {
      autoScrollFrame = null
      if (!lastPointerEvent || !applyGridDragAutoScroll(scrollViewport, lastPointerEvent)) {
        return
      }
      updatePreview(lastPointerEvent)
      scheduleAutoScroll()
    })
  }

  const move = (nativeEvent: FillHandlePointerEventLike) => {
    if (nativeEvent.pointerId !== pointerId) {
      return
    }
    lastPointerEvent = nativeEvent
    updatePreview(nativeEvent)
    scheduleAutoScroll()
  }

  const cleanup = () => {
    if (cleanupRef.current !== cleanup) {
      return
    }
    cancelAutoScroll()
    cleanupRef.current = null
    listenerTarget.removeEventListener('pointermove', move, true)
    listenerTarget.removeEventListener('pointerup', handlePointerUp, true)
    listenerTarget.removeEventListener('pointercancel', handlePointerCancel, true)
    setIsFillHandleDragging(false)
    resetHoverState()
  }

  const finish = () => {
    const previewRange = activePreviewRange
    if (previewRange) {
      const source = rectangleToAddresses(sourceRange)
      const target = rectangleToAddresses(previewRange)
      const nextSelectionRange = resolveFillHandleSelectionRange(sourceRange, previewRange)
      const nextSelection = createRectangleSelectionFromRange(nextSelectionRange)
      if (gridSelection.current?.cell && nextSelection.current) {
        nextSelection.current = {
          ...nextSelection.current,
          cell: gridSelection.current.cell,
        }
      }
      setGridSelection(nextSelection)
      onSelectionChange(nextSelection)
      if (source.startAddress !== target.startAddress || source.endAddress !== target.endAddress) {
        onFillRange(source.startAddress, source.endAddress, target.startAddress, target.endAddress)
      }
    }
    activePreviewRange = null
    setFillPreviewRangeRef(null)
    setFillPreviewRange(null)
    cleanup()
  }

  const handlePointerUp = (nativeEvent: FillHandlePointerEventLike) => {
    if (nativeEvent.pointerId !== pointerId) {
      return
    }
    lastPointerEvent = nativeEvent
    updatePreview(nativeEvent)
    finish()
  }

  const handlePointerCancel = (nativeEvent: FillHandlePointerEventLike) => {
    if (nativeEvent.pointerId !== pointerId) {
      return
    }
    activePreviewRange = null
    setFillPreviewRangeRef(null)
    setFillPreviewRange(null)
    cleanup()
  }

  cleanupRef.current = cleanup
  listenerTarget.addEventListener('pointermove', move, true)
  listenerTarget.addEventListener('pointerup', handlePointerUp, true)
  listenerTarget.addEventListener('pointercancel', handlePointerCancel, true)
}
