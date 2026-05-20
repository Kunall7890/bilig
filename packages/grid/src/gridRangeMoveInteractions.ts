import { createRectangleSelectionFromRange, rectangleToAddresses } from './gridSelection.js'
import {
  applyGridDragAutoScroll,
  resolveDefaultCancelAnimationFrame,
  resolveDefaultRequestAnimationFrame,
  type CancelGridDragFrame,
  type GridDragPointerEventLike,
  type GridDragScrollViewport,
  type RequestGridDragFrame,
} from './gridDragAutoScroll.js'
import type { GridHoverState } from './gridHover.js'
import { resolveMovedRange, sameRectangle } from './gridRangeMove.js'
import type { Item, Rectangle } from './gridTypes.js'

type PointerEventLike = GridDragPointerEventLike

interface PointerListenerTarget {
  addEventListener(type: 'pointermove' | 'pointerup', listener: (event: PointerEventLike) => void, useCapture: boolean): void
  removeEventListener(type: 'pointermove' | 'pointerup', listener: (event: PointerEventLike) => void, useCapture: boolean): void
}

export function beginWorkbookGridRangeMove(input: {
  cleanupRef: { current: (() => void) | null }
  listenerTarget: PointerListenerTarget
  sourceRange: Rectangle
  pointerCell: Item
  activeCell?: Item | null | undefined
  resolvePointerCell: (clientX: number, clientY: number) => Item | null
  setGridSelection: (selection: ReturnType<typeof createRectangleSelectionFromRange>) => void
  onSelectionChange: (selection: ReturnType<typeof createRectangleSelectionFromRange>) => void
  onMoveRange: (sourceStartAddress: string, sourceEndAddress: string, targetStartAddress: string, targetEndAddress: string) => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  setIsRangeMoveDragging: (isDragging: boolean) => void
  setHoverState: (state: GridHoverState) => void
  scrollViewport?: GridDragScrollViewport | null | undefined
  requestAnimationFrame?: RequestGridDragFrame | undefined
  cancelAnimationFrame?: CancelGridDragFrame | undefined
}): void {
  const {
    activeCell = null,
    cancelAnimationFrame = resolveDefaultCancelAnimationFrame(),
    cleanupRef,
    listenerTarget,
    onMoveRange,
    pointerCell,
    requestAnimationFrame = resolveDefaultRequestAnimationFrame(),
    refreshHoverState,
    resolvePointerCell,
    scrollViewport = null,
    setGridSelection,
    onSelectionChange,
    setHoverState,
    setIsRangeMoveDragging,
    sourceRange,
  } = input
  const anchorOffset: Item = [pointerCell[0] - sourceRange.x, pointerCell[1] - sourceRange.y]
  const activeCellOffset = resolveActiveCellOffset(sourceRange, activeCell)
  let previewRange = sourceRange
  let lastPointerEvent: PointerEventLike | null = null
  let autoScrollFrame: number | null = null

  cleanupRef.current?.()
  setIsRangeMoveDragging(true)
  setHoverState({ cell: null, header: null, cursor: 'grabbing' })

  const updatePreview = (nativeEvent: PointerEventLike) => {
    const nextPointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY)
    if (!nextPointerCell) {
      return
    }
    const nextRange = resolveMovedRange(sourceRange, nextPointerCell, anchorOffset)
    if (sameRectangle(previewRange, nextRange)) {
      return
    }
    previewRange = nextRange
    const nextSelection = createRangeMoveSelection(nextRange, activeCellOffset)
    setGridSelection(nextSelection)
    onSelectionChange(nextSelection)
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

  const move = (nativeEvent: PointerEventLike) => {
    lastPointerEvent = nativeEvent
    updatePreview(nativeEvent)
    scheduleAutoScroll()
  }

  const cleanup = (nativeEvent?: PointerEventLike) => {
    cancelAutoScroll()
    listenerTarget.removeEventListener('pointermove', move, true)
    listenerTarget.removeEventListener('pointerup', up, true)
    cleanupRef.current = null
    setIsRangeMoveDragging(false)
    if (nativeEvent) {
      refreshHoverState(nativeEvent.clientX, nativeEvent.clientY, 0)
    }
  }

  const up = (nativeEvent: PointerEventLike) => {
    lastPointerEvent = nativeEvent
    updatePreview(nativeEvent)
    cleanup(nativeEvent)
    const nextSelection = createRangeMoveSelection(previewRange, activeCellOffset)
    setGridSelection(nextSelection)
    onSelectionChange(nextSelection)
    if (sameRectangle(sourceRange, previewRange)) {
      return
    }
    const sourceAddresses = rectangleToAddresses(sourceRange)
    const targetAddresses = rectangleToAddresses(previewRange)
    onMoveRange(sourceAddresses.startAddress, sourceAddresses.endAddress, targetAddresses.startAddress, targetAddresses.endAddress)
  }

  cleanupRef.current = () => {
    cleanup()
  }
  listenerTarget.addEventListener('pointermove', move, true)
  listenerTarget.addEventListener('pointerup', up, true)
}

function resolveActiveCellOffset(sourceRange: Rectangle, activeCell: Item | null): Item | null {
  if (
    !activeCell ||
    activeCell[0] < sourceRange.x ||
    activeCell[0] >= sourceRange.x + sourceRange.width ||
    activeCell[1] < sourceRange.y ||
    activeCell[1] >= sourceRange.y + sourceRange.height
  ) {
    return null
  }
  return [activeCell[0] - sourceRange.x, activeCell[1] - sourceRange.y]
}

function createRangeMoveSelection(range: Rectangle, activeCellOffset: Item | null): ReturnType<typeof createRectangleSelectionFromRange> {
  const selection = createRectangleSelectionFromRange(range)
  if (activeCellOffset && selection.current) {
    selection.current = {
      ...selection.current,
      cell: [range.x + activeCellOffset[0], range.y + activeCellOffset[1]],
    }
  }
  return selection
}
