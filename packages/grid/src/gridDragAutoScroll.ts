const GRID_DRAG_AUTOSCROLL_EDGE_PX = 36
const GRID_DRAG_AUTOSCROLL_MAX_STEP_PX = 32

export interface GridDragPointerEventLike {
  readonly clientX: number
  readonly clientY: number
}

export interface GridDragScrollViewport {
  readonly clientHeight: number
  readonly clientWidth: number
  readonly scrollHeight: number
  readonly scrollWidth: number
  scrollLeft: number
  scrollTop: number
  dispatchEvent(event: Event): boolean
  getBoundingClientRect(): Pick<DOMRect, 'bottom' | 'left' | 'right' | 'top'>
}

export type RequestGridDragFrame = (callback: FrameRequestCallback) => number
export type CancelGridDragFrame = (handle: number) => void

export function resolveDefaultRequestAnimationFrame(): RequestGridDragFrame | undefined {
  return typeof window === 'undefined' ? undefined : window.requestAnimationFrame.bind(window)
}

export function resolveDefaultCancelAnimationFrame(): CancelGridDragFrame | undefined {
  return typeof window === 'undefined' ? undefined : window.cancelAnimationFrame.bind(window)
}

export function applyGridDragAutoScroll(scrollViewport: GridDragScrollViewport, pointer: GridDragPointerEventLike): boolean {
  const rect = scrollViewport.getBoundingClientRect()
  const deltaX = resolveGridDragAutoScrollStep(pointer.clientX, rect.left, rect.right)
  const deltaY = resolveGridDragAutoScrollStep(pointer.clientY, rect.top, rect.bottom)
  if (deltaX === 0 && deltaY === 0) {
    return false
  }

  const nextScrollLeft = clamp(scrollViewport.scrollLeft + deltaX, 0, Math.max(0, scrollViewport.scrollWidth - scrollViewport.clientWidth))
  const nextScrollTop = clamp(scrollViewport.scrollTop + deltaY, 0, Math.max(0, scrollViewport.scrollHeight - scrollViewport.clientHeight))
  if (nextScrollLeft === scrollViewport.scrollLeft && nextScrollTop === scrollViewport.scrollTop) {
    return false
  }

  scrollViewport.scrollLeft = nextScrollLeft
  scrollViewport.scrollTop = nextScrollTop
  scrollViewport.dispatchEvent(new Event('scroll'))
  return true
}

function resolveGridDragAutoScrollStep(position: number, start: number, end: number): number {
  if (position < start + GRID_DRAG_AUTOSCROLL_EDGE_PX) {
    return -resolveGridDragAutoScrollMagnitude(start + GRID_DRAG_AUTOSCROLL_EDGE_PX - position)
  }
  if (position > end - GRID_DRAG_AUTOSCROLL_EDGE_PX) {
    return resolveGridDragAutoScrollMagnitude(position - (end - GRID_DRAG_AUTOSCROLL_EDGE_PX))
  }
  return 0
}

function resolveGridDragAutoScrollMagnitude(distanceInsideEdge: number): number {
  const intensity = clamp(distanceInsideEdge / GRID_DRAG_AUTOSCROLL_EDGE_PX, 0, 1)
  return Math.max(1, Math.round(intensity * GRID_DRAG_AUTOSCROLL_MAX_STEP_PX))
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}
