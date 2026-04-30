import type { MutableRefObject } from 'react'
import type { Item } from './gridTypes.js'
import type { HeaderSelection } from './gridPointer.js'

export interface GridInteractionStateRefs {
  ignoreNextPointerSelectionRef: MutableRefObject<boolean>
  pendingPointerCellRef: MutableRefObject<Item | null>
  dragAnchorCellRef: MutableRefObject<Item | null>
  dragPointerCellRef: MutableRefObject<Item | null>
  dragHeaderSelectionRef: MutableRefObject<HeaderSelection | null>
  dragDidMoveRef: MutableRefObject<boolean>
  postDragSelectionExpiryRef: MutableRefObject<number>
  columnResizeActiveRef: MutableRefObject<boolean>
}

interface ResetGridPointerInteractionOptions {
  clearIgnoreNextPointerSelection?: boolean
  clearPostDragSelectionExpiry?: boolean
}

export function resetGridPointerInteraction(refs: GridInteractionStateRefs, options: ResetGridPointerInteractionOptions = {}): void {
  if (options.clearIgnoreNextPointerSelection) {
    refs.ignoreNextPointerSelectionRef.current = false
  }
  refs.pendingPointerCellRef.current = null
  refs.dragAnchorCellRef.current = null
  refs.dragPointerCellRef.current = null
  refs.dragHeaderSelectionRef.current = null
  refs.dragDidMoveRef.current = false
  if (options.clearPostDragSelectionExpiry ?? true) {
    refs.postDragSelectionExpiryRef.current = 0
  }
}

export function scheduleGridPointerInteractionReset(
  refs: GridInteractionStateRefs,
  options: ResetGridPointerInteractionOptions = {},
): void {
  window.requestAnimationFrame(() => {
    resetGridPointerInteraction(refs, options)
  })
}

export function startGridColumnResize(refs: GridInteractionStateRefs): void {
  refs.columnResizeActiveRef.current = true
  resetGridPointerInteraction(refs)
}

export function finishGridColumnResize(refs: GridInteractionStateRefs): void {
  window.requestAnimationFrame(() => {
    refs.columnResizeActiveRef.current = false
  })
}

export function beginGridHeaderDrag(refs: GridInteractionStateRefs, headerSelection: HeaderSelection): void {
  resetGridPointerInteraction(refs)
  refs.dragHeaderSelectionRef.current = headerSelection
}

export function beginGridBodyPointerInteraction(refs: GridInteractionStateRefs, pointerCell: Item | null): void {
  refs.ignoreNextPointerSelectionRef.current = pointerCell === null
  refs.pendingPointerCellRef.current = pointerCell
  refs.dragDidMoveRef.current = false
  refs.postDragSelectionExpiryRef.current = 0
}

export function clearGridPendingPointerActivation(refs: GridInteractionStateRefs): void {
  refs.pendingPointerCellRef.current = null
  refs.dragAnchorCellRef.current = null
  refs.dragPointerCellRef.current = null
}
