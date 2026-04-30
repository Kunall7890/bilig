import type { HeaderSelection } from './gridPointer.js'
import type { Item } from './gridTypes.js'

export type GridHoverCursor = 'default' | 'cell' | 'pointer' | 'col-resize' | 'row-resize' | 'grab' | 'grabbing'

export interface GridHoverState {
  readonly cell: Item | null
  readonly header: HeaderSelection | null
  readonly cursor: GridHoverCursor
}

export function sameGridHoverState(left: GridHoverState, right: GridHoverState): boolean {
  return sameItem(left.cell, right.cell) && sameHeaderSelection(left.header, right.header) && left.cursor === right.cursor
}

function sameItem(left: Item | null, right: Item | null): boolean {
  return left === right || (left !== null && right !== null && left[0] === right[0] && left[1] === right[1])
}

function sameHeaderSelection(left: HeaderSelection | null, right: HeaderSelection | null): boolean {
  return left === right || (left !== null && right !== null && left.kind === right.kind && left.index === right.index)
}
