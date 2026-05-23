import type { Rectangle } from './gridTypes.js'

export type GridSelectionFillRange = Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>

export function selectionFillRangesForRange(range: GridSelectionFillRange): readonly GridSelectionFillRange[] {
  if (range.width <= 0 || range.height <= 0) {
    return []
  }
  return [range]
}
