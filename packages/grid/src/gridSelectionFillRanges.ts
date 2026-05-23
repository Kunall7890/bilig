import type { Item, Rectangle } from './gridTypes.js'

export type GridSelectionFillRange = Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>

export function selectionFillRangesForRange(range: GridSelectionFillRange): readonly GridSelectionFillRange[] {
  if (range.width <= 0 || range.height <= 0) {
    return []
  }
  return [range]
}

export function selectionFillRangesForRangeExcludingCell(
  range: GridSelectionFillRange,
  excludedCell: Item | null | undefined,
): readonly GridSelectionFillRange[] {
  if (range.width <= 0 || range.height <= 0) {
    return []
  }
  if (
    !excludedCell ||
    excludedCell[0] < range.x ||
    excludedCell[0] >= range.x + range.width ||
    excludedCell[1] < range.y ||
    excludedCell[1] >= range.y + range.height
  ) {
    return [range]
  }

  const [excludedCol, excludedRow] = excludedCell
  const ranges: GridSelectionFillRange[] = []
  appendPositiveRange(ranges, {
    x: range.x,
    y: range.y,
    width: range.width,
    height: excludedRow - range.y,
  })
  appendPositiveRange(ranges, {
    x: range.x,
    y: excludedRow + 1,
    width: range.width,
    height: range.y + range.height - excludedRow - 1,
  })
  appendPositiveRange(ranges, {
    x: range.x,
    y: excludedRow,
    width: excludedCol - range.x,
    height: 1,
  })
  appendPositiveRange(ranges, {
    x: excludedCol + 1,
    y: excludedRow,
    width: range.x + range.width - excludedCol - 1,
    height: 1,
  })
  return ranges
}

function appendPositiveRange(ranges: GridSelectionFillRange[], range: GridSelectionFillRange): void {
  if (range.width > 0 && range.height > 0) {
    ranges.push(range)
  }
}
