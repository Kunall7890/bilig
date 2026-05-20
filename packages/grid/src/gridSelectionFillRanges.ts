import type { Item, Rectangle } from './gridTypes.js'

export type GridSelectionFillRange = Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>

export function splitSelectionFillRangeAroundActiveCell(
  range: GridSelectionFillRange,
  activeCell: Item | null | undefined,
): readonly GridSelectionFillRange[] {
  if (!activeCell || !cellInRange(activeCell, range)) {
    return [range]
  }

  const [activeCol, activeRow] = activeCell
  const rangeBottomExclusive = range.y + range.height
  const rangeRightExclusive = range.x + range.width
  const ranges: GridSelectionFillRange[] = []

  pushFillRange(ranges, {
    x: range.x,
    y: range.y,
    width: range.width,
    height: activeRow - range.y,
  })
  pushFillRange(ranges, {
    x: range.x,
    y: activeRow + 1,
    width: range.width,
    height: rangeBottomExclusive - activeRow - 1,
  })
  pushFillRange(ranges, {
    x: range.x,
    y: activeRow,
    width: activeCol - range.x,
    height: 1,
  })
  pushFillRange(ranges, {
    x: activeCol + 1,
    y: activeRow,
    width: rangeRightExclusive - activeCol - 1,
    height: 1,
  })

  return ranges
}

function pushFillRange(ranges: GridSelectionFillRange[], range: GridSelectionFillRange): void {
  if (range.width > 0 && range.height > 0) {
    ranges.push(range)
  }
}

function cellInRange(cell: Item, range: GridSelectionFillRange): boolean {
  return cell[0] >= range.x && cell[0] < range.x + range.width && cell[1] >= range.y && cell[1] < range.y + range.height
}
