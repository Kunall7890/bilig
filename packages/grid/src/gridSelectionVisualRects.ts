import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'

const DIRECT_SELECTION_GRIDLINE_AXIS_LIMIT = 20_000

export type GridSelectionVisualRectRole =
  | 'selection-fill'
  | 'selection-gridline'
  | 'selection-border'
  | 'active-border'
  | 'fill-handle'
  | 'header-fill'
  | 'hover-fill'

export interface GridSelectionVisualRect {
  readonly role: GridSelectionVisualRectRole
  readonly key: string
  readonly bounds: Rectangle
  readonly strokeWidth?: number | undefined
}

export function buildGridSelectionVisualRects(input: {
  readonly geometry: GridGeometrySnapshot
  readonly gridSelection: GridSelection
  readonly hoverCell?: Item | null | undefined
  readonly selectedCell: Item
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
}): readonly GridSelectionVisualRect[] {
  const rects: GridSelectionVisualRect[] = []
  appendHoverVisualRects(rects, input)
  appendAxisSelectionVisualRects(rects, input)
  appendBodySelectionVisualRects(rects, input)
  return rects
}

function appendHoverVisualRects(
  rects: GridSelectionVisualRect[],
  input: {
    readonly geometry: GridGeometrySnapshot
    readonly hoverCell?: Item | null | undefined
    readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  },
): void {
  const hoverCell = input.hoverCell ?? null
  if (!hoverCell) {
    return
  }
  if (input.selectionRange && cellInRange(hoverCell, input.selectionRange)) {
    return
  }
  let segmentIndex = 0
  for (const bounds of input.geometry.rangeScreenRects({ x: hoverCell[0], y: hoverCell[1], width: 1, height: 1 })) {
    appendInsetRect(rects, 'hover-fill', `hover-fill:cell:${hoverCell[0]}:${hoverCell[1]}:${segmentIndex}`, bounds, 1, 1)
    segmentIndex += 1
  }
}

function appendBodySelectionVisualRects(
  rects: GridSelectionVisualRect[],
  input: {
    readonly geometry: GridGeometrySnapshot
    readonly gridSelection: GridSelection
    readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
    readonly showFillHandle: boolean
  },
): void {
  if (!input.selectionRange) {
    return
  }

  const hasAxisSelection = input.gridSelection.columns.length > 0 || input.gridSelection.rows.length > 0
  if (hasAxisSelection) {
    return
  }

  const isMultiCellSelection = input.selectionRange.width > 1 || input.selectionRange.height > 1

  if (isMultiCellSelection) {
    appendRangeSelectionFillAndGridlineRects(rects, input.geometry, input.selectionRange, 'selection-fill:range')

    let segmentIndex = 0
    for (const bounds of input.geometry.rangeScreenRects(input.selectionRange)) {
      rects.push({
        role: 'selection-border',
        key: stableRangeKey('selection-border:range', input.selectionRange, 0, segmentIndex),
        bounds,
        strokeWidth: 2,
      })
      segmentIndex += 1
    }
  } else {
    appendCellBorderRects(
      rects,
      input.geometry,
      [input.selectionRange.x, input.selectionRange.y],
      'active-border',
      `active-border:cell:${input.selectionRange.x}:${input.selectionRange.y}`,
    )
  }

  if (input.showFillHandle) {
    const handle = input.geometry.fillHandleScreenRect(input.selectionRange)
    if (handle) {
      rects.push({ role: 'fill-handle', key: stableRangeKey('fill-handle:range', input.selectionRange, 0, 0), bounds: handle })
    }
  }
}

function appendAxisSelectionVisualRects(
  rects: GridSelectionVisualRect[],
  input: {
    readonly geometry: GridGeometrySnapshot
    readonly gridSelection: GridSelection
    readonly selectedCell: Item
    readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  },
): void {
  const columnRanges = resolveSelectedAxisRanges({
    axisRanges: input.gridSelection.columns.ranges,
    fallbackIndex: input.selectedCell[0],
    fallbackRange: input.selectionRange
      ? { start: input.selectionRange.x, endExclusive: input.selectionRange.x + input.selectionRange.width }
      : null,
  })
  const rowRanges = resolveSelectedAxisRanges({
    axisRanges: input.gridSelection.rows.ranges,
    fallbackIndex: input.selectedCell[1],
    fallbackRange: input.selectionRange
      ? { start: input.selectionRange.y, endExclusive: input.selectionRange.y + input.selectionRange.height }
      : null,
  })

  for (const index of visibleColumnIndexes(input.geometry)) {
    if (!isIndexSelected(index, columnRanges)) {
      continue
    }
    const bounds = input.geometry.columnHeaderScreenRect(index)
    const clip =
      index < input.geometry.camera.frozenColumnCount
        ? paneFrame(input.geometry, 'column-header-frozen')
        : paneFrame(input.geometry, 'column-header-body')
    const clipped = bounds && clip ? clipRect(bounds, clip) : null
    if (clipped) {
      appendColumnHeaderFillRect(rects, `header-fill:column:${index}`, clipped)
    }
  }
  for (const index of visibleRowIndexes(input.geometry)) {
    if (!isIndexSelected(index, rowRanges)) {
      continue
    }
    const bounds = input.geometry.rowHeaderScreenRect(index)
    const clip =
      index < input.geometry.camera.frozenRowCount
        ? paneFrame(input.geometry, 'row-header-frozen')
        : paneFrame(input.geometry, 'row-header-body')
    const clipped = bounds && clip ? clipRect(bounds, clip) : null
    if (clipped) {
      appendRowHeaderFillRect(rects, `header-fill:row:${index}`, clipped)
    }
  }

  if (input.gridSelection.columns.length > 0) {
    for (const range of columnRanges) {
      const start = Math.max(0, range.start)
      const endExclusive = Math.max(start + 1, Math.min(MAX_COLS, range.endExclusive))
      appendVisibleCellSelectionFillRects(
        rects,
        input.geometry,
        { x: start, y: 0, width: endExclusive - start, height: MAX_ROWS },
        `selection-fill:columns:${start}:${endExclusive}`,
      )
    }
  }
  if (input.gridSelection.rows.length > 0) {
    for (const range of rowRanges) {
      const start = Math.max(0, range.start)
      const endExclusive = Math.max(start + 1, Math.min(MAX_ROWS, range.endExclusive))
      appendVisibleCellSelectionFillRects(
        rects,
        input.geometry,
        { x: 0, y: start, width: MAX_COLS, height: endExclusive - start },
        `selection-fill:rows:${start}:${endExclusive}`,
      )
    }
  }
}

function appendVisibleCellSelectionFillRects(
  rects: GridSelectionVisualRect[],
  geometry: GridGeometrySnapshot,
  range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  keyPrefix: string,
): void {
  const startCol = Math.max(0, range.x)
  const startRow = Math.max(0, range.y)
  const endColExclusive = Math.max(startCol, Math.min(MAX_COLS, range.x + range.width))
  const endRowExclusive = Math.max(startRow, Math.min(MAX_ROWS, range.y + range.height))
  if (startCol >= endColExclusive || startRow >= endRowExclusive) {
    return
  }

  appendRangeSelectionFillAndGridlineRects(
    rects,
    geometry,
    { x: startCol, y: startRow, width: endColExclusive - startCol, height: endRowExclusive - startRow },
    keyPrefix,
  )
}

function appendRangeSelectionFillAndGridlineRects(
  rects: GridSelectionVisualRect[],
  geometry: GridGeometrySnapshot,
  range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  keyPrefix: string,
): void {
  let segmentIndex = 0
  for (const bounds of geometry.rangeScreenRects(range)) {
    rects.push({
      role: 'selection-fill',
      key: stableRangeKey(keyPrefix, range, 0, segmentIndex),
      bounds,
    })
    segmentIndex += 1
  }
  appendRangeSelectionGridlineRects(rects, geometry, range, keyPrefix)
}

function appendRangeSelectionGridlineRects(
  rects: GridSelectionVisualRect[],
  geometry: GridGeometrySnapshot,
  range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  keyPrefix: string,
): void {
  const startCol = Math.max(0, range.x)
  const startRow = Math.max(0, range.y)
  const endColExclusive = Math.max(startCol, Math.min(MAX_COLS, range.x + range.width))
  const endRowExclusive = Math.max(startRow, Math.min(MAX_ROWS, range.y + range.height))
  const columnIndexes =
    endColExclusive - startCol <= DIRECT_SELECTION_GRIDLINE_AXIS_LIMIT
      ? internalIndexes(startCol, endColExclusive)
      : visibleColumnIndexes(geometry).filter((col) => col > startCol && col < endColExclusive)
  const rowIndexes =
    endRowExclusive - startRow <= DIRECT_SELECTION_GRIDLINE_AXIS_LIMIT
      ? internalIndexes(startRow, endRowExclusive)
      : visibleRowIndexes(geometry).filter((row) => row > startRow && row < endRowExclusive)

  const gridlineKeyPrefix = selectionGridlineKeyPrefix(keyPrefix)
  for (const col of columnIndexes) {
    let segmentIndex = 0
    for (const bounds of geometry.rangeScreenRects({ x: col, y: startRow, width: 1, height: endRowExclusive - startRow })) {
      appendSelectionGridlineRect(rects, `${gridlineKeyPrefix}:column:${col}:${segmentIndex}`, {
        x: bounds.x,
        y: bounds.y,
        width: 1,
        height: bounds.height,
      })
      segmentIndex += 1
    }
  }
  for (const row of rowIndexes) {
    let segmentIndex = 0
    for (const bounds of geometry.rangeScreenRects({ x: startCol, y: row, width: endColExclusive - startCol, height: 1 })) {
      appendSelectionGridlineRect(rects, `${gridlineKeyPrefix}:row:${row}:${segmentIndex}`, {
        x: bounds.x,
        y: bounds.y,
        width: bounds.width,
        height: 1,
      })
      segmentIndex += 1
    }
  }
}

function internalIndexes(start: number, endExclusive: number): readonly number[] {
  const indexes: number[] = []
  for (let index = start + 1; index < endExclusive; index += 1) {
    indexes.push(index)
  }
  return indexes
}

function selectionGridlineKeyPrefix(selectionFillKeyPrefix: string): string {
  return selectionFillKeyPrefix.startsWith('selection-fill:')
    ? `selection-gridline:${selectionFillKeyPrefix.slice('selection-fill:'.length)}`
    : `selection-gridline:${selectionFillKeyPrefix}`
}

function appendSelectionGridlineRect(rects: GridSelectionVisualRect[], key: string, bounds: Rectangle): void {
  if (bounds.width > 0 && bounds.height > 0) {
    rects.push({ role: 'selection-gridline', key, bounds })
  }
}

function appendColumnHeaderFillRect(rects: GridSelectionVisualRect[], key: string, bounds: Rectangle): void {
  appendRect(rects, 'header-fill', key, {
    x: bounds.x + 1,
    y: bounds.y,
    width: Math.max(0, bounds.width - 2),
    height: bounds.height,
  })
}

function appendRowHeaderFillRect(rects: GridSelectionVisualRect[], key: string, bounds: Rectangle): void {
  appendRect(rects, 'header-fill', key, {
    x: bounds.x,
    y: bounds.y + 1,
    width: bounds.width,
    height: Math.max(0, bounds.height - 2),
  })
}

function appendCellBorderRects(
  rects: GridSelectionVisualRect[],
  geometry: GridGeometrySnapshot,
  cell: readonly [number, number],
  role: GridSelectionVisualRectRole,
  keyPrefix: string,
  options?: {
    readonly strokeWidth?: number | undefined
  },
): void {
  let segmentIndex = 0
  for (const bounds of geometry.rangeScreenRects({ x: cell[0], y: cell[1], width: 1, height: 1 })) {
    rects.push({
      role,
      key: `${keyPrefix}:${segmentIndex}`,
      bounds,
      strokeWidth: options?.strokeWidth,
    })
    segmentIndex += 1
  }
}

function appendInsetRect(
  rects: GridSelectionVisualRect[],
  role: GridSelectionVisualRectRole,
  key: string,
  bounds: Rectangle,
  insetX: number,
  insetY: number,
): void {
  appendRect(rects, role, key, {
    x: bounds.x + insetX,
    y: bounds.y + insetY,
    width: Math.max(0, bounds.width - insetX * 2),
    height: Math.max(0, bounds.height - insetY * 2),
  })
}

function appendRect(rects: GridSelectionVisualRect[], role: GridSelectionVisualRectRole, key: string, bounds: Rectangle): void {
  if (bounds.width > 0 && bounds.height > 0) {
    rects.push({ role, key, bounds })
  }
}

function stableRangeKey(
  prefix: string,
  range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  fillIndex: number,
  segmentIndex: number,
): string {
  return `${prefix}:${range.x}:${range.y}:${range.width}:${range.height}:${fillIndex}:${segmentIndex}`
}

interface AxisSelectionRange {
  readonly start: number
  readonly endExclusive: number
}

function resolveSelectedAxisRanges(input: {
  readonly axisRanges: readonly (readonly [number, number])[]
  readonly fallbackIndex: number
  readonly fallbackRange: AxisSelectionRange | null
}): readonly AxisSelectionRange[] {
  if (input.axisRanges.length > 0) {
    return input.axisRanges.map(([start, endExclusive]) => ({ start, endExclusive }))
  }
  if (input.fallbackRange) {
    return [input.fallbackRange]
  }
  return [{ start: input.fallbackIndex, endExclusive: input.fallbackIndex + 1 }]
}

function visibleColumnIndexes(geometry: GridGeometrySnapshot): readonly number[] {
  const indexes: number[] = []
  for (let index = 0; index < geometry.camera.frozenColumnCount; index += 1) {
    if (!geometry.columns.isHidden(index) && geometry.columns.sizeOf(index) > 0) {
      indexes.push(index)
    }
  }
  const bodyRange = geometry.columns.visibleRangeForWorldRect(geometry.camera.bodyWorldX, geometry.camera.bodyViewportWidth)
  for (let index = bodyRange.startIndex; index < bodyRange.endIndexExclusive; index += 1) {
    if (index < geometry.camera.frozenColumnCount) {
      continue
    }
    if (!geometry.columns.isHidden(index) && geometry.columns.sizeOf(index) > 0) {
      indexes.push(index)
    }
  }
  return indexes
}

function visibleRowIndexes(geometry: GridGeometrySnapshot): readonly number[] {
  const indexes: number[] = []
  for (let index = 0; index < geometry.camera.frozenRowCount; index += 1) {
    if (!geometry.rows.isHidden(index) && geometry.rows.sizeOf(index) > 0) {
      indexes.push(index)
    }
  }
  const bodyRange = geometry.rows.visibleRangeForWorldRect(geometry.camera.bodyWorldY, geometry.camera.bodyViewportHeight)
  for (let index = bodyRange.startIndex; index < bodyRange.endIndexExclusive; index += 1) {
    if (index < geometry.camera.frozenRowCount) {
      continue
    }
    if (!geometry.rows.isHidden(index) && geometry.rows.sizeOf(index) > 0) {
      indexes.push(index)
    }
  }
  return indexes
}

function isIndexSelected(index: number, ranges: readonly AxisSelectionRange[]): boolean {
  return ranges.some((range) => index >= range.start && index < range.endExclusive)
}

function paneFrame(geometry: GridGeometrySnapshot, kind: GridGeometrySnapshot['camera']['panes'][number]['kind']): Rectangle | null {
  return geometry.camera.panes.find((pane) => pane.kind === kind)?.frame ?? null
}

function clipRect(rect: Rectangle, clip: Rectangle): Rectangle | null {
  const x = Math.max(rect.x, clip.x)
  const y = Math.max(rect.y, clip.y)
  const right = Math.min(rect.x + rect.width, clip.x + clip.width)
  const bottom = Math.min(rect.y + rect.height, clip.y + clip.height)
  if (right <= x || bottom <= y) {
    return null
  }
  return { x, y, width: right - x, height: bottom - y }
}

function cellInRange(cell: Item, range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>): boolean {
  return cell[0] >= range.x && cell[0] < range.x + range.width && cell[1] >= range.y && cell[1] < range.y + range.height
}
