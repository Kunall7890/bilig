import { formatAddress } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import { createBorderRects } from './gridGpuBorders.js'
import { parseGpuColor, type GridGpuColor, type GridGpuRect, type GridGpuScene } from './gridGpuPrimitives.js'
import { getVisibleColumnBounds, getVisibleRowBounds, type GridMetrics } from './gridMetrics.js'
import { buildGridGpuHeaderScene } from './gridGpuHeaderScene.js'
import type { HeaderSelection } from './gridPointer.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import { resolveMergedCell, resolveMergedCellBounds } from './gridMergedRanges.js'
import { collectVisibleColumnBounds, collectVisibleRowBounds } from './visibleGridAxes.js'
import { workbookThemeColors } from './workbookTheme.js'

function appendRects(target: GridGpuRect[], rects: readonly GridGpuRect[]): void {
  for (const rect of rects) {
    target.push(rect)
  }
}

interface BuildGridGpuSceneOptions {
  readonly contentMode?: 'combined' | 'headers' | 'data'
  readonly engine: GridEngineLike
  readonly sheetName: string
  readonly visibleItems: readonly Item[]
  readonly visibleRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows?: number
    readonly freezeCols?: number
  }
  readonly gridMetrics: GridMetrics
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights?: Readonly<Record<number, number>>
  readonly hostBounds: Pick<DOMRect, 'left' | 'top'>
  readonly getCellBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridSelection: GridSelection
  readonly selectedCell: Item
  readonly selectionRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly hoveredCell?: Item | null
  readonly hoveredHeader?: HeaderSelection | null
  readonly resizeGuideColumn?: number | null
  readonly resizeGuideRow?: number | null
  readonly activeHeaderDrag?: HeaderSelection | null
  readonly includeLeadingGridLines?: boolean | undefined
}

const GRID_LINE_COLOR = parseGpuColor(workbookThemeColors.gridBorder)
const HEADER_FILL_COLOR = parseGpuColor(workbookThemeColors.surfaceSubtle)
const HEADER_SELECTED_FILL_COLOR = parseGpuColor(workbookThemeColors.selectionHeaderFill)
const HEADER_HOVER_FILL_COLOR = parseGpuColor(workbookThemeColors.muted)
const HEADER_DRAG_ANCHOR_FILL_COLOR = parseGpuColor('rgba(33, 86, 58, 0.18)')
const SELECTION_FILL_COLOR = parseGpuColor(workbookThemeColors.selectionFill)
const SELECTION_OUTLINE_COLOR = parseGpuColor(workbookThemeColors.selectionAccent)
const HOVER_FILL_COLOR = parseGpuColor(workbookThemeColors.hoverFill)
const HOVER_OUTLINE_COLOR = parseGpuColor(workbookThemeColors.hoverOutline)
const RESIZE_GUIDE_COLOR = parseGpuColor('rgba(33, 86, 58, 0.72)')
const RESIZE_GUIDE_GLOW_COLOR = parseGpuColor('rgba(191, 213, 196, 0.28)')
const RESIZE_GUIDE_CORE_THICKNESS = 1
const RESIZE_GUIDE_GLOW_THICKNESS = 3
const CHECKBOX_BORDER_COLOR = parseGpuColor(workbookThemeColors.textMuted)
const CHECKBOX_SURFACE_COLOR = parseGpuColor(workbookThemeColors.surface)
const CHECKBOX_SELECTED_COLOR = parseGpuColor(workbookThemeColors.accent)
const CHECKBOX_CHECK_COLOR = parseGpuColor(workbookThemeColors.surface)

export function buildGridGpuScene({
  contentMode = 'combined',
  engine,
  sheetName,
  visibleItems,
  visibleRegion,
  gridMetrics,
  columnWidths,
  rowHeights = {},
  hostBounds,
  getCellBounds,
  gridSelection,
  selectedCell,
  selectionRange = null,
  hoveredCell = null,
  hoveredHeader = null,
  resizeGuideColumn = null,
  resizeGuideRow = null,
  activeHeaderDrag = null,
  includeLeadingGridLines = true,
}: BuildGridGpuSceneOptions): GridGpuScene {
  const fillRects: GridGpuRect[] = []
  const borderRects: GridGpuRect[] = []
  const explicitBorderRects: GridGpuRect[] = []
  const renderHeaders = contentMode !== 'data'
  const renderData = contentMode !== 'headers'
  const hasFrozenAxes = contentMode === 'data' || (visibleRegion.freezeRows ?? 0) > 0 || (visibleRegion.freezeCols ?? 0) > 0
  const visibleColumnBounds = hasFrozenAxes
    ? collectVisibleColumnBounds(visibleItems, getCellBounds, gridMetrics)
    : getVisibleColumnBounds(
        visibleRegion.range,
        gridMetrics.rowMarkerWidth - visibleRegion.tx,
        Number.MAX_SAFE_INTEGER,
        columnWidths,
        gridMetrics.columnWidth,
      )
  const visibleRowBounds = hasFrozenAxes
    ? collectVisibleRowBounds(visibleItems, getCellBounds, gridMetrics)
    : getVisibleRowBounds(
        visibleRegion.range,
        gridMetrics.headerHeight - visibleRegion.ty,
        Number.MAX_SAFE_INTEGER,
        rowHeights,
        gridMetrics.rowHeight,
      )
  if (renderHeaders) {
    const headerScene = buildGridGpuHeaderScene({
      palette: {
        gridLineColor: GRID_LINE_COLOR,
        headerFillColor: HEADER_FILL_COLOR,
        headerSelectedFillColor: HEADER_SELECTED_FILL_COLOR,
        headerHoverFillColor: HEADER_HOVER_FILL_COLOR,
        headerDragAnchorFillColor: HEADER_DRAG_ANCHOR_FILL_COLOR,
        selectionFillColor: SELECTION_FILL_COLOR,
        resizeGuideColor: RESIZE_GUIDE_COLOR,
        resizeGuideGlowColor: RESIZE_GUIDE_GLOW_COLOR,
      },
      columnWidths,
      gridMetrics,
      gridSelection,
      rowHeights,
      activeHeaderDrag,
      hoveredHeader,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      selectedCell,
      selectionRange,
      visibleRegion,
      visibleItems,
      getCellBounds,
    })
    appendRects(fillRects, headerScene.fillRects)
    appendRects(borderRects, headerScene.borderRects)
  }
  if (visibleItems.length === 0 || !renderData) {
    return {
      fillRects,
      borderRects,
    }
  }

  const visibleCols = visibleItems.map(([col]) => col)
  const visibleRows = visibleItems.map(([, row]) => row)
  const visibleMinCol = Math.min(...visibleCols)
  const visibleMaxCol = Math.max(...visibleCols)
  const visibleMinRow = Math.min(...visibleRows)
  const visibleMaxRow = Math.max(...visibleRows)
  const selectionOutlineRange =
    gridSelection.columns.length > 0 || gridSelection.rows.length > 0
      ? { x: selectedCell[0], y: selectedCell[1], width: 1, height: 1 }
      : selectionRange

  for (const [col, row] of visibleItems) {
    const bounds = getCellBounds(col, row)
    if (!bounds) {
      continue
    }
    const merged = resolveMergedCell(engine, sheetName, row, col)
    if (merged && !merged.isAnchor) {
      continue
    }
    const cellBounds = merged ? resolveMergedCellBounds({ merged, fallback: bounds, getCellBounds }) : bounds
    const rect = {
      x: cellBounds.x - hostBounds.left,
      y: cellBounds.y - hostBounds.top,
      width: cellBounds.width,
      height: cellBounds.height,
    }

    const snapshot = engine.getCell(sheetName, formatAddress(row, col))
    const style = engine.getCellStyle(snapshot.styleId)

    if (style?.fill?.backgroundColor) {
      fillRects.push({
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
        color: parseGpuColor(style.fill.backgroundColor),
      })
    }

    pushGridLineRects(
      borderRects,
      rect,
      row,
      col,
      visibleMinRow,
      visibleMinCol,
      includeLeadingGridLines,
      renderHeaders ? gridMetrics : null,
    )

    if (snapshot.value.tag === ValueTag.Boolean) {
      pushBooleanCellRects(fillRects, borderRects, rect, snapshot.value.value)
    }

    if (!style?.borders) {
      continue
    }

    const borderEntries = [
      ['top', style.borders.top],
      ['right', style.borders.right],
      ['bottom', style.borders.bottom],
      ['left', style.borders.left],
    ] as const

    for (const [side, border] of borderEntries) {
      if (!border) {
        continue
      }
      explicitBorderRects.push(...createBorderRects(rect, side, border))
    }
  }

  if (selectionOutlineRange) {
    pushSelectionRects({
      borderRects,
      fillRects,
      getCellBounds,
      hostBounds,
      selectionRange: selectionOutlineRange,
      visibleMaxCol,
      visibleMaxRow,
      visibleMinCol,
      visibleMinRow,
    })
  }

  // Keep authored cell borders visible when the active selection sits on the
  // same edge; otherwise the selection outline completely hides border changes.
  borderRects.push(...explicitBorderRects)

  if (activeHeaderDrag?.kind === 'column' && gridSelection.columns.length > 0) {
    pushColumnHeaderDragGuideRectsTopLayer({
      borderRects,
      resizeGuideColor: RESIZE_GUIDE_COLOR,
      selectedColumns: {
        start: gridSelection.columns.first() ?? selectedCell[0],
        end: gridSelection.columns.last() ?? selectedCell[0],
      },
      visibleColumns: visibleColumnBounds,
      visibleRows: visibleRowBounds,
      gridMetrics,
    })
  }

  if (activeHeaderDrag?.kind === 'row' && gridSelection.rows.length > 0) {
    pushRowHeaderDragGuideRectsTopLayer({
      borderRects,
      resizeGuideColor: RESIZE_GUIDE_COLOR,
      selectedRows: {
        start: gridSelection.rows.first() ?? selectedCell[1],
        end: gridSelection.rows.last() ?? selectedCell[1],
      },
      visibleRows: visibleRowBounds,
      visibleWidth: visibleColumnBounds.length === 0 ? 0 : visibleColumnBounds.at(-1)!.right - gridMetrics.rowMarkerWidth,
      gridMetrics,
    })
  }

  if (hoveredCell) {
    pushHoveredCellRects({
      borderRects,
      fillRects,
      getCellBounds,
      hostBounds,
      hoveredCell,
      selectionRange,
      gridSelection,
    })
  }

  if (resizeGuideColumn !== null) {
    pushResizeGuideRectsTopLayer({
      borderRects,
      fillRects,
      gridMetrics,
      resizeGuideColumn,
      resizeGuideColor: RESIZE_GUIDE_COLOR,
      resizeGuideGlowColor: RESIZE_GUIDE_GLOW_COLOR,
      visibleColumns: visibleColumnBounds,
      visibleRows: visibleRowBounds,
    })
  }

  if (resizeGuideRow !== null) {
    pushRowResizeGuideRectsTopLayer({
      borderRects,
      fillRects,
      gridMetrics,
      resizeGuideRow,
      resizeGuideColor: RESIZE_GUIDE_COLOR,
      resizeGuideGlowColor: RESIZE_GUIDE_GLOW_COLOR,
      visibleColumns: visibleColumnBounds,
      visibleRows: visibleRowBounds,
    })
  }

  return {
    fillRects,
    borderRects,
  }
}

function pushColumnHeaderDragGuideRectsTopLayer(options: {
  borderRects: GridGpuRect[]
  resizeGuideColor: GridGpuColor
  selectedColumns: { start: number; end: number }
  visibleColumns: ReadonlyArray<{
    index: number
    left: number
    right: number
    width: number
  }>
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
  gridMetrics: GridMetrics
}) {
  const { borderRects, resizeGuideColor, selectedColumns, visibleColumns, visibleRows, gridMetrics } = options
  const startColumn = visibleColumns.find((entry) => entry.index === selectedColumns.start)
  const endColumn = visibleColumns.find((entry) => entry.index === selectedColumns.end)
  if (!startColumn || !endColumn) {
    return
  }
  const totalHeight = visibleRows.length === 0 ? gridMetrics.headerHeight : visibleRows.at(-1)!.bottom
  borderRects.push(
    {
      x: startColumn.left,
      y: 0,
      width: 1,
      height: totalHeight,
      color: resizeGuideColor,
    },
    {
      x: endColumn.right - 1,
      y: 0,
      width: 1,
      height: totalHeight,
      color: resizeGuideColor,
    },
  )
}

function pushRowHeaderDragGuideRectsTopLayer(options: {
  borderRects: GridGpuRect[]
  resizeGuideColor: GridGpuColor
  selectedRows: { start: number; end: number }
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
  visibleWidth: number
  gridMetrics: GridMetrics
}) {
  const { borderRects, resizeGuideColor, selectedRows, visibleRows, visibleWidth, gridMetrics } = options
  if (visibleWidth <= 0) {
    return
  }
  const startRow = visibleRows.find((entry) => entry.index === selectedRows.start)
  const endRow = visibleRows.find((entry) => entry.index === selectedRows.end)
  if (!startRow || !endRow) {
    return
  }
  const totalWidth = gridMetrics.rowMarkerWidth + visibleWidth
  borderRects.push(
    {
      x: 0,
      y: startRow.top,
      width: totalWidth,
      height: 1,
      color: resizeGuideColor,
    },
    {
      x: 0,
      y: endRow.bottom - 1,
      width: totalWidth,
      height: 1,
      color: resizeGuideColor,
    },
  )
}

function pushGridLineRects(
  borderRects: GridGpuRect[],
  rect: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  row: number,
  col: number,
  visibleMinRow: number,
  visibleMinCol: number,
  includeLeadingGridLines: boolean,
  headerOwnedPaneSeams: Pick<GridMetrics, 'headerHeight' | 'rowMarkerWidth'> | null,
) {
  const leadingHorizontalOwnedByHeader =
    headerOwnedPaneSeams !== null && row === visibleMinRow && sameCssPixel(rect.y, headerOwnedPaneSeams.headerHeight)
  const leadingVerticalOwnedByHeader =
    headerOwnedPaneSeams !== null && col === visibleMinCol && sameCssPixel(rect.x, headerOwnedPaneSeams.rowMarkerWidth)
  if (includeLeadingGridLines && row === visibleMinRow && !leadingHorizontalOwnedByHeader) {
    borderRects.push({
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: 1,
      color: GRID_LINE_COLOR,
    })
  }
  if (includeLeadingGridLines && col === visibleMinCol && !leadingVerticalOwnedByHeader) {
    borderRects.push({
      x: rect.x,
      y: rect.y,
      width: 1,
      height: rect.height,
      color: GRID_LINE_COLOR,
    })
  }
  borderRects.push({
    x: rect.x,
    y: rect.y + rect.height - 1,
    width: rect.width,
    height: 1,
    color: GRID_LINE_COLOR,
  })
  borderRects.push({
    x: rect.x + rect.width - 1,
    y: rect.y,
    width: 1,
    height: rect.height,
    color: GRID_LINE_COLOR,
  })
}

function sameCssPixel(left: number, right: number): boolean {
  return Math.abs(left - right) < 0.001
}

function pushSelectionRects(options: {
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  getCellBounds: (col: number, row: number) => Rectangle | undefined
  hostBounds: Pick<DOMRect, 'left' | 'top'>
  outlineColor?: GridGpuColor
  selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  visibleMaxCol: number
  visibleMaxRow: number
  visibleMinCol: number
  visibleMinRow: number
}) {
  const {
    borderRects,
    fillRects,
    getCellBounds,
    hostBounds,
    outlineColor = SELECTION_OUTLINE_COLOR,
    selectionRange,
    visibleMaxCol,
    visibleMaxRow,
    visibleMinCol,
    visibleMinRow,
  } = options
  const startCol = Math.max(selectionRange.x, visibleMinCol)
  const startRow = Math.max(selectionRange.y, visibleMinRow)
  const endCol = Math.min(selectionRange.x + selectionRange.width - 1, visibleMaxCol)
  const endRow = Math.min(selectionRange.y + selectionRange.height - 1, visibleMaxRow)
  if (startCol > endCol || startRow > endRow) {
    return
  }

  const startBounds = getCellBounds(startCol, startRow)
  const endBounds = getCellBounds(endCol, endRow)
  if (!startBounds || !endBounds) {
    return
  }

  const selectionRect = {
    x: startBounds.x - hostBounds.left,
    y: startBounds.y - hostBounds.top,
    width: endBounds.x + endBounds.width - startBounds.x,
    height: endBounds.y + endBounds.height - startBounds.y,
  }

  if (selectionRange.width > 1 || selectionRange.height > 1) {
    pushSelectionCellInteriorFillRects({
      endCol,
      endRow,
      fillRects,
      getCellBounds,
      hostBounds,
      startCol,
      startRow,
    })
  }

  const outlineThickness = 2
  borderRects.push(
    {
      x: selectionRect.x,
      y: selectionRect.y,
      width: selectionRect.width,
      height: outlineThickness,
      color: outlineColor,
    },
    {
      x: selectionRect.x,
      y: selectionRect.y + selectionRect.height - outlineThickness,
      width: selectionRect.width,
      height: outlineThickness,
      color: outlineColor,
    },
    {
      x: selectionRect.x,
      y: selectionRect.y,
      width: outlineThickness,
      height: selectionRect.height,
      color: outlineColor,
    },
    {
      x: selectionRect.x + selectionRect.width - outlineThickness,
      y: selectionRect.y,
      width: outlineThickness,
      height: selectionRect.height,
      color: outlineColor,
    },
  )
}

function pushSelectionCellInteriorFillRects(options: {
  fillRects: GridGpuRect[]
  getCellBounds: (col: number, row: number) => Rectangle | undefined
  hostBounds: Pick<DOMRect, 'left' | 'top'>
  startCol: number
  startRow: number
  endCol: number
  endRow: number
}): void {
  const { endCol, endRow, fillRects, getCellBounds, hostBounds, startCol, startRow } = options
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      const bounds = getCellBounds(col, row)
      if (!bounds) {
        continue
      }
      fillRects.push({
        x: bounds.x - hostBounds.left + 1,
        y: bounds.y - hostBounds.top + 1,
        width: Math.max(0, bounds.width - 2),
        height: Math.max(0, bounds.height - 2),
        color: SELECTION_FILL_COLOR,
      })
    }
  }
}

function pushBooleanCellRects(
  fillRects: GridGpuRect[],
  borderRects: GridGpuRect[],
  rect: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  checked: boolean,
): void {
  const size = Math.max(12, Math.min(16, Math.floor(Math.min(rect.width, rect.height) - 8)))
  const left = Math.round(rect.x + (rect.width - size) / 2)
  const top = Math.round(rect.y + (rect.height - size) / 2)
  const outline = 1
  const surfaceColor = checked ? CHECKBOX_SELECTED_COLOR : CHECKBOX_SURFACE_COLOR
  const borderColor = checked ? CHECKBOX_SELECTED_COLOR : CHECKBOX_BORDER_COLOR

  fillRects.push({
    x: left + outline,
    y: top + outline,
    width: Math.max(0, size - outline * 2),
    height: Math.max(0, size - outline * 2),
    color: surfaceColor,
  })
  borderRects.push(
    {
      x: left,
      y: top,
      width: size,
      height: outline,
      color: borderColor,
    },
    {
      x: left,
      y: top + size - outline,
      width: size,
      height: outline,
      color: borderColor,
    },
    {
      x: left,
      y: top,
      width: outline,
      height: size,
      color: borderColor,
    },
    {
      x: left + size - outline,
      y: top,
      width: outline,
      height: size,
      color: borderColor,
    },
  )

  if (!checked) {
    return
  }

  fillRects.push(
    {
      x: left + 3,
      y: top + 8,
      width: 2,
      height: 3,
      color: CHECKBOX_CHECK_COLOR,
    },
    {
      x: left + 5,
      y: top + 10,
      width: 2,
      height: 2,
      color: CHECKBOX_CHECK_COLOR,
    },
    {
      x: left + 7,
      y: top + 5,
      width: 2,
      height: 7,
      color: CHECKBOX_CHECK_COLOR,
    },
  )
}

function pushHoveredCellRects(options: {
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  getCellBounds: (col: number, row: number) => Rectangle | undefined
  gridSelection: GridSelection
  hostBounds: Pick<DOMRect, 'left' | 'top'>
  hoveredCell: Item
  selectionRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
}) {
  const { borderRects, fillRects, getCellBounds, gridSelection, hostBounds, hoveredCell, selectionRange } = options
  if (
    selectionRange &&
    hoveredCell[0] >= selectionRange.x &&
    hoveredCell[0] < selectionRange.x + selectionRange.width &&
    hoveredCell[1] >= selectionRange.y &&
    hoveredCell[1] < selectionRange.y + selectionRange.height
  ) {
    return
  }
  if (gridSelection.columns.length > 0 || gridSelection.rows.length > 0) {
    return
  }
  const bounds = getCellBounds(hoveredCell[0], hoveredCell[1])
  if (!bounds) {
    return
  }
  const rect = {
    x: bounds.x - hostBounds.left,
    y: bounds.y - hostBounds.top,
    width: bounds.width,
    height: bounds.height,
  }
  fillRects.push({
    x: rect.x + 1,
    y: rect.y + 1,
    width: Math.max(0, rect.width - 2),
    height: Math.max(0, rect.height - 2),
    color: HOVER_FILL_COLOR,
  })
  borderRects.push(
    {
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: 1,
      color: HOVER_OUTLINE_COLOR,
    },
    {
      x: rect.x,
      y: rect.y + rect.height - 1,
      width: rect.width,
      height: 1,
      color: HOVER_OUTLINE_COLOR,
    },
    {
      x: rect.x,
      y: rect.y,
      width: 1,
      height: rect.height,
      color: HOVER_OUTLINE_COLOR,
    },
    {
      x: rect.x + rect.width - 1,
      y: rect.y,
      width: 1,
      height: rect.height,
      color: HOVER_OUTLINE_COLOR,
    },
  )
}

function pushResizeGuideRectsTopLayer(options: {
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  resizeGuideColumn: number
  resizeGuideColor: GridGpuColor
  resizeGuideGlowColor: GridGpuColor
  visibleColumns: ReadonlyArray<{
    index: number
    left: number
    right: number
    width: number
  }>
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
}) {
  const { borderRects, fillRects, gridMetrics, resizeGuideColumn, resizeGuideColor, resizeGuideGlowColor, visibleColumns, visibleRows } =
    options
  const column = visibleColumns.find((entry) => entry.index === resizeGuideColumn)
  if (!column) {
    return
  }
  const lineX = column.right - 1
  const totalHeight = visibleRows.length === 0 ? gridMetrics.headerHeight : visibleRows.at(-1)!.bottom
  fillRects.push({
    x: lineX - Math.floor(RESIZE_GUIDE_GLOW_THICKNESS / 2),
    y: 0,
    width: RESIZE_GUIDE_GLOW_THICKNESS,
    height: totalHeight,
    color: resizeGuideGlowColor,
  })
  borderRects.push({
    x: lineX,
    y: 0,
    width: RESIZE_GUIDE_CORE_THICKNESS,
    height: totalHeight,
    color: resizeGuideColor,
  })
}

function pushRowResizeGuideRectsTopLayer(options: {
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  resizeGuideRow: number
  resizeGuideColor: GridGpuColor
  resizeGuideGlowColor: GridGpuColor
  visibleColumns: ReadonlyArray<{
    index: number
    left: number
    right: number
    width: number
  }>
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
}) {
  const { borderRects, fillRects, gridMetrics, resizeGuideRow, resizeGuideColor, resizeGuideGlowColor, visibleColumns, visibleRows } =
    options
  const row = visibleRows.find((entry) => entry.index === resizeGuideRow)
  if (!row) {
    return
  }
  const lineY = row.bottom - 1
  const totalWidth = visibleColumns.length === 0 ? gridMetrics.rowMarkerWidth : visibleColumns.at(-1)!.right
  fillRects.push({
    x: 0,
    y: lineY - Math.floor(RESIZE_GUIDE_GLOW_THICKNESS / 2),
    width: totalWidth,
    height: RESIZE_GUIDE_GLOW_THICKNESS,
    color: resizeGuideGlowColor,
  })
  borderRects.push({
    x: 0,
    y: lineY,
    width: totalWidth,
    height: RESIZE_GUIDE_CORE_THICKNESS,
    color: resizeGuideColor,
  })
}
