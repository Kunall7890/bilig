import type { GridGeometrySnapshot } from '../gridGeometry.js'
import { parseGpuColor, type GridGpuRect } from '../gridGpuPrimitives.js'
import type { HeaderSelection } from '../gridPointer.js'
import { createRectangleSelectionFromRange } from '../gridSelection.js'
import { buildGridSelectionVisualRects } from '../gridSelectionVisualRects.js'
import type { GridSelection, Item, Rectangle } from '../gridTypes.js'
import { workbookThemeColors } from '../workbookTheme.js'
import { GRID_RECT_FLOAT_COUNT_V3, GRID_RECT_INSTANCE_FLOAT_COUNT_V3, packGridRectBufferV3 } from './rect-instance-buffer.js'

export const DYNAMIC_OVERLAY_RECT_FLOAT_COUNT_V3 = GRID_RECT_FLOAT_COUNT_V3
export const DYNAMIC_OVERLAY_RECT_INSTANCE_FLOAT_COUNT_V3 = GRID_RECT_INSTANCE_FLOAT_COUNT_V3

export interface DynamicGridOverlayBatchV3 {
  readonly seq: number
  readonly cameraSeq: number
  readonly generatedAt: number
  readonly sheetName: string
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly rects: Float32Array
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly fillRectCount: number
  readonly borderRectCount: number
  readonly rectSignature: string
}

export interface DynamicGridPreviewRectV3 {
  readonly role: 'target' | 'source'
  readonly bounds: Rectangle
}

export type DynamicGridSelectionOverlayModeV3 = 'all' | 'fills-only'

interface BorderSides {
  readonly bottom: boolean
  readonly left: boolean
  readonly right: boolean
  readonly top: boolean
}

export function buildDynamicGridOverlayBatchV3(input: {
  readonly geometry: GridGeometrySnapshot
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly fillPreviewRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null | undefined
  readonly previewRects?: readonly DynamicGridPreviewRectV3[] | undefined
  readonly gridSelection?: GridSelection | null | undefined
  readonly selectedCell?: Item | null | undefined
  readonly hoveredCell?: readonly [number, number] | null | undefined
  readonly showFillHandle: boolean
  readonly activeHeaderDrag?: HeaderSelection | null | undefined
  readonly showHoverOverlay?: boolean | undefined
  readonly selectionOverlayMode?: DynamicGridSelectionOverlayModeV3 | undefined
  readonly showSelectionOverlay?: boolean | undefined
  readonly showStructuralGridlines?: boolean | undefined
  readonly resizeGuideColumn?: number | null | undefined
  readonly resizeGuideColumnWidth?: number | null | undefined
  readonly resizeGuideRow?: number | null | undefined
  readonly resizeGuideRowHeight?: number | null | undefined
}): DynamicGridOverlayBatchV3 {
  const fillRects: GridGpuRect[] = []
  const borderRects: GridGpuRect[] = []
  if (input.showStructuralGridlines !== false) {
    appendStructuralGridlines({
      fillRects,
      geometry: input.geometry,
    })
  }
  if (input.showSelectionOverlay !== false) {
    const selectionOverlayMode = input.selectionOverlayMode ?? 'all'
    appendSelectionVisualOverlay({
      borderRects,
      fillRects,
      geometry: input.geometry,
      gridSelection: input.gridSelection ?? null,
      selectedCell: input.selectedCell ?? null,
      selectionRange: input.selectionRange,
      showFillHandle: input.showFillHandle,
      selectionOverlayMode,
    })
  }
  appendFillPreviewOverlay({
    fillRects,
    fillPreviewRange: input.fillPreviewRange ?? null,
    geometry: input.geometry,
  })
  appendPreviewRects({
    borderRects,
    fillRects,
    previewRects: input.previewRects ?? [],
  })
  if (input.showHoverOverlay !== false) {
    appendHoverOverlay({
      fillRects,
      geometry: input.geometry,
      hoveredCell: input.hoveredCell ?? null,
      selectionRange: input.selectionRange,
    })
  }
  appendResizeGuides({
    borderRects,
    geometry: input.geometry,
    resizeGuideColumn: input.resizeGuideColumn ?? null,
    resizeGuideColumnWidth: input.resizeGuideColumnWidth ?? null,
    resizeGuideRow: input.resizeGuideRow ?? null,
    resizeGuideRowHeight: input.resizeGuideRowHeight ?? null,
  })
  appendHeaderDragGuides({
    activeHeaderDrag: input.activeHeaderDrag ?? null,
    borderRects,
    fillRects,
    geometry: input.geometry,
    gridSelection: input.gridSelection ?? null,
  })
  appendFrozenSeparators({ borderRects, geometry: input.geometry })

  const surfaceSize = resolveOverlaySurfaceSize(input.geometry)
  const rectBuffer = packGridRectBufferV3({ borderRects, fillRects }, surfaceSize)
  return {
    ...rectBuffer,
    cameraSeq: input.geometry.camera.seq,
    generatedAt: input.geometry.camera.updatedAt,
    seq: input.geometry.camera.seq,
    sheetName: input.geometry.camera.sheetName,
    surfaceSize,
  }
}

function appendStructuralGridlines(input: { readonly geometry: GridGeometrySnapshot; readonly fillRects: GridGpuRect[] }): void {
  const parsedColor = parseGpuColor(workbookThemeColors.gridBorder)
  const color = { ...parsedColor, a: 0.58 }
  appendCellPaneStructuralGridlines({
    color,
    fillRects: input.fillRects,
    geometry: input.geometry,
    paneKind: 'body',
  })
  appendCellPaneStructuralGridlines({
    color,
    fillRects: input.fillRects,
    geometry: input.geometry,
    paneKind: 'frozen-rows',
  })
  appendCellPaneStructuralGridlines({
    color,
    fillRects: input.fillRects,
    geometry: input.geometry,
    paneKind: 'frozen-columns',
  })
  appendCellPaneStructuralGridlines({
    color,
    fillRects: input.fillRects,
    geometry: input.geometry,
    paneKind: 'frozen-cells',
  })
}

function appendCellPaneStructuralGridlines(input: {
  readonly color: GridGpuRect['color']
  readonly fillRects: GridGpuRect[]
  readonly geometry: GridGeometrySnapshot
  readonly paneKind: 'body' | 'frozen-rows' | 'frozen-columns' | 'frozen-cells'
}): void {
  const pane = input.geometry.camera.panes.find((candidate) => candidate.kind === input.paneKind)
  if (!pane || pane.frame.width <= 0 || pane.frame.height <= 0) {
    return
  }
  const thickness = resolveStructuralGridlineThickness()
  const ranges = resolveCellPaneStructuralGridlineRanges(input.geometry, input.paneKind)
  if (!ranges || ranges.colStart >= ranges.colEndExclusive || ranges.rowStart >= ranges.rowEndExclusive) {
    return
  }

  for (let row = ranges.rowStart; row < ranges.rowEndExclusive; row += 1) {
    if (input.geometry.rows.isHidden(row) || input.geometry.rows.sizeOf(row) <= 0) {
      continue
    }
    for (let col = ranges.colStart; col < ranges.colEndExclusive; col += 1) {
      if (input.geometry.columns.isHidden(col) || input.geometry.columns.sizeOf(col) <= 0) {
        continue
      }
      const rect = input.geometry.cellScreenRectForPane(col, row, input.paneKind)
      if (!rect || rect.width <= 0 || rect.height <= 0) {
        continue
      }
      input.fillRects.push({
        x: rect.x,
        y: rect.y + Math.max(0, rect.height - thickness),
        width: rect.width,
        height: Math.min(thickness, rect.height),
        color: input.color,
      })
      input.fillRects.push({
        x: rect.x + Math.max(0, rect.width - thickness),
        y: rect.y,
        width: Math.min(thickness, rect.width),
        height: rect.height,
        color: input.color,
      })
    }
  }
}

function resolveStructuralGridlineThickness(): number {
  return 1
}

function resolveCellPaneStructuralGridlineRanges(
  geometry: GridGeometrySnapshot,
  paneKind: 'body' | 'frozen-rows' | 'frozen-columns' | 'frozen-cells',
): {
  readonly colEndExclusive: number
  readonly colStart: number
  readonly rowEndExclusive: number
  readonly rowStart: number
} | null {
  const camera = geometry.camera
  const bodyColumns = geometry.columns.visibleRangeForWorldRect(camera.bodyWorldX, camera.bodyViewportWidth)
  const bodyRows = geometry.rows.visibleRangeForWorldRect(camera.bodyWorldY, camera.bodyViewportHeight)
  switch (paneKind) {
    case 'body':
      return {
        colEndExclusive: bodyColumns.endIndexExclusive,
        colStart: bodyColumns.startIndex,
        rowEndExclusive: bodyRows.endIndexExclusive,
        rowStart: bodyRows.startIndex,
      }
    case 'frozen-rows':
      return camera.frozenRowCount <= 0
        ? null
        : {
            colEndExclusive: bodyColumns.endIndexExclusive,
            colStart: bodyColumns.startIndex,
            rowEndExclusive: camera.frozenRowCount,
            rowStart: 0,
          }
    case 'frozen-columns':
      return camera.frozenColumnCount <= 0
        ? null
        : {
            colEndExclusive: camera.frozenColumnCount,
            colStart: 0,
            rowEndExclusive: bodyRows.endIndexExclusive,
            rowStart: bodyRows.startIndex,
          }
    case 'frozen-cells':
      return camera.frozenColumnCount <= 0 || camera.frozenRowCount <= 0
        ? null
        : {
            colEndExclusive: camera.frozenColumnCount,
            colStart: 0,
            rowEndExclusive: camera.frozenRowCount,
            rowStart: 0,
          }
  }
}

function appendFillPreviewOverlay(input: {
  readonly geometry: GridGeometrySnapshot
  readonly fillPreviewRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly fillRects: GridGpuRect[]
}): void {
  if (!input.fillPreviewRange) {
    return
  }
  const color = parseGpuColor(workbookThemeColors.fillPreviewFill)
  for (const rect of input.geometry.rangeScreenRects(input.fillPreviewRange)) {
    input.fillRects.push({ ...rect, color })
  }
}

function appendPreviewRects(input: {
  readonly previewRects: readonly DynamicGridPreviewRectV3[]
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  for (const preview of input.previewRects) {
    const isTarget = preview.role === 'target'
    input.fillRects.push({
      ...preview.bounds,
      color: parseGpuColor(isTarget ? 'rgba(56, 189, 248, 0.08)' : 'rgba(148, 163, 184, 0.06)'),
    })
    appendBorderRects(
      input.borderRects,
      preview.bounds,
      parseGpuColor(isTarget ? 'rgba(14, 116, 144, 0.9)' : 'rgba(100, 116, 139, 0.9)'),
      1,
    )
  }
}

function resolveOverlaySurfaceSize(geometry: GridGeometrySnapshot): { readonly width: number; readonly height: number } {
  return geometry.camera.panes.reduce(
    (size, pane) => ({
      height: Math.max(size.height, pane.frame.y + pane.frame.height),
      width: Math.max(size.width, pane.frame.x + pane.frame.width),
    }),
    { height: 0, width: 0 },
  )
}

function appendHoverOverlay(input: {
  readonly geometry: GridGeometrySnapshot
  readonly hoveredCell: readonly [number, number] | null
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly fillRects: GridGpuRect[]
}): void {
  if (!input.hoveredCell) {
    return
  }
  const [col, row] = input.hoveredCell
  if (
    input.selectionRange &&
    col >= input.selectionRange.x &&
    col < input.selectionRange.x + input.selectionRange.width &&
    row >= input.selectionRange.y &&
    row < input.selectionRange.y + input.selectionRange.height
  ) {
    return
  }
  const rect = input.geometry.cellScreenRect(col, row)
  if (!rect) {
    return
  }
  input.fillRects.push({
    x: rect.x + 1,
    y: rect.y + 1,
    width: Math.max(0, rect.width - 2),
    height: Math.max(0, rect.height - 2),
    color: parseGpuColor(workbookThemeColors.hoverFill),
  })
}

function appendSelectionVisualOverlay(input: {
  readonly geometry: GridGeometrySnapshot
  readonly gridSelection: GridSelection | null
  readonly selectedCell: Item | null
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
  readonly selectionOverlayMode: DynamicGridSelectionOverlayModeV3
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  if (!input.selectionRange) {
    return
  }
  const borderColor = parseGpuColor(workbookThemeColors.selectionAccent)
  const gridSelection = input.gridSelection ?? createRectangleSelectionFromRange(input.selectionRange)
  const selectedCell = input.selectedCell ?? gridSelection.current?.cell ?? [input.selectionRange.x, input.selectionRange.y]
  const showSelectionChrome = input.selectionOverlayMode === 'all'
  const visualRects = buildGridSelectionVisualRects({
    geometry: input.geometry,
    gridSelection,
    hoverCell: null,
    selectedCell,
    selectionRange: input.selectionRange,
    showFillHandle: showSelectionChrome && input.showFillHandle,
  })

  for (const rect of visualRects) {
    switch (rect.role) {
      case 'selection-fill':
        input.fillRects.push({ ...rect.bounds, color: parseGpuColor(workbookThemeColors.selectionFill) })
        break
      case 'selection-gridline':
        input.fillRects.push({ ...rect.bounds, color: parseGpuColor(workbookThemeColors.gridBorder) })
        break
      case 'header-fill':
        input.fillRects.push({ ...rect.bounds, color: parseGpuColor(workbookThemeColors.selectionHeaderFill) })
        break
      case 'header-seam-cover':
        input.fillRects.push({ ...rect.bounds, color: parseGpuColor(workbookThemeColors.selectionHeaderSeamFill) })
        break
      case 'hover-fill':
        input.fillRects.push({ ...rect.bounds, color: parseGpuColor(workbookThemeColors.hoverFill) })
        break
      case 'selection-border':
      case 'active-border':
        if (showSelectionChrome) {
          appendBorderRects(input.borderRects, rect.bounds, borderColor, rect.strokeWidth ?? (rect.role === 'active-border' ? 2 : 1))
        }
        break
      case 'fill-handle':
        if (showSelectionChrome) {
          input.fillRects.push({ ...rect.bounds, color: borderColor })
          appendBorderRects(input.borderRects, rect.bounds, parseGpuColor(workbookThemeColors.surface), 1)
        }
        break
    }
  }
}

function appendResizeGuides(input: {
  readonly geometry: GridGeometrySnapshot
  readonly resizeGuideColumn: number | null
  readonly resizeGuideColumnWidth: number | null
  readonly resizeGuideRow: number | null
  readonly resizeGuideRowHeight: number | null
  readonly borderRects: GridGpuRect[]
}): void {
  const color = parseGpuColor(workbookThemeColors.selectionAccent)
  const glowColor = parseGpuColor(workbookThemeColors.selectionFill)
  if (input.resizeGuideColumn !== null) {
    const rect = resolveColumnResizeGuideRect(input.geometry, input.resizeGuideColumn, input.resizeGuideColumnWidth)
    if (rect) {
      input.borderRects.push({
        x: rect.x - 1,
        y: rect.y,
        width: 3,
        height: rect.height,
        color: glowColor,
      })
      input.borderRects.push({
        ...rect,
        color,
      })
    }
  }
  if (input.resizeGuideRow !== null) {
    const rect = resolveRowResizeGuideRect(input.geometry, input.resizeGuideRow, input.resizeGuideRowHeight)
    if (rect) {
      input.borderRects.push({
        x: rect.x,
        y: rect.y - 1,
        width: rect.width,
        height: 3,
        color: glowColor,
      })
      input.borderRects.push({
        ...rect,
        color,
      })
    }
  }
}

function resolveColumnResizeGuideRect(geometry: GridGeometrySnapshot, columnIndex: number, previewWidth: number | null): Rectangle | null {
  if (previewWidth === null) {
    return geometry.resizeGuideScreenRect({ kind: 'column', index: columnIndex })
  }
  const defaultRect = geometry.resizeGuideScreenRect({ kind: 'column', index: columnIndex })
  const headerRect = geometry.columnHeaderScreenRect(columnIndex)
  if (!defaultRect || !headerRect) {
    return null
  }
  const surfaceSize = resolveOverlaySurfaceSize(geometry)
  return {
    height: surfaceSize.height,
    width: defaultRect.width,
    x: headerRect.x + Math.max(0, previewWidth) - 1,
    y: defaultRect.y,
  }
}

function resolveRowResizeGuideRect(geometry: GridGeometrySnapshot, rowIndex: number, previewHeight: number | null): Rectangle | null {
  if (previewHeight === null) {
    return geometry.resizeGuideScreenRect({ kind: 'row', index: rowIndex })
  }
  const defaultRect = geometry.resizeGuideScreenRect({ kind: 'row', index: rowIndex })
  const headerRect = geometry.rowHeaderScreenRect(rowIndex)
  if (!defaultRect || !headerRect) {
    return null
  }
  const surfaceSize = resolveOverlaySurfaceSize(geometry)
  return {
    height: defaultRect.height,
    width: surfaceSize.width,
    x: defaultRect.x,
    y: headerRect.y + Math.max(0, previewHeight) - 1,
  }
}

function appendHeaderDragGuides(input: {
  readonly geometry: GridGeometrySnapshot
  readonly gridSelection: GridSelection | null
  readonly activeHeaderDrag: HeaderSelection | null
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  if (!input.activeHeaderDrag || !input.gridSelection) {
    return
  }
  const color = parseGpuColor(workbookThemeColors.selectionAccent)
  const host = hostRect(input.geometry)
  if (input.activeHeaderDrag.kind === 'column' && input.gridSelection.columns.length > 0) {
    const start = input.gridSelection.columns.first()
    const end = input.gridSelection.columns.last()
    if (start === undefined || end === undefined) {
      return
    }
    const startRect = input.geometry.columnHeaderScreenRect(start)
    const endRect = input.geometry.columnHeaderScreenRect(end)
    if (startRect && endRect) {
      input.borderRects.push(
        { x: startRect.x, y: 0, width: 1, height: host.height, color },
        { x: endRect.x + endRect.width - 1, y: 0, width: 1, height: host.height, color },
      )
    }
    const activeRect = input.geometry.columnHeaderScreenRect(input.activeHeaderDrag.index)
    if (activeRect) {
      input.fillRects.push({ x: activeRect.x, y: Math.max(0, activeRect.height - 3), width: activeRect.width, height: 3, color })
    }
  }
  if (input.activeHeaderDrag.kind === 'row' && input.gridSelection.rows.length > 0) {
    const start = input.gridSelection.rows.first()
    const end = input.gridSelection.rows.last()
    if (start === undefined || end === undefined) {
      return
    }
    const startRect = input.geometry.rowHeaderScreenRect(start)
    const endRect = input.geometry.rowHeaderScreenRect(end)
    if (startRect && endRect) {
      input.borderRects.push(
        { x: 0, y: startRect.y, width: host.width, height: 1, color },
        { x: 0, y: endRect.y + endRect.height - 1, width: host.width, height: 1, color },
      )
    }
    const activeRect = input.geometry.rowHeaderScreenRect(input.activeHeaderDrag.index)
    if (activeRect) {
      input.fillRects.push({ x: Math.max(0, activeRect.width - 3), y: activeRect.y, width: 3, height: activeRect.height, color })
    }
  }
}

function appendFrozenSeparators(input: { readonly geometry: GridGeometrySnapshot; readonly borderRects: GridGpuRect[] }): void {
  const color = parseGpuColor(workbookThemeColors.border)
  const hostWidth =
    input.geometry.camera.bodyViewportWidth +
    input.geometry.camera.frozenWidth +
    (input.geometry.camera.panes.find((pane) => pane.kind === 'row-header-body')?.frame.width ?? 0)
  const hostHeight =
    input.geometry.camera.bodyViewportHeight +
    input.geometry.camera.frozenHeight +
    (input.geometry.camera.panes.find((pane) => pane.kind === 'column-header-body')?.frame.height ?? 0)
  if (input.geometry.camera.frozenWidth > 0) {
    const x = input.geometry.camera.panes.find((pane) => pane.kind === 'body')?.frame.x ?? 0
    input.borderRects.push({ x: x - 1, y: 0, width: 1, height: hostHeight, color })
  }
  if (input.geometry.camera.frozenHeight > 0) {
    const y = input.geometry.camera.panes.find((pane) => pane.kind === 'body')?.frame.y ?? 0
    input.borderRects.push({ x: 0, y: y - 1, width: hostWidth, height: 1, color })
  }
}

function hostRect(geometry: GridGeometrySnapshot): Rectangle {
  return geometry.camera.panes.reduce(
    (current, pane) => ({
      x: 0,
      y: 0,
      width: Math.max(current.width, pane.frame.x + pane.frame.width),
      height: Math.max(current.height, pane.frame.y + pane.frame.height),
    }),
    { x: 0, y: 0, width: 0, height: 0 },
  )
}

function appendBorderRects(target: GridGpuRect[], rect: Rectangle, color: GridGpuRect['color'], thickness: number): void {
  appendBorderRectsForSides(target, rect, color, thickness, { bottom: true, left: true, right: true, top: true })
}

function appendBorderRectsForSides(
  target: GridGpuRect[],
  rect: Rectangle,
  color: GridGpuRect['color'],
  thickness: number,
  sides: BorderSides,
): void {
  const nextRects: GridGpuRect[] = []
  if (sides.top) {
    nextRects.push({ x: rect.x, y: rect.y, width: rect.width, height: thickness, color })
  }
  if (sides.bottom) {
    nextRects.push({ x: rect.x, y: rect.y + rect.height - thickness, width: rect.width, height: thickness, color })
  }
  if (sides.left) {
    nextRects.push({ x: rect.x, y: rect.y, width: thickness, height: rect.height, color })
  }
  if (sides.right) {
    nextRects.push({ x: rect.x + rect.width - thickness, y: rect.y, width: thickness, height: rect.height, color })
  }
  target.push(...nextRects.filter((candidate) => candidate.width > 0 && candidate.height > 0))
}
