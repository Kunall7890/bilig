import { useEffect, useMemo, useState, type CSSProperties } from 'react'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { workbookThemeColors } from './workbookTheme.js'

type VisualRectRole = 'selection-fill' | 'selection-border' | 'active-border' | 'fill-handle' | 'header-fill' | 'hover-fill'

export interface GridSelectionVisualRect {
  readonly role: VisualRectRole
  readonly bounds: Rectangle
}

export interface GridSelectionVisualOverlayProps {
  readonly geometry?: GridGeometrySnapshot | null | undefined
  readonly getGeometrySnapshot?: (() => GridGeometrySnapshot | null) | undefined
  readonly gridSelection: GridSelection
  readonly hoverCell?: Item | null | undefined
  readonly selectedCell: Item
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
  readonly scrollTransformStore?: WorkbookGridScrollStore | undefined
}

export function GridSelectionVisualOverlay(props: GridSelectionVisualOverlayProps) {
  const {
    geometry: staticGeometry,
    getGeometrySnapshot,
    gridSelection,
    hoverCell,
    scrollTransformStore,
    selectedCell,
    selectionRange,
    showFillHandle,
  } = props
  const [scrollVersion, setScrollVersion] = useState(0)
  useEffect(() => {
    if (!scrollTransformStore) {
      return
    }
    return scrollTransformStore.subscribe(() => {
      setScrollVersion((version) => version + 1)
    })
  }, [scrollTransformStore])
  const geometry = useMemo(() => {
    void scrollVersion
    return getGeometrySnapshot?.() ?? staticGeometry ?? null
  }, [getGeometrySnapshot, staticGeometry, scrollVersion])
  const rects = useMemo(
    () =>
      geometry
        ? buildGridSelectionVisualRects({
            geometry,
            gridSelection,
            hoverCell: hoverCell ?? null,
            selectedCell,
            selectionRange,
            showFillHandle,
          })
        : [],
    [geometry, gridSelection, hoverCell, selectedCell, selectionRange, showFillHandle],
  )

  if (rects.length === 0) {
    return null
  }

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-20 overflow-hidden"
      data-testid="grid-selection-visual-overlay"
    >
      {rects.map((rect) => (
        <div
          className={classNameForRole(rect.role)}
          data-grid-selection-visual-role={rect.role}
          key={keyForRect(rect)}
          style={styleForRect(rect)}
        />
      ))}
    </div>
  )
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
  for (const bounds of input.geometry.rangeScreenRects({ x: hoverCell[0], y: hoverCell[1], width: 1, height: 1 })) {
    appendInsetRect(rects, 'hover-fill', bounds, 1, 1)
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
    const activeCell = input.gridSelection.current?.cell ?? [input.selectionRange.x, input.selectionRange.y]
    appendCellBorderRects(rects, input.geometry, activeCell, 'active-border')
    return
  }

  const isMultiCellSelection = input.selectionRange.width > 1 || input.selectionRange.height > 1
  if (isMultiCellSelection) {
    for (const bounds of input.geometry.rangeScreenRects(input.selectionRange)) {
      appendInsetRect(rects, 'selection-fill', bounds, 1, 1)
    }
  }

  const activeCell = input.gridSelection.current?.cell ?? null
  if (isMultiCellSelection) {
    for (const bounds of input.geometry.rangeScreenRects(input.selectionRange)) {
      rects.push({ role: 'selection-border', bounds })
    }
  } else {
    appendCellBorderRects(rects, input.geometry, [input.selectionRange.x, input.selectionRange.y], 'active-border')
  }

  if (activeCell && isMultiCellSelection && cellInRange(activeCell, input.selectionRange)) {
    appendCellBorderRects(rects, input.geometry, activeCell, 'active-border')
  }

  if (input.showFillHandle) {
    const handle = input.geometry.fillHandleScreenRect(input.selectionRange)
    if (handle) {
      rects.push({ role: 'fill-handle', bounds: handle })
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
    if (bounds) {
      appendInsetRect(rects, 'header-fill', bounds, 1, 1)
    }
  }
  for (const index of visibleRowIndexes(input.geometry)) {
    if (!isIndexSelected(index, rowRanges)) {
      continue
    }
    const bounds = input.geometry.rowHeaderScreenRect(index)
    if (bounds) {
      appendInsetRect(rects, 'header-fill', bounds, 1, 1)
    }
  }

  if (input.gridSelection.columns.length > 0) {
    for (const range of columnRanges) {
      const start = Math.max(0, range.start)
      const endExclusive = Math.max(start + 1, Math.min(MAX_COLS, range.endExclusive))
      for (const bounds of input.geometry.rangeScreenRects({ x: start, y: 0, width: endExclusive - start, height: MAX_ROWS })) {
        appendInsetRect(rects, 'selection-fill', bounds, 1, 1)
      }
    }
  }
  if (input.gridSelection.rows.length > 0) {
    for (const range of rowRanges) {
      const start = Math.max(0, range.start)
      const endExclusive = Math.max(start + 1, Math.min(MAX_ROWS, range.endExclusive))
      for (const bounds of input.geometry.rangeScreenRects({ x: 0, y: start, width: MAX_COLS, height: endExclusive - start })) {
        appendInsetRect(rects, 'selection-fill', bounds, 1, 1)
      }
    }
  }
}

function appendCellBorderRects(
  rects: GridSelectionVisualRect[],
  geometry: GridGeometrySnapshot,
  cell: readonly [number, number],
  role: VisualRectRole,
): void {
  for (const bounds of geometry.rangeScreenRects({ x: cell[0], y: cell[1], width: 1, height: 1 })) {
    rects.push({ role, bounds })
  }
}

function appendInsetRect(rects: GridSelectionVisualRect[], role: VisualRectRole, bounds: Rectangle, insetX: number, insetY: number): void {
  const insetBounds = {
    x: bounds.x + insetX,
    y: bounds.y + insetY,
    width: Math.max(0, bounds.width - insetX * 2),
    height: Math.max(0, bounds.height - insetY * 2),
  }
  if (insetBounds.width > 0 && insetBounds.height > 0) {
    rects.push({ role, bounds: insetBounds })
  }
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

function cellInRange(cell: Item, range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>): boolean {
  return cell[0] >= range.x && cell[0] < range.x + range.width && cell[1] >= range.y && cell[1] < range.y + range.height
}

function classNameForRole(role: VisualRectRole): string {
  switch (role) {
    case 'selection-fill':
      return 'absolute box-border'
    case 'header-fill':
      return 'absolute box-border'
    case 'hover-fill':
      return 'absolute box-border'
    case 'selection-border':
      return 'absolute box-border border border-[var(--wb-accent)]'
    case 'active-border':
      return 'absolute box-border border-2 border-[var(--wb-accent)]'
    case 'fill-handle':
      return 'absolute box-border rounded-[1px] border border-[var(--wb-surface)] bg-[var(--wb-accent)]'
  }
}

function keyForRect(rect: GridSelectionVisualRect): string {
  return `${rect.role}:${rect.bounds.x}:${rect.bounds.y}:${rect.bounds.width}:${rect.bounds.height}`
}

function styleForRect(rect: GridSelectionVisualRect): CSSProperties {
  const base = {
    height: rect.bounds.height,
    left: rect.bounds.x,
    top: rect.bounds.y,
    width: rect.bounds.width,
  }
  if (rect.role === 'selection-border' || rect.role === 'active-border') {
    return {
      ...base,
      backgroundColor: 'transparent',
    }
  }
  if (rect.role === 'fill-handle') {
    return {
      ...base,
      boxShadow: `0 0 0 1px ${workbookThemeColors.accent}`,
    }
  }
  if (rect.role === 'header-fill') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.accentSoft,
    }
  }
  if (rect.role === 'hover-fill') {
    return {
      ...base,
      backgroundColor: 'rgba(31, 122, 67, 0.05)',
    }
  }
  return {
    ...base,
    backgroundColor: workbookThemeColors.selectionFill,
  }
}
