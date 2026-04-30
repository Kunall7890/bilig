import { useCallback, type MutableRefObject } from 'react'
import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { HeaderSelection, VisibleRegionState } from './gridPointer.js'
import type { GridSelection, Item } from './gridTypes.js'

export function useWorkbookGridPointerResolvers(input: {
  hostRef: MutableRefObject<HTMLDivElement | null>
  selectedCell: { col: number; row: number }
  gridSelection: GridSelection
  getGeometrySnapshot?: (() => GridGeometrySnapshot | null) | undefined
}) {
  const { hostRef, selectedCell, gridSelection, getGeometrySnapshot } = input

  const resolvePointerGeometry = useCallback(
    (_region?: VisibleRegionState): GridGeometrySnapshot | null => getGeometrySnapshot?.() ?? null,
    [getGeometrySnapshot],
  )

  const resolveLocalPoint = useCallback(
    (clientX: number, clientY: number): { readonly x: number; readonly y: number } | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect()
      return hostBounds ? { x: clientX - hostBounds.left, y: clientY - hostBounds.top } : null
    },
    [hostRef],
  )

  const resolveSelectedCellBounds = useCallback(
    (
      geometry: GridGeometrySnapshot,
    ): { readonly x: number; readonly y: number; readonly width: number; readonly height: number } | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect()
      const rect = geometry.cellScreenRect(selectedCell.col, selectedCell.row)
      return hostBounds && rect
        ? {
            x: hostBounds.left + rect.x,
            y: hostBounds.top + rect.y,
            width: rect.width,
            height: rect.height,
          }
        : null
    },
    [hostRef, selectedCell.col, selectedCell.row],
  )

  const resolvePointerCell = useCallback(
    (clientX: number, clientY: number, region?: VisibleRegionState, geometry?: GridGeometrySnapshot | null): Item | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      const localPoint = resolveLocalPoint(clientX, clientY)
      if (!activeGeometry || !localPoint) {
        return null
      }

      const selectedCellBounds = resolveSelectedCellBounds(activeGeometry)
      const selectionRange = gridSelection.current?.range ?? null
      if (
        selectedCellBounds &&
        gridSelection.columns.length === 0 &&
        gridSelection.rows.length === 0 &&
        selectionRange?.width === 1 &&
        selectionRange?.height === 1 &&
        clientX >= selectedCellBounds.x - 1 &&
        clientX < selectedCellBounds.x + selectedCellBounds.width &&
        clientY >= selectedCellBounds.y - 1 &&
        clientY < selectedCellBounds.y + selectedCellBounds.height
      ) {
        return [selectedCell.col, selectedCell.row]
      }

      const hit = activeGeometry.hitTestScreenPoint(localPoint)
      return hit ? [hit.col, hit.row] : null
    },
    [
      gridSelection.columns.length,
      gridSelection.current?.range,
      gridSelection.rows.length,
      resolveLocalPoint,
      resolvePointerGeometry,
      resolveSelectedCellBounds,
      selectedCell.col,
      selectedCell.row,
    ],
  )

  const resolveColumnResizeTarget = useCallback(
    (
      clientX: number,
      clientY: number,
      region?: VisibleRegionState,
      geometry?: GridGeometrySnapshot | null,
      _inputColumnWidths?: Readonly<Record<number, number>>,
      _defaultWidth?: number,
    ): number | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      const localPoint = resolveLocalPoint(clientX, clientY)
      if (!activeGeometry || !localPoint) {
        return null
      }
      const hit = activeGeometry.hitTestResizeHandleScreenPoint(localPoint)
      return hit?.kind === 'column' ? hit.index : null
    },
    [resolveLocalPoint, resolvePointerGeometry],
  )

  const resolveRowResizeTarget = useCallback(
    (
      clientX: number,
      clientY: number,
      region?: VisibleRegionState,
      geometry?: GridGeometrySnapshot | null,
      _inputRowHeights?: Readonly<Record<number, number>>,
      _defaultHeight?: number,
    ): number | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      const localPoint = resolveLocalPoint(clientX, clientY)
      if (!activeGeometry || !localPoint) {
        return null
      }
      const hit = activeGeometry.hitTestResizeHandleScreenPoint(localPoint)
      return hit?.kind === 'row' ? hit.index : null
    },
    [resolveLocalPoint, resolvePointerGeometry],
  )

  const resolveHeaderSelectionAtPointer = useCallback(
    (clientX: number, clientY: number, region?: VisibleRegionState, geometry?: GridGeometrySnapshot | null): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      const localPoint = resolveLocalPoint(clientX, clientY)
      return activeGeometry && localPoint ? activeGeometry.hitTestHeaderScreenPoint(localPoint) : null
    },
    [resolveLocalPoint, resolvePointerGeometry],
  )

  const resolveHeaderSelectionForPointerDrag = useCallback(
    (
      kind: HeaderSelection['kind'],
      clientX: number,
      clientY: number,
      region?: VisibleRegionState,
      geometry?: GridGeometrySnapshot | null,
    ): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region)
      const localPoint = resolveLocalPoint(clientX, clientY)
      return activeGeometry && localPoint ? activeGeometry.hitTestHeaderDragScreenPoint(kind, localPoint) : null
    },
    [resolveLocalPoint, resolvePointerGeometry],
  )

  return {
    resolveColumnResizeTarget,
    resolveRowResizeTarget,
    resolveHeaderSelectionAtPointer,
    resolveHeaderSelectionForPointerDrag,
    resolvePointerCell,
    resolvePointerGeometry,
  }
}
