import type { GridCameraSnapshotV2, GridPaneKind } from './gridGeometry.js'
import type { Rectangle } from './gridTypes.js'

export interface VisibleRegionState {
  range: Rectangle
  tx: number
  ty: number
  freezeRows?: number
  freezeCols?: number
}

export type HeaderSelection = { kind: 'column'; index: number } | { kind: 'row'; index: number }

export interface GridScreenPoint {
  readonly x: number
  readonly y: number
}

const DATA_PANE_KINDS: ReadonlySet<GridPaneKind> = new Set(['frozen-cells', 'frozen-rows', 'frozen-columns', 'body'])

export function clampGridScreenPointToDataPanes(camera: GridCameraSnapshotV2, point: GridScreenPoint): GridScreenPoint | null {
  let left = Number.POSITIVE_INFINITY
  let top = Number.POSITIVE_INFINITY
  let right = Number.NEGATIVE_INFINITY
  let bottom = Number.NEGATIVE_INFINITY

  for (const pane of camera.panes) {
    if (!DATA_PANE_KINDS.has(pane.kind) || pane.frame.width <= 0 || pane.frame.height <= 0) {
      continue
    }
    left = Math.min(left, pane.frame.x)
    top = Math.min(top, pane.frame.y)
    right = Math.max(right, pane.frame.x + pane.frame.width)
    bottom = Math.max(bottom, pane.frame.y + pane.frame.height)
  }

  if (!Number.isFinite(left) || !Number.isFinite(top) || !Number.isFinite(right) || !Number.isFinite(bottom)) {
    return null
  }

  return {
    x: clamp(point.x, left, Math.max(left, right - 1)),
    y: clamp(point.y, top, Math.max(top, bottom - 1)),
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value))
}
