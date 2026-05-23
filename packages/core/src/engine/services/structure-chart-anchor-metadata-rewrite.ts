import { MAX_COLS, MAX_ROWS, type WorkbookChartSnapshot } from '@bilig/protocol'
import { formatAddress, type StructuralAxisTransform } from '@bilig/formula'
import { mapStructuralAxisIndex } from '../../engine-structural-utils.js'

type WorkbookChartAnchor = NonNullable<WorkbookChartSnapshot['anchor']>
type WorkbookChartAnchorMarker = Extract<WorkbookChartAnchor, { kind: 'twoCell' | 'oneCell' }>['from']

export function rewriteChartAnchorForStructuralTransform(
  anchor: WorkbookChartAnchor,
  transform: StructuralAxisTransform,
): WorkbookChartAnchor | undefined {
  switch (anchor.kind) {
    case 'twoCell': {
      const from = rewriteChartAnchorMarkerForStructuralTransform(anchor.from, transform)
      const to = rewriteChartAnchorMarkerForStructuralTransform(anchor.to, transform)
      if (!from || !to || to.row < from.row || to.col < from.col) {
        return undefined
      }
      return { ...anchor, from, to }
    }
    case 'oneCell': {
      const from = rewriteChartAnchorMarkerForStructuralTransform(anchor.from, transform)
      return from ? { ...anchor, from } : undefined
    }
    case 'absolute':
      return structuredClone(anchor)
  }
}

function rewriteChartAnchorMarkerForStructuralTransform(
  marker: WorkbookChartAnchorMarker,
  transform: StructuralAxisTransform,
): WorkbookChartAnchorMarker | undefined {
  const row = transform.axis === 'row' ? mapStructuralAxisIndex(marker.row, transform) : marker.row
  const col = transform.axis === 'column' ? mapStructuralAxisIndex(marker.col, transform) : marker.col
  if (row === undefined || col === undefined || row >= MAX_ROWS || col >= MAX_COLS) {
    return undefined
  }
  return { ...marker, row, col }
}

export function chartGeometryFromAnchor(
  anchor: WorkbookChartAnchor | undefined,
): { readonly address: string; readonly rows: number; readonly cols: number } | undefined {
  if (!anchor) {
    return undefined
  }
  if (anchor.kind === 'twoCell') {
    return {
      address: formatAddress(anchor.from.row, anchor.from.col),
      rows: Math.max(1, anchor.to.row - anchor.from.row),
      cols: Math.max(1, anchor.to.col - anchor.from.col),
    }
  }
  if (anchor.kind === 'oneCell') {
    return {
      address: formatAddress(anchor.from.row, anchor.from.col),
      rows: 1,
      cols: 1,
    }
  }
  return undefined
}
