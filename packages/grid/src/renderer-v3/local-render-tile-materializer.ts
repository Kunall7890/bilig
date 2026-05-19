import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type CellSnapshot, type Viewport } from '@bilig/protocol'
import type { GridMetrics } from '../gridMetrics.js'
import type { GridEngineLike } from '../grid-engine.js'
import type { Item } from '../gridTypes.js'
import { materializeGridRenderTileV3 } from './grid-tile-materializer.js'
import { unpackTileKey53, tileKeysForViewport } from './tile-key.js'
import type { GridRenderTile } from './render-tile-source.js'
import { DirtyMaskV3, type DirtyTileLocalSpanV3 } from './tile-damage-index.js'

export function buildLocalFixedRenderTiles(input: {
  readonly engine: GridEngineLike
  readonly sheetName: string
  readonly sheetId: number
  readonly sheetOrdinal: number
  readonly viewport: Viewport
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
  readonly dprBucket: number
  readonly generation: number
  readonly cameraSeq: number
  readonly freezeSeq?: number | undefined
  readonly tileKeys?: readonly number[] | undefined
  readonly dirtySpansForTile?: ((tileId: number) => readonly DirtyTileLocalSpanV3[]) | undefined
  readonly editingCell?: Item | null | undefined
  readonly reuseStaticGridRectsByTileId?: ReadonlyMap<number, GridRenderTile> | undefined
  readonly selectedCell?: Item | undefined
  readonly selectedCellSnapshot?: CellSnapshot | null | undefined
}): readonly GridRenderTile[] {
  const axisVersionX = hashAxisOverrides(input.sortedColumnWidthOverrides)
  const axisVersionY = hashAxisOverrides(input.sortedRowHeightOverrides)
  const tileKeys =
    input.tileKeys ??
    tileKeysForViewport({
      dprBucket: input.dprBucket,
      sheetOrdinal: input.sheetOrdinal,
      viewport: input.viewport,
    })
  return tileKeys.map((tileId) => {
    const key = unpackTileKey53(tileId)
    const tileViewport = viewportFromTileKey(key.rowTile, key.colTile)
    const dirtySpans = addSyntheticTextDirtySpans(input.dirtySpansForTile?.(tileId) ?? [], tileViewport, {
      editingCell: input.editingCell ?? null,
      selectedCell: input.selectedCell,
      selectedCellSnapshot: input.selectedCellSnapshot,
    })
    const dirty = packDirtyLocalSpans(dirtySpans)
    return materializeGridRenderTileV3({
      ...input,
      axisSeqX: axisVersionX,
      axisSeqY: axisVersionY,
      freezeSeq: input.freezeSeq ?? 0,
      glyphAtlasSeq: 0,
      materializedAtSeq: input.generation,
      packetSeq: input.generation,
      rectSeq: input.generation,
      reuseStaticGridRectsFrom: input.reuseStaticGridRectsByTileId?.get(tileId),
      styleSeq: input.generation,
      textSeq: input.generation,
      viewport: tileViewport,
      valueSeq: input.generation,
      ...dirty,
    })
  })
}

function addSyntheticTextDirtySpans(
  spans: readonly DirtyTileLocalSpanV3[],
  viewport: Viewport,
  input: {
    readonly editingCell?: Item | null | undefined
    readonly selectedCell?: Item | undefined
    readonly selectedCellSnapshot?: CellSnapshot | null | undefined
  },
): readonly DirtyTileLocalSpanV3[] {
  let next = spans
  if (input.editingCell) {
    next = addSyntheticTextDirtySpan(next, viewport, input.editingCell)
  }
  if (input.selectedCell && input.selectedCellSnapshot) {
    next = addSyntheticTextDirtySpan(next, viewport, input.selectedCell)
  }
  return next
}

function addSyntheticTextDirtySpan(
  spans: readonly DirtyTileLocalSpanV3[],
  viewport: Viewport,
  cell: Item,
): readonly DirtyTileLocalSpanV3[] {
  const [col, row] = cell
  if (col < viewport.colStart || col > viewport.colEnd || row < viewport.rowStart || row > viewport.rowEnd) {
    return spans
  }
  const localCol = col - viewport.colStart
  const localRow = row - viewport.rowStart
  if (
    spans.some(
      (span) =>
        localRow >= span.rowStart &&
        localRow <= span.rowEnd &&
        localCol >= span.colStart &&
        localCol <= span.colEnd &&
        (span.mask & (DirtyMaskV3.Value | DirtyMaskV3.Text)) !== 0,
    )
  ) {
    return spans
  }
  return [
    ...spans,
    {
      colEnd: localCol,
      colStart: localCol,
      mask: DirtyMaskV3.Value | DirtyMaskV3.Text,
      rowEnd: localRow,
      rowStart: localRow,
    },
  ]
}

export function viewportFromTileKey(rowTile: number, colTile: number): Viewport {
  const rowStart = rowTile * VIEWPORT_TILE_ROW_COUNT
  const colStart = colTile * VIEWPORT_TILE_COLUMN_COUNT
  return {
    colEnd: Math.min(MAX_COLS - 1, colStart + VIEWPORT_TILE_COLUMN_COUNT - 1),
    colStart,
    rowEnd: Math.min(MAX_ROWS - 1, rowStart + VIEWPORT_TILE_ROW_COUNT - 1),
    rowStart,
  }
}

function packDirtyLocalSpans(spans: readonly DirtyTileLocalSpanV3[]): {
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
} {
  if (spans.length === 0) {
    return {}
  }
  const dirtyLocalRows = new Uint32Array(spans.length * 2)
  const dirtyLocalCols = new Uint32Array(spans.length * 2)
  const dirtyMasks = new Uint32Array(spans.length)
  spans.forEach((span, index) => {
    const offset = index * 2
    dirtyLocalRows[offset] = span.rowStart
    dirtyLocalRows[offset + 1] = span.rowEnd
    dirtyLocalCols[offset] = span.colStart
    dirtyLocalCols[offset + 1] = span.colEnd
    dirtyMasks[index] = span.mask
  })
  return { dirtyLocalCols, dirtyLocalRows, dirtyMasks }
}

function hashAxisOverrides(entries: readonly (readonly [number, number])[]): number {
  if (entries.length === 0) {
    return 0
  }
  let hash = 2_166_136_261
  for (const [index, size] of entries) {
    hash = mixRevisionInteger(hash, index)
    hash = mixRevisionInteger(hash, Math.round(size * 1_000))
  }
  return hash >>> 0
}

function mixRevisionInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}
