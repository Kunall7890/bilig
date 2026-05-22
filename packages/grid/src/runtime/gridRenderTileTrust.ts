import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { Item } from '../gridTypes.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import { GRID_TEXT_METRIC_FLOAT_COUNT_V3 } from '../renderer-v3/text-run-buffer.js'
import { GRID_TILE_PACKET_V3_MAGIC, GRID_TILE_PACKET_V3_VERSION } from '../renderer-v3/tile-packet-v3.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

export interface GridRenderTileCompletenessProof {
  readonly complete: boolean
  readonly gridComplete: boolean
  readonly rectBufferComplete: boolean
  readonly textBufferComplete: boolean
  readonly textRunsComplete: boolean
  readonly packetComplete: boolean
  readonly expectedCellCount: number
  readonly actualCellCount: number | null
  readonly expectedGridBorderRectCount: number
  readonly actualGridBorderRectCount: number
  readonly missingReasons: readonly string[]
}

export function hasCompleteRenderTileGrid(tile: GridRenderTile): boolean {
  return resolveRenderTileCompletenessProof(tile).gridComplete
}

export function resolveRenderTileCompletenessProof(
  tile: GridRenderTile,
  options: {
    readonly requirePacket?: boolean | undefined
  } = {},
): GridRenderTileCompletenessProof {
  const expectedCellCount = expectedRenderTileCellCount(tile)
  const actualCellCount = tile.packet?.cellCount ?? null
  const expectedGridBorderRectCount = expectedRenderTileGridBorderCount(tile)
  const actualGridBorderRectCount = countRenderTileGridBorderRects(tile)
  const rectBufferComplete = readableRectCount(tile) >= tile.rectCount
  const textBufferComplete = Math.floor(tile.textMetrics.length / GRID_TEXT_METRIC_FLOAT_COUNT_V3) >= tile.textCount
  const textRunsComplete = tile.textRuns.length === tile.textCount
  const gridComplete = expectedGridBorderRectCount === 0 || actualGridBorderRectCount >= expectedGridBorderRectCount
  const packetComplete = options.requirePacket === true ? hasCompleteRenderTilePacket(tile, expectedCellCount) : true
  const missingReasons = [
    rectBufferComplete ? null : 'rect-buffer-shorter-than-rect-count',
    textBufferComplete ? null : 'text-metrics-shorter-than-text-count',
    textRunsComplete ? null : 'text-runs-disagree-with-text-count',
    gridComplete ? null : 'grid-border-rects-incomplete',
    packetComplete ? null : 'packet-metadata-stale-or-incomplete',
  ].filter((reason): reason is string => reason !== null)
  return {
    complete: missingReasons.length === 0,
    gridComplete,
    rectBufferComplete,
    textBufferComplete,
    textRunsComplete,
    packetComplete,
    expectedCellCount,
    actualCellCount,
    expectedGridBorderRectCount,
    actualGridBorderRectCount,
    missingReasons,
  }
}

export function tileSelectedTextNeedsLocalRefresh(
  tile: GridRenderTile | null,
  selectedCell: Item | undefined,
  selectedCellSnapshot: CellSnapshot | null | undefined,
): boolean {
  if (!selectedCell) {
    return false
  }
  const selectedRun = findSelectedTextRun(tile, selectedCell)
  const expectedText = selectedSnapshotTextHint(selectedCellSnapshot)
  if (expectedText === undefined) {
    return false
  }
  if (expectedText === null) {
    return selectedRun !== null
  }
  return selectedRun?.text !== expectedText
}

function expectedRenderTileGridBorderCount(tile: GridRenderTile): number {
  const rowCount = tile.bounds.rowEnd - tile.bounds.rowStart + 1
  const colCount = tile.bounds.colEnd - tile.bounds.colStart + 1
  return rowCount > 0 && colCount > 0 ? rowCount + colCount : 0
}

function expectedRenderTileCellCount(tile: GridRenderTile): number {
  const rowCount = tile.bounds.rowEnd - tile.bounds.rowStart + 1
  const colCount = tile.bounds.colEnd - tile.bounds.colStart + 1
  return rowCount > 0 && colCount > 0 ? rowCount * colCount : 0
}

function readableRectCount(tile: GridRenderTile): number {
  return Math.min(tile.rectCount, Math.floor(tile.rectInstances.length / GRID_RECT_INSTANCE_FLOAT_COUNT_V3))
}

function countRenderTileGridBorderRects(tile: GridRenderTile): number {
  const rectCount = readableRectCount(tile)
  let count = 0
  for (let index = 0; index < rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const width = tile.rectInstances[offset + 2] ?? 0
    const height = tile.rectInstances[offset + 3] ?? 0
    const borderAlpha = tile.rectInstances[offset + 11] ?? 0
    const borderThickness = tile.rectInstances[offset + 13] ?? 0
    if (borderAlpha > 0 && borderThickness > 0 && ((width <= 1.5 && height > 0) || (height <= 1.5 && width > 0))) {
      count += 1
    }
  }
  return count
}

function hasCompleteRenderTilePacket(tile: GridRenderTile, expectedCellCount: number): boolean {
  const packet = tile.packet
  return (
    packet !== undefined &&
    packet.magic === GRID_TILE_PACKET_V3_MAGIC &&
    packet.version === GRID_TILE_PACKET_V3_VERSION &&
    packet.tileKey === tile.tileId &&
    packet.cellCount === expectedCellCount &&
    packet.rectInstanceCount === tile.rectCount &&
    packet.textRunCount === tile.textCount &&
    packet.rowStart === tile.bounds.rowStart &&
    packet.rowEnd === tile.bounds.rowEnd &&
    packet.colStart === tile.bounds.colStart &&
    packet.colEnd === tile.bounds.colEnd &&
    packet.valueSeq === tile.version.values &&
    packet.styleSeq === tile.version.styles &&
    packet.textSeq === tile.version.text &&
    packet.axisSeqX === tile.version.axisX &&
    packet.axisSeqY === tile.version.axisY &&
    packet.freezeSeq === tile.version.freeze
  )
}

function selectedSnapshotTextHint(snapshot: CellSnapshot | null | undefined): string | null | undefined {
  if (!snapshot) {
    return undefined
  }
  if (snapshot.input !== undefined && snapshot.input !== null && snapshot.input !== '') {
    return String(snapshot.input)
  }
  if (snapshot.formula !== undefined && snapshot.formula.length > 0) {
    return snapshot.formula
  }
  if (snapshot.value.tag === ValueTag.String) {
    return snapshot.value.value
  }
  if (snapshot.value.tag === ValueTag.Number) {
    return String(snapshot.value.value)
  }
  if (isDefaultPlaceholderEmptySnapshot(snapshot)) {
    return undefined
  }
  return null
}

function isDefaultPlaceholderEmptySnapshot(snapshot: CellSnapshot): boolean {
  return (
    snapshot.value.tag === ValueTag.Empty &&
    snapshot.version === 0 &&
    snapshot.flags === 0 &&
    snapshot.formula === undefined &&
    (snapshot.input === undefined || snapshot.input === '') &&
    snapshot.format === undefined &&
    snapshot.styleId === undefined &&
    snapshot.numberFormatId === undefined
  )
}

function findSelectedTextRun(tile: GridRenderTile | null, selectedCell: Item | undefined): { readonly text: string } | null {
  if (!tile || !selectedCell) {
    return null
  }
  return tile.textRuns.find((run) => run.col === selectedCell[0] && run.row === selectedCell[1] && run.text.length > 0) ?? null
}
