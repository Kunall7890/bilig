import type { TextQuadRunSpan } from './line-text-quad-buffer.js'
import type { GridRenderTile } from './render-tile-source.js'
import { DirtyMaskV3 } from './tile-damage-index.js'
import type { TypeGpuTileTextRevisionKeyV3 } from './typegpu-tile-resource-revisions.js'
import { isCellKeyDirtyForTile } from './typegpu-tile-text-patch.js'

const TEXT_CELL_SKIP_FORBIDDEN_DIRTY_MASK = DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze

type GridRenderTextRun = GridRenderTile['textRuns'][number]

interface TextResourceSkipContent {
  readonly textHandle: object | null
  readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3 | null
  readonly textRunCount: number
}

export function buildDecorationCellKeys(textRuns: readonly GridRenderTextRun[]): ReadonlySet<string> {
  const keys = new Set<string>()
  for (const run of textRuns) {
    if (!run.underline && !run.strike) {
      continue
    }
    if (run.row === undefined || run.col === undefined) {
      return new Set(['*'])
    }
    keys.add(formatDecorationCellKey(run.row, run.col))
  }
  return keys
}

export function shouldSyncTextDecorationRects(input: {
  readonly previousDecorationCellKeys: ReadonlySet<string> | null
  readonly textRuns: readonly GridRenderTextRun[]
}): boolean {
  if ((input.previousDecorationCellKeys?.size ?? 0) > 0) {
    return true
  }
  return input.textRuns.some((run) => run.underline || run.strike)
}

export function shouldSkipTextResourceSyncForNonIntersectingDirtyCells(input: {
  readonly content: TextResourceSkipContent
  readonly missingGlyphRunSpans: readonly TextQuadRunSpan[]
  readonly nextTextRunCellKeys: readonly string[] | null
  readonly previousTextRunCellKeys: readonly string[] | null
  readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3
  readonly tile: GridRenderTile
}): boolean {
  if (
    input.missingGlyphRunSpans.length > 0 ||
    input.content.textRevisionKey === null ||
    input.content.textHandle === null ||
    !input.previousTextRunCellKeys ||
    !input.nextTextRunCellKeys
  ) {
    return false
  }
  if (!hasOnlyLocalTextCellDirtyMasks(input.tile)) {
    return false
  }
  return (
    !input.previousTextRunCellKeys.some((key) => isCellKeyDirtyForTile(key, input.tile)) &&
    !input.nextTextRunCellKeys.some((key) => isCellKeyDirtyForTile(key, input.tile)) &&
    areTextRunCellKeySequencesEqual(input.previousTextRunCellKeys, input.nextTextRunCellKeys) &&
    input.content.textRunCount === input.previousTextRunCellKeys.length &&
    input.textRevisionKey.textRunCount === input.nextTextRunCellKeys.length
  )
}

function areTextRunCellKeySequencesEqual(left: readonly string[], right: readonly string[]): boolean {
  if (left.length !== right.length) {
    return false
  }
  return left.every((key, index) => key === right[index])
}

function hasOnlyLocalTextCellDirtyMasks(tile: GridRenderTile): boolean {
  const dirtyMasks = tile.dirtyMasks
  const dirtyLocalRows = tile.dirtyLocalRows
  const dirtyLocalCols = tile.dirtyLocalCols
  if (!dirtyMasks || !dirtyLocalRows || !dirtyLocalCols || dirtyMasks.length === 0) {
    return false
  }
  if (dirtyLocalRows.length !== dirtyMasks.length * 2 || dirtyLocalCols.length !== dirtyMasks.length * 2) {
    return false
  }
  for (const mask of dirtyMasks) {
    if ((mask & TEXT_CELL_SKIP_FORBIDDEN_DIRTY_MASK) !== 0) {
      return false
    }
  }
  return true
}

export function formatDecorationCellKey(row: number, col: number): string {
  return `${row}:${col}`
}
