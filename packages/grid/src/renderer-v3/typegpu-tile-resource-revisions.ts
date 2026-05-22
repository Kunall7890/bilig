import type { TextDecorationRect, TextQuadRunSpan } from './line-text-quad-buffer.js'
import type { GridRenderTile } from './render-tile-source.js'
import { DirtyMaskV3 } from './tile-damage-index.js'
import type { createGlyphAtlas } from './typegpu-atlas-manager.js'
import type { TextInstanceVertexBuffer } from './typegpu-primitives.js'
import type { GpuBufferHandleV3 } from './gpu-buffer-arena.js'

const RECT_DIRTY_MASK_V3 =
  DirtyMaskV3.Style | DirtyMaskV3.Rect | DirtyMaskV3.Border | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
const TEXT_DIRTY_MASK_V3 =
  DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
const TEXT_DECORATION_DIRTY_MASK_V3 = DirtyMaskV3.Value | DirtyMaskV3.Text

export interface TypeGpuTileTextRevisionKeyV3 {
  readonly tileId: number
  readonly textRunCount: number
  readonly textSignature: string
  readonly valueSeq: number
  readonly styleSeq: number
  readonly textSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly batchSeq: number
}

export interface TypeGpuTileRectRevisionKeyV3 {
  readonly tileId: number
  readonly rectCount: number
  readonly rectSignature: string
  readonly decorationRectSignature: string
  readonly valueSeq: number
  readonly styleSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly batchSeq: number
  readonly decorationRectCount: number
}

export function resolveGridTextTileRevisionKeyV3(tile: GridRenderTile): TypeGpuTileTextRevisionKeyV3 {
  return {
    axisSeqX: tile.version.axisX,
    axisSeqY: tile.version.axisY,
    batchSeq: tile.lastBatchId,
    freezeSeq: tile.version.freeze,
    styleSeq: tile.version.styles,
    textRunCount: tile.textCount,
    textSignature: resolveGridTextRunSignatureV3(tile),
    textSeq: tile.version.text,
    tileId: tile.tileId,
    valueSeq: tile.version.values,
  }
}

export function resolveGridRectTileRevisionKeyV3(input: {
  readonly tile: GridRenderTile
  readonly decorationRects?: readonly TextDecorationRect[] | undefined
}): TypeGpuTileRectRevisionKeyV3 {
  const decorationRects = input.decorationRects ?? []
  return {
    axisSeqX: input.tile.version.axisX,
    axisSeqY: input.tile.version.axisY,
    batchSeq: input.tile.lastBatchId,
    decorationRectCount: decorationRects.length,
    decorationRectSignature: resolveGridTextDecorationRectSignatureV3(decorationRects),
    freezeSeq: input.tile.version.freeze,
    rectCount: input.tile.rectCount,
    rectSignature: input.tile.rectSignature ?? '',
    styleSeq: input.tile.version.styles,
    tileId: input.tile.tileId,
    valueSeq: input.tile.version.values,
  }
}

export function areGridTextTileRevisionKeysEqualV3(
  left: TypeGpuTileTextRevisionKeyV3 | null | undefined,
  right: TypeGpuTileTextRevisionKeyV3 | null | undefined,
): boolean {
  return (
    left !== null &&
    left !== undefined &&
    right !== null &&
    right !== undefined &&
    left.tileId === right.tileId &&
    left.textRunCount === right.textRunCount &&
    left.textSignature === right.textSignature &&
    left.valueSeq === right.valueSeq &&
    left.styleSeq === right.styleSeq &&
    left.textSeq === right.textSeq &&
    left.axisSeqX === right.axisSeqX &&
    left.axisSeqY === right.axisSeqY &&
    left.freezeSeq === right.freezeSeq &&
    left.batchSeq === right.batchSeq
  )
}

function resolveGridTextRunSignatureV3(tile: Pick<GridRenderTile, 'textRuns' | 'textSignature'>): string {
  if (tile.textSignature) {
    return tile.textSignature
  }
  let hash = 2_166_136_261
  hash = mixRevisionNumber(hash, tile.textRuns.length)
  for (const run of tile.textRuns) {
    hash = mixRevisionString(hash, run.text)
    hash = mixRevisionNumber(hash, run.x)
    hash = mixRevisionNumber(hash, run.y)
    hash = mixRevisionNumber(hash, run.width)
    hash = mixRevisionNumber(hash, run.height)
    hash = mixRevisionNumber(hash, run.clipX)
    hash = mixRevisionNumber(hash, run.clipY)
    hash = mixRevisionNumber(hash, run.clipWidth)
    hash = mixRevisionNumber(hash, run.clipHeight)
    hash = mixRevisionString(hash, run.align ?? 'left')
    hash = mixRevisionNumber(hash, run.wrap ? 1 : 0)
    hash = mixRevisionString(hash, run.font)
    hash = mixRevisionNumber(hash, run.fontSize)
    hash = mixRevisionString(hash, run.color)
    hash = mixRevisionNumber(hash, run.underline ? 1 : 0)
    hash = mixRevisionNumber(hash, run.strike ? 1 : 0)
  }
  return hash.toString(36)
}

function resolveGridTextDecorationRectSignatureV3(decorationRects: readonly TextDecorationRect[]): string {
  let hash = 2_166_136_261
  hash = mixRevisionNumber(hash, decorationRects.length)
  for (const rect of decorationRects) {
    hash = mixRevisionNumber(hash, rect.x)
    hash = mixRevisionNumber(hash, rect.y)
    hash = mixRevisionNumber(hash, rect.width)
    hash = mixRevisionNumber(hash, rect.height)
    hash = mixRevisionString(hash, rect.color)
  }
  return hash.toString(36)
}

function mixRevisionString(hash: number, value: string): number {
  let next = hash
  for (let index = 0; index < value.length; index += 1) {
    next = mixRevisionInteger(next, value.charCodeAt(index))
  }
  return next
}

function mixRevisionNumber(hash: number, value: number): number {
  return mixRevisionInteger(hash, Math.round(value * 1_000))
}

function mixRevisionInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}

export function areGridRectTileRevisionKeysEqualV3(
  left: TypeGpuTileRectRevisionKeyV3 | null | undefined,
  right: TypeGpuTileRectRevisionKeyV3 | null | undefined,
): boolean {
  return (
    left !== null &&
    left !== undefined &&
    right !== null &&
    right !== undefined &&
    left.tileId === right.tileId &&
    left.rectCount === right.rectCount &&
    left.rectSignature === right.rectSignature &&
    left.decorationRectSignature === right.decorationRectSignature &&
    left.valueSeq === right.valueSeq &&
    left.styleSeq === right.styleSeq &&
    left.axisSeqX === right.axisSeqX &&
    left.axisSeqY === right.axisSeqY &&
    left.freezeSeq === right.freezeSeq &&
    left.batchSeq === right.batchSeq &&
    left.decorationRectCount === right.decorationRectCount
  )
}

function resolveGridTileDirtyContentMaskV3(tile: Pick<GridRenderTile, 'dirtyMasks'>): number | null {
  const masks = tile.dirtyMasks
  if (!masks || masks.length === 0) {
    return null
  }
  let mask = 0
  for (const value of masks) {
    mask |= value
  }
  return mask
}

export function shouldSyncGridTextTileResourceV3(input: {
  readonly atlasGeometryVersion?: number | undefined
  readonly content: {
    readonly textAtlasGeometryVersion: number
    readonly textCount: number
    readonly textHandle: GpuBufferHandleV3<TextInstanceVertexBuffer> | null
    readonly textRunCount: number
    readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3 | null
  }
  readonly missingGlyphDependencies?: boolean | undefined
  readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3
  readonly tile: GridRenderTile
}): boolean {
  if (input.missingGlyphDependencies) {
    return true
  }
  const atlasGeometryChanged =
    input.atlasGeometryVersion !== undefined && input.content.textAtlasGeometryVersion !== input.atlasGeometryVersion
  if (areGridTextTileRevisionKeysEqualV3(input.content.textRevisionKey, input.textRevisionKey)) {
    return atlasGeometryChanged
  }
  if (!input.content.textRevisionKey) {
    return true
  }
  if (atlasGeometryChanged) {
    return true
  }
  if (hasGridTextTileResourcePayloadChangedV3(input.content.textRevisionKey, input.textRevisionKey)) {
    return true
  }
  if (input.content.textRunCount !== input.tile.textCount) {
    return true
  }
  if (input.tile.textCount > 0 && !input.content.textHandle) {
    return true
  }
  const dirtyMask = resolveGridTileDirtyContentMaskV3(input.tile)
  if (dirtyMask === null && areGridTextTilePayloadsEqualV3(input.content.textRevisionKey, input.textRevisionKey)) {
    return input.atlasGeometryVersion !== undefined && input.content.textAtlasGeometryVersion !== input.atlasGeometryVersion
  }
  return dirtyMask === null || (dirtyMask & TEXT_DIRTY_MASK_V3) !== 0
}

function areGridTextTilePayloadsEqualV3(
  left: TypeGpuTileTextRevisionKeyV3 | null | undefined,
  right: TypeGpuTileTextRevisionKeyV3 | null | undefined,
): boolean {
  return (
    left !== null &&
    left !== undefined &&
    right !== null &&
    right !== undefined &&
    left.tileId === right.tileId &&
    left.textRunCount === right.textRunCount &&
    left.textSignature === right.textSignature
  )
}

export function resolveMissingTextGlyphRunSpansV3(input: {
  readonly atlas: Pick<ReturnType<typeof createGlyphAtlas>, 'resolveGlyphRecord'>
  readonly content: {
    readonly textGlyphIds: readonly number[] | null
    readonly textGlyphPageIds: readonly number[] | null
    readonly textRunCount: number
    readonly textRunGlyphIds: readonly (readonly number[])[] | null
  }
}): readonly TextQuadRunSpan[] {
  const runCount = input.content.textRunCount
  if (runCount <= 0) {
    return []
  }

  const glyphIds = input.content.textGlyphIds
  const pageIds = input.content.textGlyphPageIds
  const runGlyphIds = input.content.textRunGlyphIds
  if (!glyphIds || glyphIds.length === 0) {
    return []
  }
  if (!pageIds || pageIds.length !== glyphIds.length || !runGlyphIds || runGlyphIds.length !== runCount) {
    return [{ length: runCount, offset: 0 }]
  }

  const expectedPageByGlyph = new Map<number, number>()
  for (let index = 0; index < glyphIds.length; index += 1) {
    const glyphId = glyphIds[index]
    const pageId = pageIds[index]
    if (glyphId === undefined || pageId === undefined) {
      return [{ length: runCount, offset: 0 }]
    }
    expectedPageByGlyph.set(glyphId, pageId)
  }

  const missing: TextQuadRunSpan[] = []
  for (let runIndex = 0; runIndex < runCount; runIndex += 1) {
    const runGlyphs = runGlyphIds[runIndex] ?? []
    let runMissing = false
    for (const glyphId of runGlyphs) {
      const expectedPageId = expectedPageByGlyph.get(glyphId)
      const glyphRecord = input.atlas.resolveGlyphRecord(glyphId)
      if (expectedPageId === undefined || !glyphRecord || glyphRecord.pageId !== expectedPageId) {
        runMissing = true
        break
      }
    }
    if (runMissing) {
      missing.push({ length: 1, offset: runIndex })
    }
  }
  return missing
}

export function shouldSyncGridRectTileResourceV3(input: {
  readonly content: {
    readonly decorationCellKeys: ReadonlySet<string> | null
    readonly decorationRects: readonly TextDecorationRect[] | null
    readonly rectCount: number
    readonly rectHandle: GpuBufferHandleV3 | null
    readonly rectRevisionKey: TypeGpuTileRectRevisionKeyV3 | null
  }
  readonly rectRevisionKey: TypeGpuTileRectRevisionKeyV3
  readonly tile: GridRenderTile
}): boolean {
  if (areGridRectTileRevisionKeysEqualV3(input.content.rectRevisionKey, input.rectRevisionKey)) {
    return false
  }
  if (!input.content.rectRevisionKey) {
    return true
  }
  if (hasGridRectTileResourcePayloadChangedV3(input.content.rectRevisionKey, input.rectRevisionKey)) {
    return true
  }
  if (input.content.rectCount !== input.tile.rectCount) {
    return true
  }
  if (input.tile.rectCount > 0 && !input.content.rectHandle) {
    return true
  }
  const dirtyMask = resolveGridTileDirtyContentMaskV3(input.tile)
  if (dirtyMask === null) {
    return false
  }
  if ((dirtyMask & RECT_DIRTY_MASK_V3) !== 0) {
    return true
  }
  if ((dirtyMask & TEXT_DECORATION_DIRTY_MASK_V3) === 0) {
    return false
  }
  return hasDirtyTextDecorationResourceV3({
    previousDecorationCellKeys: input.content.decorationCellKeys,
    tile: input.tile,
  })
}

function hasGridTextTileResourcePayloadChangedV3(previous: TypeGpuTileTextRevisionKeyV3, next: TypeGpuTileTextRevisionKeyV3): boolean {
  return previous.textRunCount !== next.textRunCount || previous.textSignature !== next.textSignature
}

function hasGridRectTileResourcePayloadChangedV3(previous: TypeGpuTileRectRevisionKeyV3, next: TypeGpuTileRectRevisionKeyV3): boolean {
  return (
    previous.rectCount !== next.rectCount ||
    previous.rectSignature !== next.rectSignature ||
    previous.decorationRectCount !== next.decorationRectCount ||
    previous.decorationRectSignature !== next.decorationRectSignature
  )
}

function hasDirtyTextDecorationResourceV3(input: {
  readonly previousDecorationCellKeys: ReadonlySet<string> | null
  readonly tile: GridRenderTile
}): boolean {
  const dirtyMasks = input.tile.dirtyMasks
  const dirtyLocalRows = input.tile.dirtyLocalRows
  const dirtyLocalCols = input.tile.dirtyLocalCols
  if (!dirtyMasks || !dirtyLocalRows || !dirtyLocalCols || dirtyMasks.length === 0) {
    return input.tile.textRuns.some((run) => run.underline || run.strike) || (input.previousDecorationCellKeys?.size ?? 0) > 0
  }

  for (const run of input.tile.textRuns) {
    if (!run.underline && !run.strike) {
      continue
    }
    if (run.row === undefined || run.col === undefined) {
      return true
    }
    if (
      isCellInDirtyTextDecorationRange(
        run.row,
        run.col,
        dirtyMasks,
        dirtyLocalRows,
        dirtyLocalCols,
        input.tile.bounds.rowStart,
        input.tile.bounds.colStart,
      )
    ) {
      return true
    }
  }

  const previousDecorationCellKeys = input.previousDecorationCellKeys
  if (!previousDecorationCellKeys || previousDecorationCellKeys.size === 0) {
    return false
  }
  if (previousDecorationCellKeys.has('*')) {
    return true
  }
  for (const key of previousDecorationCellKeys) {
    const cell = parseDecorationCellKey(key)
    if (
      cell &&
      isCellInDirtyTextDecorationRange(
        cell.row,
        cell.col,
        dirtyMasks,
        dirtyLocalRows,
        dirtyLocalCols,
        input.tile.bounds.rowStart,
        input.tile.bounds.colStart,
      )
    ) {
      return true
    }
  }
  return false
}

function isCellInDirtyTextDecorationRange(
  row: number,
  col: number,
  dirtyMasks: Uint32Array,
  dirtyLocalRows: Uint32Array,
  dirtyLocalCols: Uint32Array,
  tileRowStart: number,
  tileColStart: number,
): boolean {
  const localRow = row - tileRowStart
  const localCol = col - tileColStart
  for (let index = 0; index < dirtyMasks.length; index += 1) {
    const mask = dirtyMasks[index] ?? 0
    if ((mask & TEXT_DECORATION_DIRTY_MASK_V3) === 0) {
      continue
    }
    const offset = index * 2
    const rowStart = dirtyLocalRows[offset] ?? 0
    const rowEnd = dirtyLocalRows[offset + 1] ?? rowStart
    const colStart = dirtyLocalCols[offset] ?? 0
    const colEnd = dirtyLocalCols[offset + 1] ?? colStart
    if (localRow >= rowStart && localRow <= rowEnd && localCol >= colStart && localCol <= colEnd) {
      return true
    }
  }
  return false
}

function parseDecorationCellKey(key: string): { readonly row: number; readonly col: number } | null {
  const separator = key.indexOf(':')
  if (separator <= 0 || separator === key.length - 1) {
    return null
  }
  const row = Number(key.slice(0, separator))
  const col = Number(key.slice(separator + 1))
  return Number.isInteger(row) && row >= 0 && Number.isInteger(col) && col >= 0 ? { col, row } : null
}
