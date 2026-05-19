import { noteTypeGpuTextPayload } from '../grid-render-counters.js'
import {
  resolveTextQuadRunContentSignatureV3,
  resolveTextQuadRunSignatureV3,
  type TextQuadRunPayloadV3,
  type TextQuadRunSpan,
} from './line-text-quad-buffer.js'
import type { GpuBufferHandleV3 } from './gpu-buffer-arena.js'
import type { GridRenderTile, GridRenderTileTextRun } from './render-tile-source.js'
import { DirtyMaskV3 } from './tile-damage-index.js'
import type { TypeGpuTileTextRevisionKeyV3 } from './typegpu-tile-resource-revisions.js'
import { type TextInstanceVertexBuffer, writeTypeGpuVertexBufferSubrange } from './typegpu-primitives.js'

const TEXT_INSTANCE_FLOAT_COUNT = 16
type AxisOnlyTextGeometryRejectReason = 'invalid-state' | 'missing-glyph' | 'signature'

interface AxisOnlyTextContentEntryV3 {
  textRevisionKey: TypeGpuTileTextRevisionKeyV3 | null
  textRunCount: number
  textRunPayloads: readonly TextQuadRunPayloadV3[] | null
  textRunQuadSpans: readonly TextQuadRunSpan[] | null
}

export function syncAxisOnlyTileTextGeometryResource(input: {
  readonly content: AxisOnlyTextContentEntryV3
  readonly dirtyTextRunSpans: readonly TextQuadRunSpan[] | undefined
  readonly hasMissingGlyphDependencies?: boolean | undefined
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer> | null
  readonly label: string
  readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3
  readonly tile: GridRenderTile
}): boolean {
  const dirtySpanResolution = resolveAxisOnlyTextDirtySpans({
    contentRevisionKey: input.content.textRevisionKey,
    dirtyTextRunSpans: input.dirtyTextRunSpans,
    textRevisionKey: input.textRevisionKey,
    tile: input.tile,
  })
  const dirtySpans = dirtySpanResolution.spans
  const attemptedAxisOnlySync = dirtySpans.length > 0 && dirtySpanResolution.isAxisOnly
  if (dirtySpans.length === 0 || !dirtySpanResolution.isAxisOnly) {
    return false
  }
  if (input.hasMissingGlyphDependencies === true) {
    return rejectAxisOnlyTextGeometrySync('missing-glyph')
  }
  if (
    !input.handle ||
    input.content.textRevisionKey === null ||
    input.content.textRunPayloads === null ||
    input.content.textRunQuadSpans === null ||
    input.content.textRunCount !== input.tile.textCount ||
    input.content.textRunPayloads.length !== input.tile.textRuns.length ||
    input.content.textRunQuadSpans.length !== input.tile.textRuns.length
  ) {
    return attemptedAxisOnlySync ? rejectAxisOnlyTextGeometrySync('invalid-state') : false
  }

  const nextPayloads = [...input.content.textRunPayloads]
  let reusedRunPayloads = 0
  for (const dirtySpan of dirtySpans) {
    const start = dirtySpan.offset
    const end = start + dirtySpan.length
    if (start < 0 || dirtySpan.length <= 0 || end > input.tile.textRuns.length) {
      return false
    }
    let pendingWrite: PendingAxisOnlyTextWrite | null = null
    for (let runIndex = start; runIndex < end; runIndex += 1) {
      const run = input.tile.textRuns[runIndex]
      const previousPayload = input.content.textRunPayloads[runIndex]
      const previousRunSpan = input.content.textRunQuadSpans[runIndex]
      if (!run || !previousPayload || !previousRunSpan) {
        return rejectAxisOnlyTextGeometrySync('invalid-state')
      }
      const translatedPayload = translateAxisOnlyTextRunPayload(run, previousPayload)
      if (!translatedPayload || translatedPayload.quadCount !== previousRunSpan.length) {
        return rejectAxisOnlyTextGeometrySync('signature')
      }
      nextPayloads[runIndex] = translatedPayload
      reusedRunPayloads += 1
      pendingWrite = queueAxisOnlyTextGeometryWrite({
        handle: input.handle,
        label: input.label,
        pendingWrite,
        previousPayload,
        previousRunSpan,
        translatedPayload,
      })
    }
    flushAxisOnlyTextGeometryWrite(input.handle, input.label, pendingWrite)
  }

  input.content.textRunPayloads = nextPayloads
  input.content.textRevisionKey = input.textRevisionKey
  noteTypeGpuTextPayload({
    atlasGeometryResyncs: 0,
    atlasGeometryRetries: 0,
    glyphDependencies: 0,
    pageDependencies: 0,
    rebuiltRunPayloads: 0,
    reusedRunPayloads,
    axisOnlySyncAccepts: 1,
    axisOnlySyncAuthoritativeFullTile: dirtySpanResolution.authoritativeFullTile ? 1 : 0,
  })
  return true
}

interface PendingAxisOnlyTextWrite {
  readonly chunks: Float32Array[]
  readonly floatCount: number
  readonly quadEnd: number
  readonly quadOffset: number
}

export function shouldAttemptAxisOnlyTileTextGeometryResourceSync(input: {
  readonly contentRevisionKey: TypeGpuTileTextRevisionKeyV3 | null
  readonly dirtyTextRunSpans: readonly TextQuadRunSpan[] | undefined
  readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3
  readonly tile: GridRenderTile
}): boolean {
  const resolution = resolveAxisOnlyTextDirtySpans(input)
  return resolution.isAxisOnly && resolution.spans.length > 0
}

function resolveAxisOnlyTextDirtySpans(input: {
  readonly contentRevisionKey: TypeGpuTileTextRevisionKeyV3 | null
  readonly dirtyTextRunSpans: readonly TextQuadRunSpan[] | undefined
  readonly textRevisionKey: TypeGpuTileTextRevisionKeyV3
  readonly tile: GridRenderTile
}): {
  readonly authoritativeFullTile: boolean
  readonly isAxisOnly: boolean
  readonly spans: readonly TextQuadRunSpan[]
} {
  const dirtySpans = input.dirtyTextRunSpans ?? []
  if (dirtySpans.length > 0) {
    return {
      authoritativeFullTile: false,
      isAxisOnly: isAxisOnlyTextGeometryDirty(input.tile),
      spans: dirtySpans,
    }
  }
  if (!isAxisOnlyTextRevisionChange(input.contentRevisionKey, input.textRevisionKey)) {
    return {
      authoritativeFullTile: false,
      isAxisOnly: false,
      spans: [],
    }
  }
  return {
    authoritativeFullTile: true,
    isAxisOnly: true,
    spans: input.tile.textCount > 0 ? [{ offset: 0, length: input.tile.textCount }] : [],
  }
}

function isAxisOnlyTextRevisionChange(previous: TypeGpuTileTextRevisionKeyV3 | null, next: TypeGpuTileTextRevisionKeyV3): boolean {
  return (
    previous !== null &&
    previous.tileId === next.tileId &&
    previous.textRunCount === next.textRunCount &&
    previous.freezeSeq === next.freezeSeq &&
    (previous.axisSeqX !== next.axisSeqX || previous.axisSeqY !== next.axisSeqY)
  )
}

function isAxisOnlyTextGeometryDirty(tile: GridRenderTile): boolean {
  const dirtyMasks = tile.dirtyMasks
  if (!dirtyMasks || dirtyMasks.length === 0) {
    return false
  }
  const allowedMask = DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect
  const requiredMask = DirtyMaskV3.AxisX | DirtyMaskV3.AxisY
  for (const mask of dirtyMasks) {
    if ((mask & requiredMask) === 0 || (mask & ~allowedMask) !== 0) {
      return false
    }
  }
  return true
}

function rejectAxisOnlyTextGeometrySync(reason: AxisOnlyTextGeometryRejectReason): false {
  noteTypeGpuTextPayload({
    atlasGeometryResyncs: 0,
    atlasGeometryRetries: 0,
    glyphDependencies: 0,
    pageDependencies: 0,
    rebuiltRunPayloads: 0,
    reusedRunPayloads: 0,
    axisOnlySyncMissingGlyphRejects: reason === 'missing-glyph' ? 1 : 0,
    axisOnlySyncRejects: 1,
    axisOnlySyncSignatureRejects: reason === 'signature' ? 1 : 0,
  })
  return false
}

function queueAxisOnlyTextGeometryWrite(input: {
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
  readonly label: string
  readonly pendingWrite: PendingAxisOnlyTextWrite | null
  readonly previousPayload: TextQuadRunPayloadV3
  readonly previousRunSpan: TextQuadRunSpan
  readonly translatedPayload: TextQuadRunPayloadV3
}): PendingAxisOnlyTextWrite | null {
  const floatCount = input.translatedPayload.quadCount * TEXT_INSTANCE_FLOAT_COUNT
  if (floatCount === 0 || areTextRunPayloadFloatsEqual(input.previousPayload.floats, input.translatedPayload.floats, floatCount)) {
    return input.pendingWrite
  }

  const runQuadOffset = input.previousRunSpan.offset
  const runQuadEnd = runQuadOffset + input.previousRunSpan.length
  const runFloats = input.translatedPayload.floats.subarray(0, floatCount)
  if (!input.pendingWrite) {
    return {
      chunks: [runFloats],
      floatCount,
      quadEnd: runQuadEnd,
      quadOffset: runQuadOffset,
    }
  }
  if (input.pendingWrite.quadEnd === runQuadOffset) {
    input.pendingWrite.chunks.push(runFloats)
    return {
      chunks: input.pendingWrite.chunks,
      floatCount: input.pendingWrite.floatCount + floatCount,
      quadEnd: runQuadEnd,
      quadOffset: input.pendingWrite.quadOffset,
    }
  }

  flushAxisOnlyTextGeometryWrite(input.handle, input.label, input.pendingWrite)
  return {
    chunks: [runFloats],
    floatCount,
    quadEnd: runQuadEnd,
    quadOffset: runQuadOffset,
  }
}

function flushAxisOnlyTextGeometryWrite(
  handle: GpuBufferHandleV3<TextInstanceVertexBuffer>,
  label: string,
  pendingWrite: PendingAxisOnlyTextWrite | null,
): void {
  if (!pendingWrite || pendingWrite.floatCount === 0) {
    return
  }
  const floats = mergePendingAxisOnlyTextChunks(pendingWrite)
  writeTypeGpuVertexBufferSubrange({
    buffer: handle.buffer,
    floatCount: pendingWrite.floatCount,
    floats,
    label: `${label}:axis-span`,
    sourceStartFloat: 0,
    startFloat: pendingWrite.quadOffset * TEXT_INSTANCE_FLOAT_COUNT,
  })
}

function mergePendingAxisOnlyTextChunks(pendingWrite: PendingAxisOnlyTextWrite): Float32Array {
  if (pendingWrite.chunks.length === 1 && pendingWrite.chunks[0]?.length === pendingWrite.floatCount) {
    return pendingWrite.chunks[0]
  }
  const floats = new Float32Array(pendingWrite.floatCount)
  let offset = 0
  for (const chunk of pendingWrite.chunks) {
    floats.set(chunk, offset)
    offset += chunk.length
  }
  return floats
}

function areTextRunPayloadFloatsEqual(left: Float32Array, right: Float32Array, floatCount: number): boolean {
  for (let index = 0; index < floatCount; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

function translateAxisOnlyTextRunPayload(run: GridRenderTileTextRun, previousPayload: TextQuadRunPayloadV3): TextQuadRunPayloadV3 | null {
  if (run.wrap === true) {
    return null
  }
  const contentSignature = resolveTextQuadRunContentSignatureV3(run)
  if (!canReuseAxisOnlyTextRunPayload(run, previousPayload, contentSignature)) {
    return null
  }
  if (previousPayload.quadCount <= 0) {
    return {
      ...previousPayload,
      contentSignature,
      clipHeight: run.clipHeight,
      clipWidth: run.clipWidth,
      clipX: run.clipX,
      clipY: run.clipY,
      height: run.height,
      signature: resolveTextQuadRunSignatureV3(run),
      width: run.width,
      x: run.x,
      y: run.y,
    }
  }
  const dx = run.x - previousPayload.x
  const dy = run.y - previousPayload.y
  const dTextX = dx + resolveHorizontalTextGeometryDelta(run, previousPayload)
  const dTextY = dy + resolveVerticalTextGeometryDelta(run, previousPayload)
  const dClipX = run.clipX - (previousPayload.clipX ?? previousPayload.x)
  const dClipY = run.clipY - (previousPayload.clipY ?? previousPayload.y)
  const previousClipRight = (previousPayload.clipX ?? previousPayload.x) + (previousPayload.clipWidth ?? previousPayload.width ?? 0)
  const previousClipBottom = (previousPayload.clipY ?? previousPayload.y) + (previousPayload.clipHeight ?? previousPayload.height ?? 0)
  const dClipRight = run.clipX + run.clipWidth - previousClipRight
  const dClipBottom = run.clipY + run.clipHeight - previousClipBottom
  const floats = previousPayload.floats.slice()
  const floatCount = previousPayload.quadCount * TEXT_INSTANCE_FLOAT_COUNT
  for (let offset = 0; offset < floatCount; offset += TEXT_INSTANCE_FLOAT_COUNT) {
    floats[offset + 0] = (floats[offset + 0] ?? 0) + dTextX
    floats[offset + 1] = (floats[offset + 1] ?? 0) + dTextY
    floats[offset + 12] = (floats[offset + 12] ?? 0) + dClipX
    floats[offset + 13] = (floats[offset + 13] ?? 0) + dClipY
    floats[offset + 14] = (floats[offset + 14] ?? 0) + dClipRight
    floats[offset + 15] = (floats[offset + 15] ?? 0) + dClipBottom
  }
  return {
    ...previousPayload,
    clipHeight: run.clipHeight,
    clipWidth: run.clipWidth,
    clipX: run.clipX,
    clipY: run.clipY,
    contentSignature,
    floats,
    height: run.height,
    signature: resolveTextQuadRunSignatureV3(run),
    width: run.width,
    x: run.x,
    y: run.y,
  }
}

function canReuseAxisOnlyTextRunPayload(
  run: GridRenderTileTextRun,
  previousPayload: TextQuadRunPayloadV3,
  contentSignature: string,
): boolean {
  if (previousPayload.contentSignature === contentSignature) {
    return true
  }
  if (run.wrap === true) {
    return false
  }
  return (
    previousPayload.contentSignature ===
    resolveTextQuadRunContentSignatureV3({
      ...run,
      ...(previousPayload.clipHeight === undefined ? {} : { clipHeight: previousPayload.clipHeight }),
      ...(previousPayload.clipWidth === undefined ? {} : { clipWidth: previousPayload.clipWidth }),
      ...(previousPayload.height === undefined ? {} : { height: previousPayload.height }),
      ...(previousPayload.width === undefined ? {} : { width: previousPayload.width }),
    })
  )
}

function resolveHorizontalTextGeometryDelta(run: GridRenderTileTextRun, previousPayload: TextQuadRunPayloadV3): number {
  const widthDelta = (run.clipWidth ?? run.width) - (previousPayload.clipWidth ?? previousPayload.width ?? run.clipWidth ?? run.width)
  if (run.align === 'right') {
    return widthDelta
  }
  if (run.align === 'center') {
    return widthDelta / 2
  }
  return 0
}

function resolveVerticalTextGeometryDelta(run: GridRenderTileTextRun, previousPayload: TextQuadRunPayloadV3): number {
  const heightDelta =
    (run.clipHeight ?? run.height) - (previousPayload.clipHeight ?? previousPayload.height ?? run.clipHeight ?? run.height)
  return heightDelta / 2
}
