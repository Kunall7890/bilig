import type { GpuBufferHandleV3 } from './gpu-buffer-arena.js'
import { buildTextQuadsFromRunsWithSpans, type TextQuadRunPayloadV3, type TextQuadRunSpan } from './line-text-quad-buffer.js'
import { isFullGridRenderTileDirtySpanV3 } from './render-tile-dirty-spans.js'
import type { GridRenderTile } from './render-tile-source.js'
import { DirtyMaskV3 } from './tile-damage-index.js'
import type { createGlyphAtlas } from './typegpu-atlas-manager.js'
import type { TypeGpuTileContentResourceEntryV3 } from './typegpu-tile-buffer-pool.js'
import { type TextInstanceVertexBuffer, writeTypeGpuVertexBuffer, writeTypeGpuVertexBufferSubrange } from './typegpu-primitives.js'

const TEXT_INSTANCE_FLOAT_COUNT = 16
const TEXT_INSTANCE_BYTE_COUNT = TEXT_INSTANCE_FLOAT_COUNT * Float32Array.BYTES_PER_ELEMENT
const MAX_VARIABLE_TEXT_PATCH_RUNS = 128

type TextTilePayloadV3 = ReturnType<typeof buildTextQuadsFromRunsWithSpans>

export function buildSmallDirtyTextPatchPayload(input: {
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly dirtyRunSpans: readonly TextQuadRunSpan[] | undefined
  readonly forceRebuildDirtyRunSpans: boolean
  readonly previousGlyphIds: readonly number[] | null
  readonly previousPageIds: readonly number[] | null
  readonly previousRunGlyphIds: readonly (readonly number[])[] | null
  readonly previousRunPayloads: readonly TextQuadRunPayloadV3[] | null
  readonly previousRunSpans: readonly TextQuadRunSpan[] | null
  readonly previousTextCount: number
  readonly runs: readonly GridRenderTile['textRuns'][number][]
}): TextTilePayloadV3 | null {
  const dirtyRunSpans = input.dirtyRunSpans ?? []
  if (
    dirtyRunSpans.length === 0 ||
    dirtyRunSpans.some((span) => isFullGridRenderTileDirtySpanV3(span, input.runs.length)) ||
    !input.previousRunPayloads ||
    !input.previousRunSpans ||
    input.previousRunPayloads.length !== input.runs.length ||
    input.previousRunSpans.length !== input.runs.length
  ) {
    return null
  }
  const dirtyRunCount = countDirtyTextRuns(dirtyRunSpans, input.runs.length)
  if (dirtyRunCount <= 0 || dirtyRunCount > MAX_VARIABLE_TEXT_PATCH_RUNS) {
    return null
  }

  const runPayloads = input.previousRunPayloads.slice()
  const runGlyphIds =
    input.previousRunGlyphIds && input.previousRunGlyphIds.length === input.runs.length
      ? input.previousRunGlyphIds.slice()
      : input.previousRunPayloads.map((payload) => payload.glyphIds)
  const runSpans = input.previousRunSpans.slice()
  const glyphIds: number[] = input.previousGlyphIds ? [...input.previousGlyphIds] : []
  const pageIds: number[] = input.previousPageIds ? [...input.previousPageIds] : []
  const glyphDependencyPages = new Map<number, number>()
  for (let index = 0; index < glyphIds.length; index += 1) {
    const glyphId = glyphIds[index]
    const pageId = pageIds[index]
    if (glyphId !== undefined && pageId !== undefined) {
      glyphDependencyPages.set(glyphId, pageId)
    }
  }
  let quadCount = input.previousTextCount
  let rebuiltRunPayloads = 0
  let atlasGeometryRetries = 0
  const visited = new Set<number>()

  for (const span of dirtyRunSpans) {
    const start = Math.max(0, Math.min(input.runs.length, span.offset))
    const end = Math.max(start, Math.min(input.runs.length, span.offset + span.length))
    for (let index = start; index < end; index += 1) {
      if (visited.has(index)) {
        continue
      }
      visited.add(index)
      const run = input.runs[index]
      const previousPayload = input.previousRunPayloads[index]
      const previousSpan = input.previousRunSpans[index]
      if (!run || !previousPayload || !previousSpan) {
        return null
      }
      const rebuilt = buildTextQuadsFromRunsWithSpans([run], input.atlas, undefined, {
        dirtyRunSpans: [{ length: 1, offset: 0 }],
        forceRebuildDirtyRunSpans: input.forceRebuildDirtyRunSpans,
        packFullPayload: true,
        previousRunPayloads: [previousPayload],
      })
      const nextPayload = rebuilt.runPayloads[0]
      if (!nextPayload) {
        return null
      }
      runPayloads[index] = nextPayload
      runGlyphIds[index] = rebuilt.runGlyphIds[0] ?? nextPayload.glyphIds
      runSpans[index] = { offset: previousSpan.offset, length: nextPayload.quadCount }
      quadCount += nextPayload.quadCount - previousPayload.quadCount
      appendUniqueTextGlyphDependencies(glyphIds, pageIds, glyphDependencyPages, nextPayload.glyphIds, nextPayload.pageIds)
      rebuiltRunPayloads += rebuilt.diagnostics.rebuiltRunPayloads
      atlasGeometryRetries += rebuilt.diagnostics.atlasGeometryRetries
    }
  }

  return {
    floats: new Float32Array(0),
    glyphIds,
    pageIds,
    quadCount,
    runGlyphIds,
    runPayloads,
    runSpans,
    diagnostics: {
      atlasGeometryRetries,
      rebuiltRunPayloads,
      reusedRunPayloads: 0,
    },
  }
}

function appendUniqueTextGlyphDependencies(
  glyphIds: number[],
  pageIds: number[],
  glyphDependencyPages: Map<number, number>,
  nextGlyphIds: readonly number[],
  nextPageIds: readonly number[],
): void {
  for (let index = 0; index < nextGlyphIds.length; index += 1) {
    const glyphId = nextGlyphIds[index]
    const pageId = nextPageIds[index]
    if (glyphId === undefined || pageId === undefined || glyphDependencyPages.has(glyphId)) {
      continue
    }
    glyphDependencyPages.set(glyphId, pageId)
    glyphIds.push(glyphId)
    pageIds.push(pageId)
  }
}

interface TextPatchDirtySpanResolution {
  readonly authoritative: boolean
  readonly spans: readonly TextQuadRunSpan[] | undefined
}

export function resolveTextPatchDirtySpans(input: {
  readonly dirtySpans: readonly TextQuadRunSpan[] | undefined
  readonly previousRunPayloads: readonly TextQuadRunPayloadV3[] | null
  readonly textPayload: { readonly runPayloads: readonly TextQuadRunPayloadV3[] }
  readonly tileTextCount: number
}): TextPatchDirtySpanResolution {
  const dirtySpans = input.dirtySpans ?? []
  if (!shouldDeriveTextPatchDirtySpans(dirtySpans, input.tileTextCount) || input.previousRunPayloads === null) {
    return { authoritative: false, spans: input.dirtySpans }
  }
  const changedSpans = resolveChangedTextRunPayloadSpans({
    nextRunPayloads: input.textPayload.runPayloads,
    previousRunPayloads: input.previousRunPayloads,
    textRunCount: input.tileTextCount,
  })
  return changedSpans === null ? { authoritative: false, spans: input.dirtySpans } : { authoritative: true, spans: changedSpans }
}

function shouldDeriveTextPatchDirtySpans(dirtySpans: readonly TextQuadRunSpan[], textRunCount: number): boolean {
  return dirtySpans.length === 0 || dirtySpans.some((span) => isFullGridRenderTileDirtySpanV3(span, textRunCount))
}

function resolveChangedTextRunPayloadSpans(input: {
  readonly nextRunPayloads: readonly TextQuadRunPayloadV3[]
  readonly previousRunPayloads: readonly TextQuadRunPayloadV3[]
  readonly textRunCount: number
}): readonly TextQuadRunSpan[] | null {
  if (
    input.textRunCount <= 0 ||
    input.previousRunPayloads.length !== input.textRunCount ||
    input.nextRunPayloads.length !== input.textRunCount
  ) {
    return null
  }
  const spans: TextQuadRunSpan[] = []
  let spanStart = -1
  let spanLength = 0
  for (let index = 0; index < input.textRunCount; index += 1) {
    const previousPayload = input.previousRunPayloads[index]
    const nextPayload = input.nextRunPayloads[index]
    const changed = !previousPayload || !nextPayload || !areTextRunPayloadsVertexEquivalent(previousPayload, nextPayload)
    if (changed) {
      if (spanStart < 0) {
        spanStart = index
        spanLength = 1
      } else {
        spanLength += 1
      }
      continue
    }
    if (spanStart >= 0) {
      spans.push({ offset: spanStart, length: spanLength })
      spanStart = -1
      spanLength = 0
    }
  }
  if (spanStart >= 0) {
    spans.push({ offset: spanStart, length: spanLength })
  }
  return spans
}

function areTextRunPayloadsVertexEquivalent(left: TextQuadRunPayloadV3, right: TextQuadRunPayloadV3): boolean {
  if (left.quadCount !== right.quadCount) {
    return false
  }
  const floatCount = left.quadCount * TEXT_INSTANCE_FLOAT_COUNT
  if (left.floats.length < floatCount || right.floats.length < floatCount) {
    return false
  }
  for (let offset = 0; offset < floatCount; offset += 1) {
    if (left.floats[offset] !== right.floats[offset]) {
      return false
    }
  }
  return true
}

export function writeTileTextPayload(input: {
  readonly canWritePartialPayload: boolean
  readonly content: TypeGpuTileContentResourceEntryV3
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
  readonly label: string
  readonly nextTextRunCellKeys: readonly string[] | null
  readonly previousRunSpans: readonly TextQuadRunSpan[] | null
  readonly previousTextRunCellKeys: readonly string[] | null
  readonly previousTextCount: number
  readonly textPayload: {
    readonly floats: Float32Array
    readonly quadCount: number
    readonly runPayloads: readonly TextQuadRunPayloadV3[]
    readonly runSpans: readonly TextQuadRunSpan[]
  }
  readonly tile: GridRenderTile
  readonly textDirtySpans?: readonly TextQuadRunSpan[] | undefined
  readonly textDirtySpansAuthoritative?: boolean | undefined
}): { readonly runSpans: readonly TextQuadRunSpan[]; readonly textCount: number } | null {
  const dirtySpans = input.textDirtySpans ?? input.tile.dirty?.textSpans ?? []
  if (!input.canWritePartialPayload) {
    return writeFullTileTextPayload(input, 'full-no-partial')
      ? { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
      : null
  }
  if (dirtySpans.length === 0 && input.previousTextCount === input.textPayload.quadCount) {
    if (input.textDirtySpansAuthoritative === true) {
      return { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
    }
    return writeFullTileTextPayload(input, 'full-clean')
      ? { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
      : null
  }
  if (dirtySpans.some((span) => isFullGridRenderTileDirtySpanV3(span, input.tile.textCount))) {
    const keyedPatchResult = writeKeyedDirtyTextPayload({
      dirtySpans,
      handle: input.handle,
      label: input.label,
      nextTextRunCellKeys: input.nextTextRunCellKeys,
      previousRunSpans: input.previousRunSpans,
      previousTextCount: input.previousTextCount,
      previousTextRunCellKeys: input.previousTextRunCellKeys,
      textPayload: input.textPayload,
      tile: input.tile,
    })
    if (keyedPatchResult) {
      return keyedPatchResult
    }
    return writeFullTileTextPayload(input, 'full-dirty-span')
      ? { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
      : null
  }

  const previousRunSpans = input.previousRunSpans
  if (!previousRunSpans || previousRunSpans.length !== input.textPayload.runSpans.length) {
    const keyedPatchResult = writeKeyedDirtyTextPayload({
      dirtySpans,
      handle: input.handle,
      label: input.label,
      nextTextRunCellKeys: input.nextTextRunCellKeys,
      previousRunSpans,
      previousTextCount: input.previousTextCount,
      previousTextRunCellKeys: input.previousTextRunCellKeys,
      textPayload: input.textPayload,
      tile: input.tile,
    })
    if (keyedPatchResult) {
      return keyedPatchResult
    }
    return writeFullTileTextPayload(input, 'full-keyed-fallback')
      ? { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
      : null
  }

  if (input.previousTextCount !== input.textPayload.quadCount) {
    const variablePatchResult = writeVariableLengthDirtyTextPayload({
      dirtySpans,
      handle: input.handle,
      label: input.label,
      previousRunSpans,
      previousTextCount: input.previousTextCount,
      textPayload: input.textPayload,
      tileTextCount: input.tile.textCount,
    })
    if (variablePatchResult) {
      return variablePatchResult
    }
    return writeFullTileTextPayload(input, 'full-variable-fallback')
      ? { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
      : null
  }

  for (const dirtySpan of dirtySpans) {
    const quadSpan = resolveStableTextQuadSpan({
      dirtySpan,
      nextRunSpans: input.textPayload.runSpans,
      previousRunSpans,
    })
    if (!quadSpan || quadSpan.length === 0) {
      return writeFullTileTextPayload(input, 'full-stable-fallback')
        ? { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
        : null
    }
    writeTypeGpuVertexBufferSubrange({
      buffer: input.handle.buffer,
      floatCount: quadSpan.length * TEXT_INSTANCE_FLOAT_COUNT,
      floats: input.textPayload.floats,
      label: `${input.label}:span`,
      startFloat: quadSpan.offset * TEXT_INSTANCE_FLOAT_COUNT,
    })
  }
  return { runSpans: input.textPayload.runSpans, textCount: input.textPayload.quadCount }
}

function writeFullTileTextPayload(
  input: {
    readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
    readonly label: string
    readonly textPayload: { readonly floats: Float32Array; readonly quadCount: number }
  },
  reason: string,
): boolean {
  if (input.textPayload.floats.length < input.textPayload.quadCount * TEXT_INSTANCE_FLOAT_COUNT) {
    return false
  }
  writeTypeGpuVertexBuffer(input.handle.buffer, input.textPayload.floats, `${input.label}:${reason}`)
  return true
}

interface VariableLengthTextPatchPlan {
  readonly runSpans: readonly TextQuadRunSpan[]
  readonly textCount: number
  readonly writes: readonly VariableLengthTextPatchWrite[]
}

interface VariableLengthTextPatchWrite {
  readonly floatCount: number
  readonly floats: Float32Array
  readonly startFloat: number
  readonly suffix: string
}

export function resolveVariableLengthTextPatchReserveCount(input: {
  readonly dirtySpans: readonly TextQuadRunSpan[] | undefined
  readonly previousTextCount: number
  readonly textPayload: { readonly runPayloads?: readonly TextQuadRunPayloadV3[] | undefined; readonly quadCount: number }
  readonly tile: GridRenderTile
}): number {
  const dirtySpans = input.dirtySpans ?? input.tile.dirty?.textSpans ?? []
  if (dirtySpans.length === 0 || dirtySpans.some((span) => isFullGridRenderTileDirtySpanV3(span, input.tile.textCount))) {
    return 0
  }
  const dirtyRunCount = countDirtyTextRuns(dirtySpans, input.tile.textCount)
  if (dirtyRunCount <= 0 || dirtyRunCount > MAX_VARIABLE_TEXT_PATCH_RUNS || !input.textPayload.runPayloads) {
    return input.textPayload.quadCount
  }
  let dirtyQuadCount = 0
  const visited = new Set<number>()
  for (const span of dirtySpans) {
    const start = Math.max(0, Math.min(input.tile.textCount, span.offset))
    const end = Math.max(start, Math.min(input.tile.textCount, span.offset + span.length))
    for (let index = start; index < end; index += 1) {
      if (visited.has(index)) {
        continue
      }
      visited.add(index)
      dirtyQuadCount += input.textPayload.runPayloads[index]?.quadCount ?? 0
    }
  }
  return Math.max(input.textPayload.quadCount, input.previousTextCount + dirtyQuadCount)
}

function writeVariableLengthDirtyTextPayload(input: {
  readonly dirtySpans: readonly TextQuadRunSpan[]
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
  readonly label: string
  readonly previousRunSpans: readonly TextQuadRunSpan[]
  readonly previousTextCount: number
  readonly textPayload: {
    readonly runPayloads?: readonly TextQuadRunPayloadV3[] | undefined
    readonly runSpans: readonly TextQuadRunSpan[]
  }
  readonly tileTextCount: number
}): { readonly runSpans: readonly TextQuadRunSpan[]; readonly textCount: number } | null {
  const plan = buildVariableLengthTextPatchPlan(input)
  if (!plan) {
    return null
  }
  for (const write of plan.writes) {
    writeTypeGpuVertexBufferSubrange({
      buffer: input.handle.buffer,
      floatCount: write.floatCount,
      floats: write.floats,
      label: `${input.label}:${write.suffix}`,
      sourceStartFloat: 0,
      startFloat: write.startFloat,
    })
  }
  return { runSpans: plan.runSpans, textCount: plan.textCount }
}

function buildVariableLengthTextPatchPlan(input: {
  readonly dirtySpans: readonly TextQuadRunSpan[]
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
  readonly previousRunSpans: readonly TextQuadRunSpan[]
  readonly previousTextCount: number
  readonly textPayload: {
    readonly runPayloads?: readonly TextQuadRunPayloadV3[] | undefined
    readonly runSpans: readonly TextQuadRunSpan[]
  }
  readonly tileTextCount: number
}): VariableLengthTextPatchPlan | null {
  const runPayloads = input.textPayload.runPayloads
  if (
    !runPayloads ||
    input.previousRunSpans.length !== input.textPayload.runSpans.length ||
    input.previousRunSpans.length !== runPayloads.length
  ) {
    return null
  }
  const dirtyRunCount = countDirtyTextRuns(input.dirtySpans, input.tileTextCount)
  if (dirtyRunCount <= 0 || dirtyRunCount > MAX_VARIABLE_TEXT_PATCH_RUNS) {
    return null
  }
  const writes: VariableLengthTextPatchWrite[] = []
  const runSpans = [...input.previousRunSpans]
  const visited = new Set<number>()
  let textCount = input.previousTextCount
  for (const span of input.dirtySpans) {
    const start = Math.max(0, Math.min(input.tileTextCount, span.offset))
    const end = Math.max(start, Math.min(input.tileTextCount, span.offset + span.length))
    for (let index = start; index < end; index += 1) {
      if (visited.has(index)) {
        continue
      }
      visited.add(index)
      const previousSpan = input.previousRunSpans[index]
      const nextPayload = runPayloads[index]
      if (!previousSpan || !nextPayload) {
        return null
      }
      if (previousSpan.length > 0) {
        writes.push({
          floatCount: previousSpan.length * TEXT_INSTANCE_FLOAT_COUNT,
          floats: new Float32Array(previousSpan.length * TEXT_INSTANCE_FLOAT_COUNT),
          startFloat: previousSpan.offset * TEXT_INSTANCE_FLOAT_COUNT,
          suffix: 'text-run-clear',
        })
      }
      if (nextPayload.quadCount <= 0) {
        runSpans[index] = { offset: previousSpan.offset, length: 0 }
        continue
      }
      if (nextPayload.quadCount <= previousSpan.length) {
        runSpans[index] = { offset: previousSpan.offset, length: nextPayload.quadCount }
        writes.push({
          floatCount: nextPayload.quadCount * TEXT_INSTANCE_FLOAT_COUNT,
          floats: nextPayload.floats,
          startFloat: previousSpan.offset * TEXT_INSTANCE_FLOAT_COUNT,
          suffix: 'text-run-rewrite',
        })
        continue
      }
      const requiredTextCount = textCount + nextPayload.quadCount
      if (requiredTextCount * TEXT_INSTANCE_BYTE_COUNT > input.handle.capacityBytes) {
        return null
      }
      runSpans[index] = { offset: textCount, length: nextPayload.quadCount }
      writes.push({
        floatCount: nextPayload.quadCount * TEXT_INSTANCE_FLOAT_COUNT,
        floats: nextPayload.floats,
        startFloat: textCount * TEXT_INSTANCE_FLOAT_COUNT,
        suffix: 'text-run-append',
      })
      textCount = requiredTextCount
    }
  }
  return { runSpans, textCount, writes }
}

function writeKeyedDirtyTextPayload(input: {
  readonly dirtySpans: readonly TextQuadRunSpan[]
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
  readonly label: string
  readonly nextTextRunCellKeys: readonly string[] | null
  readonly previousRunSpans: readonly TextQuadRunSpan[] | null
  readonly previousTextCount: number
  readonly previousTextRunCellKeys: readonly string[] | null
  readonly textPayload: {
    readonly runPayloads: readonly TextQuadRunPayloadV3[]
    readonly runSpans: readonly TextQuadRunSpan[]
  }
  readonly tile: GridRenderTile
}): { readonly runSpans: readonly TextQuadRunSpan[]; readonly textCount: number } | null {
  const plan = buildKeyedDirtyTextPatchPlan(input)
  if (!plan) {
    return null
  }
  for (const write of plan.writes) {
    writeTypeGpuVertexBufferSubrange({
      buffer: input.handle.buffer,
      floatCount: write.floatCount,
      floats: write.floats,
      label: `${input.label}:${write.suffix}`,
      sourceStartFloat: 0,
      startFloat: write.startFloat,
    })
  }
  return { runSpans: plan.runSpans, textCount: plan.textCount }
}

function buildKeyedDirtyTextPatchPlan(input: {
  readonly dirtySpans: readonly TextQuadRunSpan[]
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
  readonly nextTextRunCellKeys: readonly string[] | null
  readonly previousRunSpans: readonly TextQuadRunSpan[] | null
  readonly previousTextCount: number
  readonly previousTextRunCellKeys: readonly string[] | null
  readonly textPayload: {
    readonly runPayloads: readonly TextQuadRunPayloadV3[]
    readonly runSpans: readonly TextQuadRunSpan[]
  }
  readonly tile: GridRenderTile
}): VariableLengthTextPatchPlan | null {
  if (!input.previousRunSpans || !input.previousTextRunCellKeys || !input.nextTextRunCellKeys) {
    return null
  }
  if (
    input.previousRunSpans.length !== input.previousTextRunCellKeys.length ||
    input.nextTextRunCellKeys.length !== input.textPayload.runPayloads.length
  ) {
    return null
  }
  const previousByKey = buildUniqueTextRunKeyIndexMap(input.previousTextRunCellKeys)
  const nextByKey = buildUniqueTextRunKeyIndexMap(input.nextTextRunCellKeys)
  if (!previousByKey || !nextByKey) {
    return null
  }

  const hasFullDirtySpan = input.dirtySpans.some((span) => isFullGridRenderTileDirtySpanV3(span, input.tile.textCount))
  const dirtyNextIndexes = hasFullDirtySpan ? new Set<number>() : buildDirtyTextRunIndexSet(input.dirtySpans, input.tile.textCount)
  const writes: VariableLengthTextPatchWrite[] = []
  const runSpans: TextQuadRunSpan[] = []
  let textCount = input.previousTextCount
  let patchedRunCount = 0

  for (let nextIndex = 0; nextIndex < input.nextTextRunCellKeys.length; nextIndex += 1) {
    const key = input.nextTextRunCellKeys[nextIndex]
    const previousIndex = key === undefined ? undefined : previousByKey.get(key)
    const nextPayload = input.textPayload.runPayloads[nextIndex]
    if (!key || !nextPayload) {
      return null
    }
    const isDirty =
      dirtyNextIndexes.has(nextIndex) || previousIndex === undefined || (hasFullDirtySpan && isCellKeyDirtyForTile(key, input.tile))
    if (!isDirty) {
      const previousSpan = previousIndex === undefined ? undefined : input.previousRunSpans[previousIndex]
      if (!previousSpan) {
        return null
      }
      runSpans.push(previousSpan)
      continue
    }
    if (!isCellKeyDirtyForTile(key, input.tile)) {
      return null
    }
    if (patchedRunCount >= MAX_VARIABLE_TEXT_PATCH_RUNS) {
      return null
    }
    patchedRunCount += 1
    const previousSpan = previousIndex === undefined ? null : (input.previousRunSpans[previousIndex] ?? null)
    const patch = appendVariableTextRunPatch({
      handle: input.handle,
      nextPayload,
      previousSpan,
      textCount,
      writes,
    })
    if (!patch) {
      return null
    }
    runSpans.push(patch.runSpan)
    textCount = patch.textCount
  }

  for (let previousIndex = 0; previousIndex < input.previousTextRunCellKeys.length; previousIndex += 1) {
    const key = input.previousTextRunCellKeys[previousIndex]
    if (!key || nextByKey.has(key)) {
      continue
    }
    if (!isCellKeyDirtyForTile(key, input.tile)) {
      return null
    }
    if (patchedRunCount >= MAX_VARIABLE_TEXT_PATCH_RUNS) {
      return null
    }
    patchedRunCount += 1
    const previousSpan = input.previousRunSpans[previousIndex]
    if (previousSpan && previousSpan.length > 0) {
      writes.push({
        floatCount: previousSpan.length * TEXT_INSTANCE_FLOAT_COUNT,
        floats: new Float32Array(previousSpan.length * TEXT_INSTANCE_FLOAT_COUNT),
        startFloat: previousSpan.offset * TEXT_INSTANCE_FLOAT_COUNT,
        suffix: 'text-run-clear',
      })
    }
  }

  return patchedRunCount === 0 ? null : { runSpans, textCount, writes }
}

function appendVariableTextRunPatch(input: {
  readonly handle: GpuBufferHandleV3<TextInstanceVertexBuffer>
  readonly nextPayload: TextQuadRunPayloadV3
  readonly previousSpan: TextQuadRunSpan | null
  readonly textCount: number
  readonly writes: VariableLengthTextPatchWrite[]
}): { readonly runSpan: TextQuadRunSpan; readonly textCount: number } | null {
  if (input.previousSpan && input.previousSpan.length > 0) {
    input.writes.push({
      floatCount: input.previousSpan.length * TEXT_INSTANCE_FLOAT_COUNT,
      floats: new Float32Array(input.previousSpan.length * TEXT_INSTANCE_FLOAT_COUNT),
      startFloat: input.previousSpan.offset * TEXT_INSTANCE_FLOAT_COUNT,
      suffix: 'text-run-clear',
    })
  }
  if (input.nextPayload.quadCount <= 0) {
    return { runSpan: { offset: input.previousSpan?.offset ?? input.textCount, length: 0 }, textCount: input.textCount }
  }
  if (input.previousSpan && input.nextPayload.quadCount <= input.previousSpan.length) {
    input.writes.push({
      floatCount: input.nextPayload.quadCount * TEXT_INSTANCE_FLOAT_COUNT,
      floats: input.nextPayload.floats,
      startFloat: input.previousSpan.offset * TEXT_INSTANCE_FLOAT_COUNT,
      suffix: 'text-run-rewrite',
    })
    return {
      runSpan: { offset: input.previousSpan.offset, length: input.nextPayload.quadCount },
      textCount: input.textCount,
    }
  }
  const requiredTextCount = input.textCount + input.nextPayload.quadCount
  if (requiredTextCount * TEXT_INSTANCE_BYTE_COUNT > input.handle.capacityBytes) {
    return null
  }
  input.writes.push({
    floatCount: input.nextPayload.quadCount * TEXT_INSTANCE_FLOAT_COUNT,
    floats: input.nextPayload.floats,
    startFloat: input.textCount * TEXT_INSTANCE_FLOAT_COUNT,
    suffix: 'text-run-append',
  })
  return {
    runSpan: { offset: input.textCount, length: input.nextPayload.quadCount },
    textCount: requiredTextCount,
  }
}

function countDirtyTextRuns(dirtySpans: readonly TextQuadRunSpan[], textRunCount: number): number {
  let count = 0
  for (const span of dirtySpans) {
    const start = Math.max(0, Math.min(textRunCount, span.offset))
    const end = Math.max(start, Math.min(textRunCount, span.offset + span.length))
    count += end - start
  }
  return count
}

export function buildTextRunCellKeys(textRuns: readonly GridRenderTile['textRuns'][number][]): readonly string[] | null {
  const keys: string[] = []
  for (const run of textRuns) {
    if (run.row === undefined || run.col === undefined) {
      return null
    }
    keys.push(`${run.row}:${run.col}`)
  }
  return keys
}

function buildUniqueTextRunKeyIndexMap(keys: readonly string[]): Map<string, number> | null {
  const indexes = new Map<string, number>()
  for (let index = 0; index < keys.length; index += 1) {
    const key = keys[index]
    if (!key || indexes.has(key)) {
      return null
    }
    indexes.set(key, index)
  }
  return indexes
}

function buildDirtyTextRunIndexSet(dirtySpans: readonly TextQuadRunSpan[], textRunCount: number): ReadonlySet<number> {
  const indexes = new Set<number>()
  for (const span of dirtySpans) {
    const start = Math.max(0, Math.min(textRunCount, span.offset))
    const end = Math.max(start, Math.min(textRunCount, span.offset + span.length))
    for (let index = start; index < end; index += 1) {
      indexes.add(index)
    }
  }
  return indexes
}

export function isCellKeyDirtyForTile(key: string, tile: GridRenderTile): boolean {
  const parsed = parseTextRunCellKey(key)
  const dirtyMasks = tile.dirtyMasks
  const dirtyLocalRows = tile.dirtyLocalRows
  const dirtyLocalCols = tile.dirtyLocalCols
  if (!parsed || !dirtyMasks || !dirtyLocalRows || !dirtyLocalCols || dirtyLocalRows.length !== dirtyMasks.length * 2) {
    return false
  }
  if (dirtyLocalCols.length !== dirtyMasks.length * 2) {
    return false
  }
  const localRow = parsed.row - tile.bounds.rowStart
  const localCol = parsed.col - tile.bounds.colStart
  for (let index = 0; index < dirtyMasks.length; index += 1) {
    const mask = dirtyMasks[index] ?? 0
    if ((mask & (DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Style)) === 0) {
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

function parseTextRunCellKey(key: string): { readonly col: number; readonly row: number } | null {
  const separator = key.indexOf(':')
  if (separator <= 0 || separator >= key.length - 1) {
    return null
  }
  const row = Number.parseInt(key.slice(0, separator), 10)
  const col = Number.parseInt(key.slice(separator + 1), 10)
  return Number.isFinite(row) && Number.isFinite(col) ? { col, row } : null
}

function resolveStableTextQuadSpan(input: {
  readonly dirtySpan: { readonly offset: number; readonly length: number }
  readonly nextRunSpans: readonly TextQuadRunSpan[]
  readonly previousRunSpans: readonly TextQuadRunSpan[]
}): TextQuadRunSpan | null {
  const start = input.dirtySpan.offset
  const endExclusive = start + input.dirtySpan.length
  if (start < 0 || endExclusive > input.nextRunSpans.length || input.dirtySpan.length <= 0) {
    return null
  }

  let offset = Number.MAX_SAFE_INTEGER
  let end = 0
  for (let index = start; index < endExclusive; index += 1) {
    const next = input.nextRunSpans[index]
    const previous = input.previousRunSpans[index]
    if (!next || !previous || next.offset !== previous.offset || next.length !== previous.length) {
      return null
    }
    offset = Math.min(offset, next.offset)
    end = Math.max(end, next.offset + next.length)
  }
  return offset === Number.MAX_SAFE_INTEGER ? null : { offset, length: end - offset }
}

export function mergeTextRunDirtySpans(
  tileDirtySpans: readonly TextQuadRunSpan[] | undefined,
  missingGlyphRunSpans: readonly TextQuadRunSpan[],
): readonly TextQuadRunSpan[] | undefined {
  if (!tileDirtySpans || tileDirtySpans.length === 0) {
    return missingGlyphRunSpans.length === 0 ? undefined : missingGlyphRunSpans
  }
  if (missingGlyphRunSpans.length === 0) {
    return tileDirtySpans
  }
  return [...tileDirtySpans, ...missingGlyphRunSpans]
}
