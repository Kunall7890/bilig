import type { GridGpuColor, GridGpuRect, GridGpuScene } from '../gridGpuPrimitives.js'

export const GRID_RECT_FLOAT_COUNT_V3 = 8
export const GRID_RECT_INSTANCE_FLOAT_COUNT_V3 = 20

export interface PackedGridRectBufferV3 {
  readonly rects: Float32Array
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly fillRectCount: number
  readonly borderRectCount: number
  readonly rectSignature: string
}

export function packGridRectBufferV3(
  scene: GridGpuScene,
  surfaceSize: { readonly width: number; readonly height: number },
): PackedGridRectBufferV3 {
  const fillRects = coalesceGridRectsV3(scene.fillRects)
  const borderRects = coalesceGridRectsV3(scene.borderRects)
  const coalescedScene = {
    borderRects,
    fillRects,
  }
  return {
    borderRectCount: borderRects.length,
    fillRectCount: fillRects.length,
    rectCount: fillRects.length + borderRects.length,
    rectInstances: packGridRectInstancesV3(coalescedScene, surfaceSize),
    rects: packGridRectsV3(coalescedScene),
    rectSignature: resolveGridRectSignatureV3(coalescedScene, surfaceSize),
  }
}

function packGridRectInstancesV3(scene: GridGpuScene, surfaceSize: { readonly width: number; readonly height: number }): Float32Array {
  const rectCount = scene.fillRects.length + scene.borderRects.length
  const floats = new Float32Array(Math.max(1, rectCount) * GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
  const clipX = 0
  const clipY = 0
  const clipX1 = surfaceSize.width
  const clipY1 = surfaceSize.height
  let offset = 0
  for (const rect of scene.fillRects) {
    offset = writeFillRectInstance(floats, offset, rect, clipX, clipY, clipX1, clipY1)
  }
  for (const rect of scene.borderRects) {
    offset = writeBorderRectInstance(floats, offset, rect, clipX, clipY, clipX1, clipY1)
  }
  return floats
}

function packGridRectsV3(scene: GridGpuScene): Float32Array {
  const rectCount = scene.fillRects.length + scene.borderRects.length
  const floats = new Float32Array(Math.max(1, rectCount) * GRID_RECT_FLOAT_COUNT_V3)
  let offset = 0
  for (const rect of scene.fillRects) {
    offset = writeGridRect(floats, offset, rect)
  }
  for (const rect of scene.borderRects) {
    offset = writeGridRect(floats, offset, rect)
  }
  return floats
}

function resolveGridRectSignatureV3(scene: GridGpuScene, surfaceSize: { readonly width: number; readonly height: number }): string {
  let hash = createHash()
  hash = mixNumber(hash, surfaceSize.width)
  hash = mixNumber(hash, surfaceSize.height)
  hash = mixNumber(hash, scene.fillRects.length)
  hash = mixNumber(hash, scene.borderRects.length)
  for (const rect of scene.fillRects) {
    hash = mixRect(hash, rect)
  }
  for (const rect of scene.borderRects) {
    hash = mixRect(hash, rect)
  }
  return hash.toString(36)
}

export function coalesceGridRectsV3(rects: readonly GridGpuRect[]): readonly GridGpuRect[] {
  if (rects.length <= 1) {
    return rects
  }
  const coalesced: GridGpuRect[] = []
  const horizontalRuns = new Map<string, Map<string, number>>()
  const verticalRuns = new Map<string, Map<string, number>>()

  for (const rect of rects) {
    if (rect.width <= 0 || rect.height <= 0) {
      continue
    }
    const horizontalKey = createHorizontalRunKey(rect)
    const horizontalEndKey = coordinateKey(rect.x)
    const horizontalIndex = horizontalRuns.get(horizontalKey)?.get(horizontalEndKey)
    if (horizontalIndex !== undefined && mergeHorizontalRect(coalesced, horizontalRuns, horizontalKey, horizontalIndex, rect)) {
      continue
    }

    const verticalKey = createVerticalRunKey(rect)
    const verticalEndKey = coordinateKey(rect.y)
    const verticalIndex = verticalRuns.get(verticalKey)?.get(verticalEndKey)
    if (verticalIndex !== undefined && mergeVerticalRect(coalesced, verticalRuns, verticalKey, verticalIndex, rect)) {
      continue
    }

    const index = coalesced.length
    coalesced.push(rect)
    setRunIndex(horizontalRuns, horizontalKey, coordinateKey(rect.x + rect.width), index)
    setRunIndex(verticalRuns, verticalKey, coordinateKey(rect.y + rect.height), index)
  }

  return coalesced
}

function mergeHorizontalRect(
  rects: GridGpuRect[],
  runs: Map<string, Map<string, number>>,
  runKey: string,
  index: number,
  rect: GridGpuRect,
): boolean {
  const previous = rects[index]
  if (
    !previous ||
    !sameGridRectColor(previous.color, rect.color) ||
    !sameNumber(previous.y, rect.y) ||
    !sameNumber(previous.height, rect.height)
  ) {
    return false
  }
  const previousEnd = previous.x + previous.width
  if (!sameNumber(previousEnd, rect.x)) {
    return false
  }
  runs.get(runKey)?.delete(coordinateKey(previousEnd))
  rects[index] = {
    ...previous,
    width: rect.x + rect.width - previous.x,
  }
  setRunIndex(runs, runKey, coordinateKey(rect.x + rect.width), index)
  return true
}

function mergeVerticalRect(
  rects: GridGpuRect[],
  runs: Map<string, Map<string, number>>,
  runKey: string,
  index: number,
  rect: GridGpuRect,
): boolean {
  const previous = rects[index]
  if (
    !previous ||
    !sameGridRectColor(previous.color, rect.color) ||
    !sameNumber(previous.x, rect.x) ||
    !sameNumber(previous.width, rect.width)
  ) {
    return false
  }
  const previousEnd = previous.y + previous.height
  if (!sameNumber(previousEnd, rect.y)) {
    return false
  }
  runs.get(runKey)?.delete(coordinateKey(previousEnd))
  rects[index] = {
    ...previous,
    height: rect.y + rect.height - previous.y,
  }
  setRunIndex(runs, runKey, coordinateKey(rect.y + rect.height), index)
  return true
}

function setRunIndex(runs: Map<string, Map<string, number>>, runKey: string, endKey: string, index: number): void {
  let entries = runs.get(runKey)
  if (!entries) {
    entries = new Map()
    runs.set(runKey, entries)
  }
  entries.set(endKey, index)
}

function createHorizontalRunKey(rect: GridGpuRect): string {
  return ['h', coordinateKey(rect.y), coordinateKey(rect.height), colorKey(rect.color)].join('|')
}

function createVerticalRunKey(rect: GridGpuRect): string {
  return ['v', coordinateKey(rect.x), coordinateKey(rect.width), colorKey(rect.color)].join('|')
}

function colorKey(color: GridGpuColor): string {
  return [coordinateKey(color.r), coordinateKey(color.g), coordinateKey(color.b), coordinateKey(color.a)].join(',')
}

function coordinateKey(value: number): string {
  return String(Math.round(value * 1000))
}

function sameNumber(left: number, right: number): boolean {
  return coordinateKey(left) === coordinateKey(right)
}

function sameGridRectColor(left: GridGpuColor, right: GridGpuColor): boolean {
  return sameNumber(left.r, right.r) && sameNumber(left.g, right.g) && sameNumber(left.b, right.b) && sameNumber(left.a, right.a)
}

function writeFillRectInstance(
  floats: Float32Array,
  offset: number,
  rect: GridGpuRect,
  clipX: number,
  clipY: number,
  clipX1: number,
  clipY1: number,
): number {
  floats[offset + 0] = rect.x
  floats[offset + 1] = rect.y
  floats[offset + 2] = rect.width
  floats[offset + 3] = rect.height
  floats[offset + 4] = rect.color.r
  floats[offset + 5] = rect.color.g
  floats[offset + 6] = rect.color.b
  floats[offset + 7] = rect.color.a
  floats[offset + 8] = 0
  floats[offset + 9] = 0
  floats[offset + 10] = 0
  floats[offset + 11] = 0
  floats[offset + 12] = rect.color.a < 0.2 ? 2 : 0
  floats[offset + 13] = 0
  floats[offset + 14] = 0
  floats[offset + 15] = 0
  floats[offset + 16] = clipX
  floats[offset + 17] = clipY
  floats[offset + 18] = clipX1
  floats[offset + 19] = clipY1
  return offset + GRID_RECT_INSTANCE_FLOAT_COUNT_V3
}

function writeBorderRectInstance(
  floats: Float32Array,
  offset: number,
  rect: GridGpuRect,
  clipX: number,
  clipY: number,
  clipX1: number,
  clipY1: number,
): number {
  floats[offset + 0] = rect.x
  floats[offset + 1] = rect.y
  floats[offset + 2] = rect.width
  floats[offset + 3] = rect.height
  floats[offset + 4] = 0
  floats[offset + 5] = 0
  floats[offset + 6] = 0
  floats[offset + 7] = 0
  floats[offset + 8] = rect.color.r
  floats[offset + 9] = rect.color.g
  floats[offset + 10] = rect.color.b
  floats[offset + 11] = rect.color.a
  floats[offset + 12] = 0
  floats[offset + 13] = 1
  floats[offset + 14] = 0
  floats[offset + 15] = 0
  floats[offset + 16] = clipX
  floats[offset + 17] = clipY
  floats[offset + 18] = clipX1
  floats[offset + 19] = clipY1
  return offset + GRID_RECT_INSTANCE_FLOAT_COUNT_V3
}

function writeGridRect(floats: Float32Array, offset: number, rect: GridGpuRect): number {
  floats[offset + 0] = rect.x
  floats[offset + 1] = rect.y
  floats[offset + 2] = rect.width
  floats[offset + 3] = rect.height
  floats[offset + 4] = rect.color.r
  floats[offset + 5] = rect.color.g
  floats[offset + 6] = rect.color.b
  floats[offset + 7] = rect.color.a
  return offset + GRID_RECT_FLOAT_COUNT_V3
}

function mixRect(hash: number, rect: GridGpuRect): number {
  let next = hash
  next = mixNumber(next, rect.x)
  next = mixNumber(next, rect.y)
  next = mixNumber(next, rect.width)
  next = mixNumber(next, rect.height)
  next = mixColor(next, rect.color)
  return next
}

function mixColor(hash: number, color: GridGpuColor): number {
  let next = hash
  next = mixNumber(next, color.r)
  next = mixNumber(next, color.g)
  next = mixNumber(next, color.b)
  next = mixNumber(next, color.a)
  return next
}

function createHash(): number {
  return 2_166_136_261
}

function mixNumber(hash: number, value: number): number {
  return mixInteger(hash, Math.round(value * 1_000))
}

function mixInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}
