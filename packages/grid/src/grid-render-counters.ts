export interface RendererTileReadinessCounterInput {
  readonly exactHits: number
  readonly staleHits: number
  readonly misses: number
  readonly visibleDirtyTiles: number
  readonly warmDirtyTiles: number
}

export interface TypeGpuTileCacheStaleLookupCounterInput {
  readonly hits: number
  readonly lookups: number
  readonly scannedEntries: number
}

export interface TypeGpuTextPayloadCounterInput {
  readonly reusedRunPayloads: number
  readonly rebuiltRunPayloads: number
  readonly atlasGeometryRetries: number
  readonly atlasGeometryResyncs?: number | undefined
  readonly axisOnlySyncAccepts?: number | undefined
  readonly axisOnlySyncRejects?: number | undefined
  readonly axisOnlySyncMissingGlyphRejects?: number | undefined
  readonly axisOnlySyncSignatureRejects?: number | undefined
  readonly axisOnlySyncFallbackRebuilds?: number | undefined
  readonly axisOnlySyncAuthoritativeFullTile?: number | undefined
  readonly glyphDependencies: number
  readonly pageDependencies: number
  readonly textBuildMs?: number | undefined
  readonly textDecorationMs?: number | undefined
  readonly textSyncMs?: number | undefined
  readonly textWriteMs?: number | undefined
}

type ScrollPerfCounterSink = Partial<{
  noteTypeGpuConfigure: () => void
  noteTypeGpuSubmit: () => void
  noteTypeGpuDrawCall: (count: number) => void
  noteTypeGpuPaneDraw: (count: number) => void
  noteTypeGpuUniformWrite: (bytes: number, label: string) => void
  noteTypeGpuBufferWrite: (bytes: number, label: string) => void
  noteTypeGpuOverlayWrite: (bytes: number) => void
  noteTypeGpuBufferAllocation: (bytes: number, label: string) => void
  noteTypeGpuAtlasUpload: (bytes: number) => void
  noteTypeGpuAtlasDirtyPageUpload: (bytes: number, pageCount: number) => void
  noteTypeGpuSurfaceResize: (width: number, height: number, dpr: number) => void
  noteTypeGpuTileMiss: (tileKey: number | string) => void
  noteTypeGpuTileCacheEviction: (count: number) => void
  noteTypeGpuTileCacheStaleLookups: (input: TypeGpuTileCacheStaleLookupCounterInput) => void
  noteTypeGpuTileCacheVisibleMark: (count: number) => void
  noteTypeGpuTextPayload: (input: TypeGpuTextPayloadCounterInput) => void
  noteRendererTileReadiness: (input: RendererTileReadinessCounterInput) => void
  noteGridScrollInput: (timestamp: number) => void
  noteGridDrawFrame: (timestamp: number) => void
}>

function getCounterSink(): ScrollPerfCounterSink | null {
  if (typeof window === 'undefined') {
    return null
  }
  return (window as Window & { __biligScrollPerf?: ScrollPerfCounterSink }).__biligScrollPerf ?? null
}

export function noteTypeGpuConfigure(): void {
  getCounterSink()?.noteTypeGpuConfigure?.()
}

export function noteTypeGpuSubmit(): void {
  getCounterSink()?.noteTypeGpuSubmit?.()
}

export function noteTypeGpuDrawCall(count = 1): void {
  getCounterSink()?.noteTypeGpuDrawCall?.(count)
}

export function noteTypeGpuPaneDraw(count = 1): void {
  getCounterSink()?.noteTypeGpuPaneDraw?.(count)
}

export function noteTypeGpuUniformWrite(bytes: number, label: string): void {
  getCounterSink()?.noteTypeGpuUniformWrite?.(bytes, label)
}

export function noteTypeGpuBufferWrite(bytes: number, label: string): void {
  getCounterSink()?.noteTypeGpuBufferWrite?.(bytes, label)
  if (label.startsWith('overlay:')) {
    getCounterSink()?.noteTypeGpuOverlayWrite?.(bytes)
  }
}

export function noteTypeGpuBufferAllocation(bytes: number, label: string): void {
  getCounterSink()?.noteTypeGpuBufferAllocation?.(bytes, label)
}

export function noteTypeGpuAtlasUpload(bytes: number): void {
  getCounterSink()?.noteTypeGpuAtlasUpload?.(bytes)
}

export function noteTypeGpuAtlasDirtyPageUpload(bytes: number, pageCount: number): void {
  getCounterSink()?.noteTypeGpuAtlasDirtyPageUpload?.(bytes, pageCount)
}

export function noteTypeGpuSurfaceResize(width: number, height: number, dpr: number): void {
  getCounterSink()?.noteTypeGpuSurfaceResize?.(width, height, dpr)
}

export function noteTypeGpuTileMiss(tileKey: number | string): void {
  getCounterSink()?.noteTypeGpuTileMiss?.(tileKey)
}

export function noteTypeGpuTileCacheEviction(count = 1): void {
  getCounterSink()?.noteTypeGpuTileCacheEviction?.(count)
}

export function noteTypeGpuTileCacheStaleLookups(input: TypeGpuTileCacheStaleLookupCounterInput): void {
  getCounterSink()?.noteTypeGpuTileCacheStaleLookups?.(input)
}

export function noteTypeGpuTileCacheVisibleMark(count: number): void {
  getCounterSink()?.noteTypeGpuTileCacheVisibleMark?.(count)
}

export function noteTypeGpuTextPayload(input: TypeGpuTextPayloadCounterInput): void {
  getCounterSink()?.noteTypeGpuTextPayload?.(input)
}

export function noteRendererTileReadiness(input: RendererTileReadinessCounterInput): void {
  getCounterSink()?.noteRendererTileReadiness?.(input)
}

export function noteGridScrollInput(timestamp = performance.now()): void {
  getCounterSink()?.noteGridScrollInput?.(timestamp)
}

export function noteGridDrawFrame(timestamp = performance.now()): void {
  getCounterSink()?.noteGridDrawFrame?.(timestamp)
}
