export interface BiligRenderedCanvasState {
  readonly authoritativeRenderRevision?: string | null | undefined
  readonly backendStatus?: string | null | undefined
  readonly currentContentSignature?: string | null | undefined
  readonly currentRectCount?: number | undefined
  readonly currentRectSignature?: string | null | undefined
  readonly currentTextRunCount?: number | undefined
  readonly currentTextSignature?: string | null | undefined
  readonly frameProofStatus?: string | null | undefined
  readonly frameProofSignature?: string | null | undefined
  readonly headerPaneCount: number
  readonly hasPresentedFrame?: boolean | undefined
  readonly hasPresentedVisibleFrame?: boolean | undefined
  readonly localRenderRevision?: string | null | undefined
  readonly mode: string | null
  readonly pixelHeight: number
  readonly pixelWidth: number
  readonly presentedContentSignature?: string | null | undefined
  readonly presentedFrameProofSignature?: string | null | undefined
  readonly presentedHeaderPaneCount?: number | undefined
  readonly presentedRectCount?: number | undefined
  readonly presentedRectSignature?: string | null | undefined
  readonly presentedTextRunCount?: number | undefined
  readonly presentedTextSignature?: string | null | undefined
  readonly presentedTilePaneCount?: number | undefined
  readonly projectedRenderRevision?: string | null | undefined
  readonly tilePaneCount: number
  readonly tileSceneRevision?: string | null | undefined
  readonly visibleAuthoritativeRenderRevision?: string | null | undefined
  readonly visibleLocalRenderRevision?: string | null | undefined
  readonly visibleProjectedRenderRevision?: string | null | undefined
  readonly visibleRenderRevision?: string | null | undefined
  readonly visiblePixelCount?: number | undefined
}

export interface BiligRenderedSurfaceState {
  readonly dpr: number
  readonly fallback: BiligRenderedCanvasState | null
  readonly gridAuthoritativeRenderRevision?: string | null | undefined
  readonly gridHeight: number
  readonly gridLocalRenderRevision?: string | null | undefined
  readonly gridProjectedRenderRevision?: string | null | undefined
  readonly gridWidth: number
  readonly typeGpu: BiligRenderedCanvasState | null
}

export interface BiligRenderedSurfaceReadiness {
  readonly evidence: readonly string[]
  readonly gaps: readonly string[]
  readonly ready: boolean
}

export function isBiligRenderedSurfaceReady(state: BiligRenderedSurfaceState | null): boolean {
  return biligRenderedSurfaceReadiness(state).ready
}

export function biligRenderedSurfaceReadiness(state: BiligRenderedSurfaceState | null): BiligRenderedSurfaceReadiness {
  if (!state) {
    return {
      evidence: ['missing sheet grid surface'],
      gaps: ['missing sheet grid surface'],
      ready: false,
    }
  }
  const gaps: string[] = []
  if (state.gridWidth <= 0 || state.gridHeight <= 0) {
    gaps.push('grid viewport has no measurable size')
  }
  if (state.fallback) {
    gaps.push('fallback canvas is mounted')
  }
  const canvas = state.typeGpu
  if (!canvas) {
    gaps.push('TypeGPU canvas is missing')
    return {
      evidence: [...baseSurfaceEvidence(state), ...gaps.map((gap) => `gap=${gap}`)],
      gaps,
      ready: false,
    }
  }
  if (canvas.mode !== 'typegpu-v3') {
    gaps.push(`renderer mode is ${canvas.mode ?? 'missing'}`)
  }
  if (canvas.backendStatus !== 'ready') {
    gaps.push(`TypeGPU backend is ${canvas.backendStatus ?? 'missing'}`)
  }
  if (canvas.frameProofStatus !== 'presented') {
    gaps.push(`frame proof is ${canvas.frameProofStatus ?? 'missing'}`)
  }
  if (!hasText(canvas.frameProofSignature)) {
    gaps.push('frame proof signature is missing')
  }
  if (!hasText(canvas.presentedFrameProofSignature)) {
    gaps.push('presented frame proof signature is missing')
  }
  if (canvas.hasPresentedFrame !== true) {
    gaps.push('current frame signature has not been presented')
  }
  if (
    hasText(canvas.frameProofSignature) &&
    hasText(canvas.presentedFrameProofSignature) &&
    canvas.presentedFrameProofSignature !== canvas.frameProofSignature
  ) {
    gaps.push('presented frame proof signature does not match current frame')
  }
  if (canvas.hasPresentedVisibleFrame !== true) {
    gaps.push('visible frame has not been presented')
  }
  if (canvas.tilePaneCount <= 0 || canvas.headerPaneCount <= 0) {
    gaps.push('current tile/header pane counts are empty')
  }
  if ((canvas.presentedTilePaneCount ?? 0) <= 0 || (canvas.presentedHeaderPaneCount ?? 0) <= 0) {
    gaps.push('presented tile/header pane counts are empty')
  }
  if (
    canvas.tilePaneCount > 0 &&
    canvas.headerPaneCount > 0 &&
    (canvas.presentedTilePaneCount ?? 0) > 0 &&
    (canvas.presentedHeaderPaneCount ?? 0) > 0 &&
    (canvas.presentedTilePaneCount !== canvas.tilePaneCount || canvas.presentedHeaderPaneCount !== canvas.headerPaneCount)
  ) {
    gaps.push('presented tile/header pane counts do not cover the current visible panes')
  }
  if (!canvasPixelsMatchViewport(canvas, state)) {
    gaps.push('TypeGPU canvas backing pixels do not cover the viewport')
  }
  if (!hasText(canvas.currentContentSignature)) {
    gaps.push('current visible content signature is missing')
  }
  if (!hasText(canvas.presentedContentSignature)) {
    gaps.push('presented visible content signature is missing')
  }
  if (
    hasText(canvas.currentContentSignature) &&
    hasText(canvas.presentedContentSignature) &&
    canvas.presentedContentSignature !== canvas.currentContentSignature
  ) {
    gaps.push('presented visible content signature does not match current tiles')
  }
  if (!hasText(canvas.currentTextSignature)) {
    gaps.push('current visible text signature is missing')
  }
  if (!hasText(canvas.presentedTextSignature)) {
    gaps.push('presented visible text signature is missing')
  }
  if (
    hasText(canvas.currentTextSignature) &&
    hasText(canvas.presentedTextSignature) &&
    canvas.presentedTextSignature !== canvas.currentTextSignature
  ) {
    gaps.push('presented visible text signature does not match current tiles')
  }
  if (!hasText(canvas.currentRectSignature)) {
    gaps.push('current visible rect signature is missing')
  }
  if (!hasText(canvas.presentedRectSignature)) {
    gaps.push('presented visible rect signature is missing')
  }
  if (
    hasText(canvas.currentRectSignature) &&
    hasText(canvas.presentedRectSignature) &&
    canvas.presentedRectSignature !== canvas.currentRectSignature
  ) {
    gaps.push('presented visible rect signature does not match current tiles')
  }
  if ((canvas.currentRectCount ?? 0) <= 0 || (canvas.presentedRectCount ?? 0) <= 0) {
    gaps.push('visible rect payload counts are empty')
  }
  if ((canvas.currentTextRunCount ?? 0) !== (canvas.presentedTextRunCount ?? 0)) {
    gaps.push('presented visible text run count does not match current tiles')
  }
  if ((canvas.currentRectCount ?? 0) !== (canvas.presentedRectCount ?? 0)) {
    gaps.push('presented visible rect count does not match current tiles')
  }
  if (!hasText(state.gridAuthoritativeRenderRevision)) {
    gaps.push('grid authoritative render revision is missing')
  }
  if (!hasText(canvas.authoritativeRenderRevision)) {
    gaps.push('TypeGPU authoritative render revision is missing')
  }
  if (!hasText(canvas.visibleAuthoritativeRenderRevision)) {
    gaps.push('visible authoritative render revision is missing')
  }
  if (
    hasText(state.gridAuthoritativeRenderRevision) &&
    hasText(canvas.authoritativeRenderRevision) &&
    canvas.authoritativeRenderRevision !== state.gridAuthoritativeRenderRevision
  ) {
    gaps.push('TypeGPU authoritative render revision does not match the grid revision')
  }
  if (
    hasText(state.gridAuthoritativeRenderRevision) &&
    hasText(canvas.visibleAuthoritativeRenderRevision) &&
    canvas.visibleAuthoritativeRenderRevision !== state.gridAuthoritativeRenderRevision
  ) {
    gaps.push('visible authoritative render revision does not match the grid revision')
  }
  if (!hasText(state.gridLocalRenderRevision)) {
    gaps.push('grid local render revision is missing')
  }
  if (!hasText(canvas.localRenderRevision)) {
    gaps.push('TypeGPU local render revision is missing')
  }
  if (!hasText(canvas.visibleLocalRenderRevision)) {
    gaps.push('visible local render revision is missing')
  }
  if (
    hasText(state.gridLocalRenderRevision) &&
    hasText(canvas.localRenderRevision) &&
    canvas.localRenderRevision !== state.gridLocalRenderRevision
  ) {
    gaps.push('TypeGPU local render revision does not match the grid revision')
  }
  if (
    hasText(state.gridLocalRenderRevision) &&
    hasText(canvas.visibleLocalRenderRevision) &&
    canvas.visibleLocalRenderRevision !== state.gridLocalRenderRevision
  ) {
    gaps.push('visible local render revision does not match the grid revision')
  }
  if (!hasText(state.gridProjectedRenderRevision)) {
    gaps.push('grid projected render revision is missing')
  }
  if (!hasText(canvas.projectedRenderRevision)) {
    gaps.push('TypeGPU projected render revision is missing')
  }
  if (!hasText(canvas.visibleProjectedRenderRevision)) {
    gaps.push('visible projected render revision is missing')
  }
  if (
    hasText(state.gridProjectedRenderRevision) &&
    hasText(canvas.projectedRenderRevision) &&
    canvas.projectedRenderRevision !== state.gridProjectedRenderRevision
  ) {
    gaps.push('TypeGPU projected render revision does not match the grid revision')
  }
  if (
    hasText(state.gridProjectedRenderRevision) &&
    hasText(canvas.visibleProjectedRenderRevision) &&
    canvas.visibleProjectedRenderRevision !== state.gridProjectedRenderRevision
  ) {
    gaps.push('visible projected render revision does not match the grid revision')
  }
  if (!hasText(canvas.tileSceneRevision)) {
    gaps.push('tile scene revision is missing')
  }
  if (!hasText(canvas.visibleRenderRevision)) {
    gaps.push('visible render revision is missing')
  }
  if (
    hasText(canvas.tileSceneRevision) &&
    hasText(canvas.visibleRenderRevision) &&
    canvas.visibleRenderRevision !== canvas.tileSceneRevision
  ) {
    gaps.push('visible render revision does not match the tile scene revision')
  }
  return {
    evidence: [...baseSurfaceEvidence(state), ...canvasEvidence(canvas, state), ...gaps.map((gap) => `gap=${gap}`)],
    gaps,
    ready: gaps.length === 0,
  }
}

function canvasPixelsMatchViewport(canvas: BiligRenderedCanvasState, state: BiligRenderedSurfaceState): boolean {
  const expectedPixelWidth = Math.max(1, Math.floor(state.gridWidth * state.dpr))
  const expectedPixelHeight = Math.max(1, Math.floor(state.gridHeight * state.dpr))
  return canvas.pixelWidth >= expectedPixelWidth - 2 && canvas.pixelHeight >= expectedPixelHeight - 2
}

function baseSurfaceEvidence(state: BiligRenderedSurfaceState): string[] {
  const expectedPixelWidth = Math.max(1, Math.floor(state.gridWidth * state.dpr))
  const expectedPixelHeight = Math.max(1, Math.floor(state.gridHeight * state.dpr))
  return [
    `gridCssWidth=${String(state.gridWidth)}`,
    `gridCssHeight=${String(state.gridHeight)}`,
    `devicePixelRatio=${String(state.dpr)}`,
    `expectedPixelWidth=${String(expectedPixelWidth)}`,
    `expectedPixelHeight=${String(expectedPixelHeight)}`,
    `gridAuthoritativeRevision=${state.gridAuthoritativeRenderRevision ?? ''}`,
    `gridLocalRevision=${state.gridLocalRenderRevision ?? ''}`,
    `gridProjectedRevision=${state.gridProjectedRenderRevision ?? ''}`,
    `fallbackMounted=${String(Boolean(state.fallback))}`,
  ]
}

function canvasEvidence(canvas: BiligRenderedCanvasState, state: BiligRenderedSurfaceState): string[] {
  return [
    `mode=${canvas.mode ?? ''}`,
    `backendStatus=${canvas.backendStatus ?? ''}`,
    `frameProofStatus=${canvas.frameProofStatus ?? ''}`,
    `frameProofSignature=${canvas.frameProofSignature ?? ''}`,
    `hasPresentedFrame=${String(canvas.hasPresentedFrame === true)}`,
    `hasPresentedVisibleFrame=${String(canvas.hasPresentedVisibleFrame === true)}`,
    `presentedFrameProofSignature=${canvas.presentedFrameProofSignature ?? ''}`,
    `currentContentSignature=${canvas.currentContentSignature ?? ''}`,
    `presentedContentSignature=${canvas.presentedContentSignature ?? ''}`,
    `currentTextRunCount=${String(canvas.currentTextRunCount ?? 0)}`,
    `presentedTextRunCount=${String(canvas.presentedTextRunCount ?? 0)}`,
    `currentTextSignature=${canvas.currentTextSignature ?? ''}`,
    `presentedTextSignature=${canvas.presentedTextSignature ?? ''}`,
    `currentRectCount=${String(canvas.currentRectCount ?? 0)}`,
    `presentedRectCount=${String(canvas.presentedRectCount ?? 0)}`,
    `currentRectSignature=${canvas.currentRectSignature ?? ''}`,
    `presentedRectSignature=${canvas.presentedRectSignature ?? ''}`,
    `tilePaneCount=${String(canvas.tilePaneCount)}`,
    `headerPaneCount=${String(canvas.headerPaneCount)}`,
    `presentedTilePaneCount=${String(canvas.presentedTilePaneCount ?? 0)}`,
    `presentedHeaderPaneCount=${String(canvas.presentedHeaderPaneCount ?? 0)}`,
    `canvasPixelWidth=${String(canvas.pixelWidth)}`,
    `canvasPixelHeight=${String(canvas.pixelHeight)}`,
    `canvasCoversViewport=${String(canvasPixelsMatchViewport(canvas, state))}`,
    `typeGpuAuthoritativeRevision=${canvas.authoritativeRenderRevision ?? ''}`,
    `typeGpuLocalRevision=${canvas.localRenderRevision ?? ''}`,
    `typeGpuProjectedRevision=${canvas.projectedRenderRevision ?? ''}`,
    `visibleAuthoritativeRevision=${canvas.visibleAuthoritativeRenderRevision ?? ''}`,
    `visibleLocalRevision=${canvas.visibleLocalRenderRevision ?? ''}`,
    `visibleProjectedRevision=${canvas.visibleProjectedRenderRevision ?? ''}`,
    `tileSceneRevision=${canvas.tileSceneRevision ?? ''}`,
    `visibleRenderRevision=${canvas.visibleRenderRevision ?? ''}`,
  ]
}

function hasText(value: string | null | undefined): value is string {
  return typeof value === 'string' && value.length > 0
}
