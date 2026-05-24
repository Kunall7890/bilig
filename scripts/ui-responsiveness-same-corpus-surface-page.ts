import type { Page } from '@playwright/test'

import type { BiligRenderedCanvasState, BiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface.ts'

export async function readBiligRenderedSurfaceState(page: Page): Promise<BiligRenderedSurfaceState | null> {
  return await page.evaluate(() => {
    const grid = document.querySelector('[data-testid="sheet-grid"]')
    if (!(grid instanceof HTMLElement)) {
      return null
    }
    const typeGpu = document.querySelector('[data-testid="grid-pane-renderer"]')
    const fallback = document.querySelector('[data-testid="grid-pane-renderer-fallback"]')
    const nativeRectLayer = document.querySelector('[data-testid="grid-native-rect-layer"]')
    const nativeTextLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
    const nativeRectCount =
      nativeRectLayer instanceof HTMLElement
        ? Number.parseInt(nativeRectLayer.getAttribute('data-v3-native-rect-count') ?? '0', 10) || 0
        : 0
    const nativeTextRunCount =
      nativeTextLayer instanceof HTMLElement
        ? Number.parseInt(nativeTextLayer.getAttribute('data-v3-native-text-run-count') ?? '0', 10) || 0
        : 0
    const fallbackState: BiligRenderedCanvasState | null =
      fallback instanceof HTMLCanvasElement
        ? {
            authoritativeRenderRevision: fallback.getAttribute('data-v3-authoritative-render-revision'),
            backendStatus: fallback.getAttribute('data-v3-backend-status'),
            currentContentSignature: fallback.getAttribute('data-v3-current-content-signature'),
            currentFillHandleRevision: fallback.getAttribute('data-v3-current-fill-handle-revision'),
            currentSceneEpochSignature: fallback.getAttribute('data-v3-current-scene-epoch-signature'),
            currentSceneOwnershipSignature: fallback.getAttribute('data-v3-current-scene-ownership-signature'),
            currentSelectionRevision: fallback.getAttribute('data-v3-current-selection-revision'),
            currentSemanticMutationRevision: fallback.getAttribute('data-v3-current-semantic-mutation-revision'),
            currentRectCount: Number.parseInt(fallback.getAttribute('data-v3-current-rect-count') ?? '0', 10) || 0,
            currentRectSignature: fallback.getAttribute('data-v3-current-rect-signature'),
            currentTextRunCount: Number.parseInt(fallback.getAttribute('data-v3-current-text-run-count') ?? '0', 10) || 0,
            currentTextSignature: fallback.getAttribute('data-v3-current-text-signature'),
            currentViewportRevision: fallback.getAttribute('data-v3-current-viewport-revision'),
            currentWorkbookRevision: fallback.getAttribute('data-v3-current-workbook-revision'),
            drawText: fallback.getAttribute('data-v3-draw-text') === 'true',
            frameProofStatus: fallback.getAttribute('data-v3-frame-proof-status'),
            frameProofSignature: fallback.getAttribute('data-v3-frame-proof-signature'),
            headerPaneCount: Number.parseInt(fallback.getAttribute('data-v3-header-pane-count') ?? '0', 10) || 0,
            hasPresentedFrame: fallback.getAttribute('data-v3-has-presented-frame') === 'true',
            hasPresentedVisibleFrame: fallback.getAttribute('data-v3-has-presented-visible-frame') === 'true',
            localRenderRevision: fallback.getAttribute('data-v3-local-render-revision'),
            mode: fallback.getAttribute('data-renderer-mode'),
            nativeHeaderPaneCount: Number.parseInt(fallback.getAttribute('data-v3-native-header-pane-count') ?? '0', 10) || 0,
            nativeHeaderTextRunCount: Number.parseInt(fallback.getAttribute('data-v3-native-header-text-run-count') ?? '0', 10) || 0,
            nativeLayerSource: fallback.getAttribute('data-v3-native-layer-source'),
            nativeTilePaneCount: Number.parseInt(fallback.getAttribute('data-v3-native-tile-pane-count') ?? '0', 10) || 0,
            nativeTileTextRunCount: Number.parseInt(fallback.getAttribute('data-v3-native-text-run-count') ?? '0', 10) || 0,
            pixelHeight: fallback.height,
            pixelWidth: fallback.width,
            presentedContentSignature: fallback.getAttribute('data-v3-presented-content-signature'),
            presentedFillHandleRevision: fallback.getAttribute('data-v3-presented-fill-handle-revision'),
            presentedSceneEpochSignature: fallback.getAttribute('data-v3-presented-scene-epoch-signature'),
            presentedSceneOwnershipSignature: fallback.getAttribute('data-v3-presented-scene-ownership-signature'),
            presentedFrameProofSignature: fallback.getAttribute('data-v3-presented-frame-proof-signature'),
            presentedHeaderPaneCount: Number.parseInt(fallback.getAttribute('data-v3-presented-header-pane-count') ?? '0', 10) || 0,
            presentedRectCount: Number.parseInt(fallback.getAttribute('data-v3-presented-rect-count') ?? '0', 10) || 0,
            presentedRectSignature: fallback.getAttribute('data-v3-presented-rect-signature'),
            presentedSelectionRevision: fallback.getAttribute('data-v3-presented-selection-revision'),
            presentedSemanticMutationRevision: fallback.getAttribute('data-v3-presented-semantic-mutation-revision'),
            presentedTextRunCount: Number.parseInt(fallback.getAttribute('data-v3-presented-visible-text-run-count') ?? '0', 10) || 0,
            presentedTextSignature: fallback.getAttribute('data-v3-presented-text-signature'),
            presentedTilePaneCount: Number.parseInt(fallback.getAttribute('data-v3-presented-tile-pane-count') ?? '0', 10) || 0,
            presentedViewportRevision: fallback.getAttribute('data-v3-presented-viewport-revision'),
            presentedWorkbookRevision: fallback.getAttribute('data-v3-presented-workbook-revision'),
            projectedRenderRevision: fallback.getAttribute('data-v3-projected-render-revision'),
            tilePaneCount: Number.parseInt(fallback.getAttribute('data-v3-tile-pane-count') ?? '0', 10) || 0,
            tileSceneRevision: fallback.getAttribute('data-v3-tile-scene-revision'),
            visibleAuthoritativeRenderRevision: fallback.getAttribute('data-v3-visible-authoritative-render-revision'),
            visibleLocalRenderRevision: fallback.getAttribute('data-v3-visible-local-render-revision'),
            visibleProjectedRenderRevision: fallback.getAttribute('data-v3-visible-projected-render-revision'),
            visibleRenderRevision: fallback.getAttribute('data-v3-visible-render-revision'),
          }
        : null
    const typeGpuState: BiligRenderedCanvasState | null =
      typeGpu instanceof HTMLCanvasElement
        ? {
            authoritativeRenderRevision: typeGpu.getAttribute('data-v3-authoritative-render-revision'),
            backendStatus: typeGpu.getAttribute('data-v3-backend-status'),
            currentContentSignature: typeGpu.getAttribute('data-v3-current-content-signature'),
            currentFillHandleRevision: typeGpu.getAttribute('data-v3-current-fill-handle-revision'),
            currentSceneEpochSignature: typeGpu.getAttribute('data-v3-current-scene-epoch-signature'),
            currentSceneOwnershipSignature: typeGpu.getAttribute('data-v3-current-scene-ownership-signature'),
            currentSelectionRevision: typeGpu.getAttribute('data-v3-current-selection-revision'),
            currentSemanticMutationRevision: typeGpu.getAttribute('data-v3-current-semantic-mutation-revision'),
            currentRectCount: Number.parseInt(typeGpu.getAttribute('data-v3-current-rect-count') ?? '0', 10) || 0,
            currentRectSignature: typeGpu.getAttribute('data-v3-current-rect-signature'),
            currentTextRunCount: Number.parseInt(typeGpu.getAttribute('data-v3-current-text-run-count') ?? '0', 10) || 0,
            currentTextSignature: typeGpu.getAttribute('data-v3-current-text-signature'),
            currentViewportRevision: typeGpu.getAttribute('data-v3-current-viewport-revision'),
            currentWorkbookRevision: typeGpu.getAttribute('data-v3-current-workbook-revision'),
            drawText: typeGpu.getAttribute('data-v3-draw-text') === 'true',
            frameProofStatus: typeGpu.getAttribute('data-v3-frame-proof-status'),
            frameProofSignature: typeGpu.getAttribute('data-v3-frame-proof-signature'),
            headerPaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-header-pane-count') ?? '0', 10) || 0,
            hasPresentedFrame: typeGpu.getAttribute('data-v3-has-presented-frame') === 'true',
            hasPresentedVisibleFrame: typeGpu.getAttribute('data-v3-has-presented-visible-frame') === 'true',
            localRenderRevision: typeGpu.getAttribute('data-v3-local-render-revision'),
            mode: typeGpu.getAttribute('data-renderer-mode'),
            nativeHeaderPaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-native-header-pane-count') ?? '0', 10) || 0,
            nativeHeaderTextRunCount: Number.parseInt(typeGpu.getAttribute('data-v3-native-header-text-run-count') ?? '0', 10) || 0,
            nativeLayerSource: typeGpu.getAttribute('data-v3-native-layer-source'),
            nativeTilePaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-native-tile-pane-count') ?? '0', 10) || 0,
            nativeTileTextRunCount: Number.parseInt(typeGpu.getAttribute('data-v3-native-text-run-count') ?? '0', 10) || 0,
            pixelHeight: typeGpu.height,
            pixelWidth: typeGpu.width,
            presentedContentSignature: typeGpu.getAttribute('data-v3-presented-content-signature'),
            presentedFillHandleRevision: typeGpu.getAttribute('data-v3-presented-fill-handle-revision'),
            presentedSceneEpochSignature: typeGpu.getAttribute('data-v3-presented-scene-epoch-signature'),
            presentedSceneOwnershipSignature: typeGpu.getAttribute('data-v3-presented-scene-ownership-signature'),
            presentedFrameProofSignature: typeGpu.getAttribute('data-v3-presented-frame-proof-signature'),
            presentedHeaderPaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-presented-header-pane-count') ?? '0', 10) || 0,
            presentedRectCount: Number.parseInt(typeGpu.getAttribute('data-v3-presented-rect-count') ?? '0', 10) || 0,
            presentedRectSignature: typeGpu.getAttribute('data-v3-presented-rect-signature'),
            presentedSelectionRevision: typeGpu.getAttribute('data-v3-presented-selection-revision'),
            presentedSemanticMutationRevision: typeGpu.getAttribute('data-v3-presented-semantic-mutation-revision'),
            presentedTextRunCount: Number.parseInt(typeGpu.getAttribute('data-v3-presented-visible-text-run-count') ?? '0', 10) || 0,
            presentedTextSignature: typeGpu.getAttribute('data-v3-presented-text-signature'),
            presentedTilePaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-presented-tile-pane-count') ?? '0', 10) || 0,
            presentedViewportRevision: typeGpu.getAttribute('data-v3-presented-viewport-revision'),
            presentedWorkbookRevision: typeGpu.getAttribute('data-v3-presented-workbook-revision'),
            projectedRenderRevision: typeGpu.getAttribute('data-v3-projected-render-revision'),
            tilePaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-tile-pane-count') ?? '0', 10) || 0,
            tileSceneRevision: typeGpu.getAttribute('data-v3-tile-scene-revision'),
            visibleAuthoritativeRenderRevision: typeGpu.getAttribute('data-v3-visible-authoritative-render-revision'),
            visibleLocalRenderRevision: typeGpu.getAttribute('data-v3-visible-local-render-revision'),
            visibleProjectedRenderRevision: typeGpu.getAttribute('data-v3-visible-projected-render-revision'),
            visibleRenderRevision: typeGpu.getAttribute('data-v3-visible-render-revision'),
          }
        : null
    return {
      dpr: Math.max(1, window.devicePixelRatio || 1),
      fallback: fallbackState,
      gridAuthoritativeRenderRevision: grid.getAttribute('data-render-authoritative-revision'),
      gridHeight: Math.max(0, Math.floor(grid.clientHeight)),
      gridLocalRenderRevision: grid.getAttribute('data-render-local-revision'),
      gridProjectedRenderRevision: grid.getAttribute('data-render-projected-revision'),
      gridWidth: Math.max(0, Math.floor(grid.clientWidth)),
      nativeRectCount,
      nativeRectLayerMounted: nativeRectLayer instanceof HTMLElement,
      nativeTextLayerMounted: nativeTextLayer instanceof HTMLElement,
      nativeTextRunCount,
      typeGpu: typeGpuState,
    }
  })
}
