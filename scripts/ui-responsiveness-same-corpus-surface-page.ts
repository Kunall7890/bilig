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
    const fallbackState: BiligRenderedCanvasState | null =
      fallback instanceof HTMLCanvasElement
        ? {
            authoritativeRenderRevision: fallback.getAttribute('data-v3-authoritative-render-revision'),
            backendStatus: fallback.getAttribute('data-v3-backend-status'),
            frameProofStatus: fallback.getAttribute('data-v3-frame-proof-status'),
            headerPaneCount: Number.parseInt(fallback.getAttribute('data-v3-header-pane-count') ?? '0', 10) || 0,
            hasPresentedVisibleFrame: fallback.getAttribute('data-v3-has-presented-visible-frame') === 'true',
            localRenderRevision: fallback.getAttribute('data-v3-local-render-revision'),
            mode: fallback.getAttribute('data-renderer-mode'),
            pixelHeight: fallback.height,
            pixelWidth: fallback.width,
            presentedHeaderPaneCount: Number.parseInt(fallback.getAttribute('data-v3-presented-header-pane-count') ?? '0', 10) || 0,
            presentedTilePaneCount: Number.parseInt(fallback.getAttribute('data-v3-presented-tile-pane-count') ?? '0', 10) || 0,
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
            frameProofStatus: typeGpu.getAttribute('data-v3-frame-proof-status'),
            headerPaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-header-pane-count') ?? '0', 10) || 0,
            hasPresentedVisibleFrame: typeGpu.getAttribute('data-v3-has-presented-visible-frame') === 'true',
            localRenderRevision: typeGpu.getAttribute('data-v3-local-render-revision'),
            mode: typeGpu.getAttribute('data-renderer-mode'),
            pixelHeight: typeGpu.height,
            pixelWidth: typeGpu.width,
            presentedHeaderPaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-presented-header-pane-count') ?? '0', 10) || 0,
            presentedTilePaneCount: Number.parseInt(typeGpu.getAttribute('data-v3-presented-tile-pane-count') ?? '0', 10) || 0,
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
      typeGpu: typeGpuState,
    }
  })
}
