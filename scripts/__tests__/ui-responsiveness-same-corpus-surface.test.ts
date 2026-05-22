import { describe, expect, it } from 'vitest'

import {
  biligRenderedSurfaceReadiness,
  isBiligRenderedSurfaceReady,
  type BiligRenderedSurfaceState,
} from '../ui-responsiveness-same-corpus-surface.ts'

const readyTypeGpuSurface: BiligRenderedSurfaceState = {
  dpr: 2,
  fallback: null,
  gridHeight: 300,
  gridProjectedRenderRevision: 'rev-3',
  gridWidth: 500,
  typeGpu: {
    backendStatus: 'ready',
    frameProofStatus: 'presented',
    headerPaneCount: 1,
    hasPresentedVisibleFrame: true,
    mode: 'typegpu-v3',
    pixelHeight: 600,
    pixelWidth: 1000,
    presentedHeaderPaneCount: 1,
    presentedTilePaneCount: 1,
    projectedRenderRevision: 'rev-3',
    tilePaneCount: 1,
    tileSceneRevision: 'scene-7',
    visibleProjectedRenderRevision: 'rev-3',
    visibleRenderRevision: 'scene-7',
  },
}

describe('same-corpus Bilig rendered surface proof', () => {
  it('requires the TypeGPU renderer without accepting a fallback canvas', () => {
    expect(isBiligRenderedSurfaceReady(readyTypeGpuSurface)).toBe(true)
    expect(
      isBiligRenderedSurfaceReady({
        ...readyTypeGpuSurface,
        fallback: {
          headerPaneCount: 1,
          mode: 'legacy-fallback',
          pixelHeight: 600,
          pixelWidth: 1000,
          tilePaneCount: 1,
        },
      }),
    ).toBe(false)
  })

  it('requires a presented visible TypeGPU frame before capture', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        frameProofStatus: 'pending',
        hasPresentedVisibleFrame: false,
        presentedHeaderPaneCount: 0,
        presentedTilePaneCount: 0,
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toEqual(
      expect.arrayContaining([
        'frame proof is pending',
        'visible frame has not been presented',
        'presented tile/header pane counts are empty',
      ]),
    )
  })

  it('rejects stale projected and visible render revisions', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        projectedRenderRevision: 'rev-2',
        tileSceneRevision: 'scene-6',
        visibleProjectedRenderRevision: 'rev-1',
        visibleRenderRevision: 'scene-5',
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toEqual(
      expect.arrayContaining([
        'TypeGPU projected render revision does not match the grid revision',
        'visible projected render revision does not match the grid revision',
        'visible render revision does not match the tile scene revision',
      ]),
    )
  })

  it('rejects a canvas that does not cover the workbook viewport', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        pixelHeight: 200,
        pixelWidth: 200,
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toContain('TypeGPU canvas backing pixels do not cover the viewport')
  })
})
