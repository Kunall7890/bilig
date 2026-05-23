import { describe, expect, it } from 'vitest'

import {
  biligRenderedSurfaceReadiness,
  isBiligRenderedSurfaceReady,
  type BiligRenderedSurfaceState,
} from '../ui-responsiveness-same-corpus-surface.ts'

const readyTypeGpuSurface: BiligRenderedSurfaceState = {
  dpr: 2,
  fallback: null,
  gridAuthoritativeRenderRevision: 'rev-3',
  gridHeight: 300,
  gridLocalRenderRevision: 'rev-local-2',
  gridProjectedRenderRevision: 'rev-3',
  gridWidth: 500,
  typeGpu: {
    authoritativeRenderRevision: 'rev-3',
    backendStatus: 'ready',
    currentContentSignature: 'content:current',
    currentSceneEpochSignature: 'epoch:current',
    currentSceneOwnershipSignature: 'scene:current',
    currentRectCount: 88,
    currentRectSignature: 'rect:current',
    currentTextRunCount: 12,
    currentTextSignature: 'text:current',
    frameProofStatus: 'presented',
    frameProofSignature: 'frame-current',
    headerPaneCount: 1,
    hasPresentedFrame: true,
    hasPresentedVisibleFrame: true,
    localRenderRevision: 'rev-local-2',
    mode: 'typegpu-v3',
    pixelHeight: 600,
    pixelWidth: 1000,
    presentedContentSignature: 'content:current',
    presentedSceneEpochSignature: 'epoch:current',
    presentedSceneOwnershipSignature: 'scene:current',
    presentedFrameProofSignature: 'frame-current',
    presentedHeaderPaneCount: 1,
    presentedRectCount: 88,
    presentedRectSignature: 'rect:current',
    presentedTextRunCount: 12,
    presentedTextSignature: 'text:current',
    presentedTilePaneCount: 1,
    projectedRenderRevision: 'rev-3',
    tilePaneCount: 1,
    tileSceneRevision: 'scene-7',
    visibleAuthoritativeRenderRevision: 'rev-3',
    visibleLocalRenderRevision: 'rev-local-2',
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
        frameProofSignature: 'frame-next',
        hasPresentedFrame: false,
        hasPresentedVisibleFrame: false,
        presentedFrameProofSignature: 'frame-current',
        presentedHeaderPaneCount: 0,
        presentedTilePaneCount: 0,
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toEqual(
      expect.arrayContaining([
        'frame proof is pending',
        'current frame signature has not been presented',
        'presented frame proof signature does not match current frame',
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
        authoritativeRenderRevision: 'rev-2',
        localRenderRevision: 'rev-local-1',
        projectedRenderRevision: 'rev-2',
        tileSceneRevision: 'scene-6',
        visibleAuthoritativeRenderRevision: 'rev-2',
        visibleLocalRenderRevision: 'rev-local-1',
        visibleProjectedRenderRevision: 'rev-1',
        visibleRenderRevision: 'scene-5',
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toEqual(
      expect.arrayContaining([
        'TypeGPU authoritative render revision does not match the grid revision',
        'visible authoritative render revision does not match the grid revision',
        'TypeGPU local render revision does not match the grid revision',
        'visible local render revision does not match the grid revision',
        'TypeGPU projected render revision does not match the grid revision',
        'visible projected render revision does not match the grid revision',
        'visible render revision does not match the tile scene revision',
      ]),
    )
  })

  it('rejects missing or stale presented-frame lineage', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        frameProofSignature: 'frame-next',
        hasPresentedFrame: true,
        presentedFrameProofSignature: 'frame-current',
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toContain('presented frame proof signature does not match current frame')
  })

  it('rejects stale visible-scene ownership even when payload signatures match', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        presentedSceneOwnershipSignature: 'scene:stale',
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toContain('presented visible-scene ownership does not match current scene')
  })

  it('rejects stale visible-scene epochs even when ownership payload signatures match', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        presentedSceneEpochSignature: 'epoch:stale',
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toContain('presented visible-scene epoch does not match current authoritative scene')
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

  it('rejects partially presented pane coverage', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        presentedHeaderPaneCount: 1,
        presentedTilePaneCount: 2,
        tilePaneCount: 3,
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toContain('presented tile/header pane counts do not cover the current visible panes')
  })

  it('rejects stale visible text and rect payloads even when the frame signature is presented', () => {
    const readiness = biligRenderedSurfaceReadiness({
      ...readyTypeGpuSurface,
      typeGpu: {
        ...readyTypeGpuSurface.typeGpu!,
        presentedContentSignature: 'content:stale',
        presentedRectCount: 77,
        presentedRectSignature: 'rect:stale',
        presentedTextRunCount: 11,
        presentedTextSignature: 'text:stale',
      },
    })

    expect(readiness.ready).toBe(false)
    expect(readiness.gaps).toEqual(
      expect.arrayContaining([
        'presented visible content signature does not match current tiles',
        'presented visible text signature does not match current tiles',
        'presented visible rect signature does not match current tiles',
        'presented visible text run count does not match current tiles',
        'presented visible rect count does not match current tiles',
      ]),
    )
  })
})
