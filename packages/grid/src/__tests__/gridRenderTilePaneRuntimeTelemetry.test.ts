import { describe, expect, it, vi } from 'vitest'
import { GridRenderTilePaneRuntime } from '../runtime/gridRenderTilePaneRuntime.js'

describe('GridRenderTilePaneRuntime telemetry', () => {
  it('reports tile cache lookup telemetry from readiness snapshots', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const scrollPerf = {
      readiness: null as null | {
        exactHits: number
        misses: number
        staleHits: number
        visibleDirtyTiles: number
        warmDirtyTiles: number
      },
      staleLookups: null as null | { hits: number; lookups: number; scannedEntries: number },
      visibleMarks: 0,
      noteRendererTileReadiness(input: {
        exactHits: number
        misses: number
        staleHits: number
        visibleDirtyTiles: number
        warmDirtyTiles: number
      }): void {
        this.readiness = input
      },
      noteTypeGpuTileCacheStaleLookups(input: { hits: number; lookups: number; scannedEntries: number }): void {
        this.staleLookups = input
      },
      noteTypeGpuTileCacheVisibleMark(count: number): void {
        this.visibleMarks += count
      },
    }
    vi.stubGlobal('window', { __biligScrollPerf: scrollPerf })
    try {
      runtime.noteTileReadiness({
        exactHits: [1],
        misses: [],
        staleHits: [2],
        staleLookupCount: 2,
        staleLookupScannedEntries: 7,
        visibleDirtyTileKeys: [],
        visibleMarkedTiles: 3,
        warmDirtyTileKeys: [],
      })

      expect(scrollPerf.readiness).toEqual({
        exactHits: 1,
        misses: 0,
        staleHits: 1,
        visibleDirtyTiles: 0,
        warmDirtyTiles: 0,
      })
      expect(scrollPerf.staleLookups).toEqual({ hits: 1, lookups: 2, scannedEntries: 7 })
      expect(scrollPerf.visibleMarks).toBe(3)
    } finally {
      vi.unstubAllGlobals()
    }
  })
})
