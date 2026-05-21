// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import {
  WORKBOOK_ATLAS_TEXT_RENDERING,
  configureTextContext,
  createGlyphAtlas,
  resolveGlyphAtlasScale,
  type TextContextConfigurationTarget,
} from '../renderer-v3/typegpu-atlas-manager.js'

describe('glyph-atlas', () => {
  it('configures atlas text for small spreadsheet legibility', () => {
    const context: TextContextConfigurationTarget = {
      fontKerning: 'none',
      imageSmoothingEnabled: false,
      imageSmoothingQuality: 'low',
      textBaseline: 'top',
      textRendering: 'geometricPrecision',
    }

    configureTextContext(context)

    expect(context.textBaseline).toBe('alphabetic')
    expect(context.imageSmoothingEnabled).toBe(true)
    expect(context.imageSmoothingQuality).toBe('high')
    expect(context.fontKerning).toBe('normal')
    expect(context.textRendering).toBe(WORKBOOK_ATLAS_TEXT_RENDERING)
  })

  it('resolves atlas scale from the active device pixel ratio bucket', () => {
    expect(resolveGlyphAtlasScale(0)).toBe(2)
    expect(resolveGlyphAtlasScale(1)).toBe(2)
    expect(resolveGlyphAtlasScale(1.25)).toBe(2)
    expect(resolveGlyphAtlasScale(2)).toBe(2)
    expect(resolveGlyphAtlasScale(99)).toBe(4)
  })

  it('returns stable glyph keys for repeated runs', () => {
    const atlas = createGlyphAtlas()
    const first = atlas.intern('400 11px Geist', 'A')
    const second = atlas.intern('400 11px Geist', 'A')

    expect(second.key).toBe(first.key)
    expect(second.x).toBe(first.x)
    expect(second.y).toBe(first.y)
  })

  it('tracks atlas version after adding glyphs', () => {
    const atlas = createGlyphAtlas()
    expect(atlas.getVersion()).toBe(0)
    expect(atlas.getGlyphGeometryVersion()).toBe(0)
    atlas.intern('400 11px Geist', 'A')
    expect(atlas.getVersion()).toBeGreaterThan(0)
    expect(atlas.getGlyphGeometryVersion()).toBe(0)
  })

  it('bumps glyph geometry version only when atlas texture growth remaps UVs', () => {
    const originalOffscreenCanvas = globalThis.OffscreenCanvas
    Object.defineProperty(globalThis, 'OffscreenCanvas', {
      configurable: true,
      value: class {
        height: number
        width: number
        constructor(width: number, height: number) {
          this.width = width
          this.height = height
        }
        getContext() {
          return null
        }
      },
    })
    try {
      const atlas = createGlyphAtlas({ initialHeight: 512, initialWidth: 512 })
      let glyph = 0
      const firstGeometryVersion = atlas.getGlyphGeometryVersion()

      while (atlas.getGlyphGeometryVersion() === firstGeometryVersion && glyph < 4000) {
        atlas.intern('400 48px Geist', `glyph-${glyph}`)
        glyph += 1
      }

      expect(atlas.getGlyphGeometryVersion()).toBeGreaterThan(firstGeometryVersion)
    } finally {
      Object.defineProperty(globalThis, 'OffscreenCanvas', {
        configurable: true,
        value: originalOffscreenCanvas,
      })
    }
  })

  it('tracks dirty atlas pages for V3 glyph inserts', () => {
    const atlas = createGlyphAtlas()

    const entry = atlas.intern('400 11px Geist', 'A')

    const stats = atlas.getDirtyPageStats()
    expect(atlas.getTextAtlasPagesSeq()).toBeGreaterThan(0)
    expect(stats.dirtyPageCount).toBeGreaterThan(0)
    expect(stats.dirtyUploadBytes).toBeGreaterThan(0)
    expect(atlas.getTextAtlasPagesStats()).toMatchObject({
      dirtyPageCount: stats.dirtyPageCount,
      glyphCount: 1,
    })
    expect(stats.dirtyUploadBytes).toBeLessThanOrEqual(2 * 32 * 32 * 4)
    expect(stats.dirtyUploadBytes).toBeLessThan(64 * 64 * 4)
    expect(atlas.resolveGlyphRecord(entry.glyphId)).toMatchObject({
      glyphId: entry.glyphId,
      pageId: entry.pageId,
    })

    const pages = atlas.drainDirtyPages()
    expect(pages.length).toBe(stats.dirtyPageCount)
    expect(pages.every((page) => page.byteSize === page.width * page.height * 4)).toBe(true)
    expect(atlas.getDirtyPageStats().dirtyPageCount).toBe(0)
    expect(atlas.getTextAtlasPagesStats().dirtyPageCount).toBe(0)
  })

  it('assigns stable glyph identities without ref-counting repeated reads as new glyph registrations', () => {
    const atlas = createGlyphAtlas()

    const first = atlas.intern('400 11px Geist', 'A')
    atlas.drainDirtyPages()
    const second = atlas.intern('400 11px Geist', 'A')

    expect(second.glyphId).toBe(first.glyphId)
    expect(second.pageId).toBe(first.pageId)
    expect(atlas.getTextAtlasPagesStats()).toMatchObject({
      dirtyPageCount: 0,
      glyphCount: 1,
    })
  })

  it('rekeys and invalidates glyph geometry when the surface DPR bucket changes', () => {
    const atlas = createGlyphAtlas({ initialHeight: 512, initialWidth: 512, scale: 1 })
    const first = atlas.intern('400 11px Geist', 'A')
    atlas.drainDirtyPages()
    const geometryVersion = atlas.getGlyphGeometryVersion()
    const version = atlas.getVersion()

    expect(atlas.getScale()).toBe(2)
    expect(atlas.getSize()).toEqual({ height: 1024, width: 1024 })
    expect(atlas.setScale(1)).toBe(false)
    expect(atlas.getGlyphGeometryVersion()).toBe(geometryVersion)

    expect(atlas.setScale(3)).toBe(true)
    expect(atlas.getScale()).toBe(3)
    expect(atlas.getSize()).toEqual({ height: 1536, width: 1536 })
    expect(atlas.getVersion()).toBeGreaterThan(version)
    expect(atlas.getGlyphGeometryVersion()).toBeGreaterThan(geometryVersion)
    expect(atlas.getTextAtlasPagesStats()).toMatchObject({ glyphCount: 0 })

    const second = atlas.intern('400 11px Geist', 'A')
    expect(second.key).not.toBe(first.key)
    expect(second.key.startsWith('3:')).toBe(true)
  })
})
