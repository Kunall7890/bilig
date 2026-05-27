import { describe, expect, it } from 'vitest'

import { sameCorpusVisibleCellInteriorClip } from '../ui-responsiveness-same-corpus-mutation-readback.ts'

describe('same-corpus mutation visible-cell readback', () => {
  it('samples the middle of a selected cell instead of border and fill-handle pixels', () => {
    expect(sameCorpusVisibleCellInteriorClip({ height: 40, width: 100, x: 10, y: 20 }, { height: 200, width: 300 })).toEqual({
      height: 24,
      width: 60,
      x: 30,
      y: 28,
    })
  })

  it('keeps small cells measurable while still avoiding the outer edge', () => {
    expect(sameCorpusVisibleCellInteriorClip({ height: 8, width: 12, x: 5, y: 7 }, { height: 50, width: 50 })).toEqual({
      height: 2,
      width: 6,
      x: 8,
      y: 10,
    })
  })

  it('clips the sampling window to the viewport', () => {
    expect(sameCorpusVisibleCellInteriorClip({ height: 30, width: 60, x: 260, y: 170 }, { height: 190, width: 300 })).toEqual({
      height: 14,
      width: 28,
      x: 272,
      y: 176,
    })
  })

  it('rejects cells whose interior crop is outside the viewport', () => {
    expect(sameCorpusVisibleCellInteriorClip({ height: 20, width: 20, x: 40, y: 40 }, { height: 10, width: 10 })).toBeNull()
  })
})
