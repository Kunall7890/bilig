import { describe, expect, it } from 'vitest'

import { resolveKeepWorkpaperSmokeStage } from '../workpaper-external-smoke-config.ts'

describe('workpaper external smoke config', () => {
  it('defaults to cleaning up the smoke stage', () => {
    expect(resolveKeepWorkpaperSmokeStage({})).toBe(false)
    expect(resolveKeepWorkpaperSmokeStage({ KEEP_WORKPAPER_SMOKE_STAGE: '' })).toBe(false)
  })

  it('accepts explicit boolean values for preserving the smoke stage', () => {
    expect(resolveKeepWorkpaperSmokeStage({ KEEP_WORKPAPER_SMOKE_STAGE: '1' })).toBe(true)
    expect(resolveKeepWorkpaperSmokeStage({ KEEP_WORKPAPER_SMOKE_STAGE: 'true' })).toBe(true)
    expect(resolveKeepWorkpaperSmokeStage({ KEEP_WORKPAPER_SMOKE_STAGE: '0' })).toBe(false)
    expect(resolveKeepWorkpaperSmokeStage({ KEEP_WORKPAPER_SMOKE_STAGE: 'false' })).toBe(false)
  })

  it('rejects malformed preserve-stage flags', () => {
    expect(() => resolveKeepWorkpaperSmokeStage({ KEEP_WORKPAPER_SMOKE_STAGE: 'yes' })).toThrow(
      'KEEP_WORKPAPER_SMOKE_STAGE must be "1", "true", "0", or "false" when set, got yes',
    )
  })
})
