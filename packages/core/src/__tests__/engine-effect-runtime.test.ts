import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'

import { runEngineEffect, runEngineEffectPromise } from '../engine/live-effect.js'

describe('engine effect runtime helpers', () => {
  it('unwraps synchronous effect successes and failures', () => {
    expect(runEngineEffect(Effect.succeed(42))).toBe(42)
    expect(() => runEngineEffect(Effect.fail(new Error('sync-failure')))).toThrow('sync-failure')
  })

  it('unwraps asynchronous effect successes and failures', async () => {
    await expect(runEngineEffectPromise(Effect.succeed('ok'))).resolves.toBe('ok')
    await expect(runEngineEffectPromise(Effect.fail(new Error('async-failure')))).rejects.toThrow('async-failure')
  })
})
