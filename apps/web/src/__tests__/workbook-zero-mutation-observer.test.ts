import { describe, expect, it, vi } from 'vitest'
import { waitForZeroMutationObserver, type ZeroMutationObserverFailure } from '../workbook-zero-mutation-observer.js'

describe('workbook zero mutation observer', () => {
  it('releases timed-out observers while still surfacing late server failures', async () => {
    vi.useFakeTimers()
    let resolveObserver: ((value: unknown) => void) | null = null
    const observer = new Promise<unknown>((resolve) => {
      resolveObserver = resolve
    })
    const lateFailures: ZeroMutationObserverFailure[] = []

    const waitPromise = waitForZeroMutationObserver({
      observer,
      timeoutMs: 50,
      onLateFailure: (failure) => {
        lateFailures.push(failure)
      },
    })

    await vi.advanceTimersByTimeAsync(50)
    await expect(waitPromise).resolves.toEqual({ ok: true })

    resolveObserver?.({
      type: 'error',
      error: {
        type: 'app',
        message: 'server rejected mutation',
      },
    })
    await vi.runAllTimersAsync()

    expect(lateFailures).toHaveLength(1)
    expect(lateFailures[0]?.retryable).toBe(false)
    expect(lateFailures[0]?.error.message).toBe('server rejected mutation')
  })
})
