import { isMutationErrorResult, toErrorMessage } from './worker-workbook-app-model.js'

export const ZERO_MUTATION_OBSERVER_SETTLE_TIMEOUT_MS = 2_000

export interface ZeroMutationObserverFailure {
  readonly retryable: boolean
  readonly error: Error
}

export type ZeroMutationObserverOutcome = { readonly ok: true } | ({ readonly ok: false } & ZeroMutationObserverFailure)

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function observeZeroMutationResult(result: unknown): Promise<unknown> | null {
  if (!isRecord(result)) {
    return null
  }
  const observer = result['server'] ?? result['client']
  return observer instanceof Promise ? observer : null
}

export async function resolveZeroMutationObserverOutcome(observer: Promise<unknown>): Promise<ZeroMutationObserverOutcome> {
  try {
    const remoteResult = await observer
    if (!isMutationErrorResult(remoteResult)) {
      return { ok: true }
    }
    const details =
      remoteResult.error.type === 'app' && remoteResult.error.details !== undefined
        ? ` (${JSON.stringify(remoteResult.error.details)})`
        : ''
    return {
      ok: false,
      retryable: remoteResult.error.type === 'zero',
      error: new Error(`${remoteResult.error.message}${details}`),
    }
  } catch (error) {
    return {
      ok: false,
      retryable: true,
      error: error instanceof Error ? error : new Error(toErrorMessage(error)),
    }
  }
}

export async function waitForZeroMutationObserver(input: {
  readonly observer: Promise<unknown>
  readonly timeoutMs?: number
  readonly onLateFailure?: (failure: ZeroMutationObserverFailure) => void
}): Promise<ZeroMutationObserverOutcome> {
  let timeoutId: ReturnType<typeof setTimeout> | null = null
  const outcomePromise = resolveZeroMutationObserverOutcome(input.observer)
  const timeoutPromise = new Promise<'timeout'>((resolve) => {
    timeoutId = setTimeout(() => resolve('timeout'), input.timeoutMs ?? ZERO_MUTATION_OBSERVER_SETTLE_TIMEOUT_MS)
  })
  const firstResult = await Promise.race([outcomePromise, timeoutPromise])
  if (firstResult === 'timeout') {
    void (async () => {
      try {
        const outcome = await outcomePromise
        if (!outcome.ok) {
          input.onLateFailure?.({
            retryable: outcome.retryable,
            error: outcome.error,
          })
        }
      } catch {
        // resolveZeroMutationObserverOutcome converts observer failures into outcomes.
      }
    })()
    return { ok: true }
  }
  if (timeoutId !== null) {
    clearTimeout(timeoutId)
  }
  return firstResult
}
