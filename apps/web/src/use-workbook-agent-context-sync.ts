import { useCallback, useRef, type MutableRefObject } from 'react'
import type { WorkbookAgentThreadSnapshot, WorkbookAgentUiContext } from '@bilig/contracts'
import { logDebug } from './runtime-logger.js'
import { stringifyWorkbookAgentContextSyncKey } from './workbook-agent-pane-helpers.js'

const AGENT_CONTEXT_SYNC_DEBOUNCE_MS = 150
const AGENT_CONTEXT_SYNC_MIN_INTERVAL_MS = 750
const AGENT_CONTEXT_SYNC_FAILURE_RETRY_INITIAL_MS = 2_000
const AGENT_CONTEXT_SYNC_FAILURE_RETRY_MAX_MS = 15_000

interface WorkbookAgentContextSyncClient {
  syncThreadContext(threadId: string, context: WorkbookAgentUiContext): Promise<void>
}

interface WorkbookAgentContextSyncSession {
  readonly threadId: string
}

function stringifyWorkbookAgentImmediateContextKey(context: WorkbookAgentUiContext): string {
  return JSON.stringify({
    selection: context.selection,
    viewport: context.viewport,
  })
}

export function useWorkbookAgentContextSync(input: {
  readonly client: WorkbookAgentContextSyncClient
  readonly documentId: string
  readonly enabled: boolean
  readonly getContextRef: MutableRefObject<() => WorkbookAgentUiContext>
  readonly sessionRef: MutableRefObject<WorkbookAgentContextSyncSession | null>
  readonly snapshot: WorkbookAgentThreadSnapshot | null
}) {
  const lastContextKeyRef = useRef<string>('')
  const lastImmediateContextKeyRef = useRef<string>('')
  const hasSyncedContextRef = useRef(false)
  const contextSyncInFlightRef = useRef(false)
  const contextSyncGenerationRef = useRef(0)
  const contextSyncFailureCountRef = useRef(0)
  const lastContextSyncAtRef = useRef(0)
  const nextContextSyncRetryAtRef = useRef(0)
  const pendingContextSyncTimeoutRef = useRef<number | null>(null)
  const pendingContextSyncRef = useRef<{
    readonly context: WorkbookAgentUiContext
    readonly generation: number
    readonly key: string
    readonly immediateKey: string
    readonly threadId: string
  } | null>(null)

  const clearPendingContextSync = useCallback(() => {
    contextSyncGenerationRef.current += 1
    const pendingTimeout = pendingContextSyncTimeoutRef.current
    if (pendingTimeout !== null) {
      window.clearTimeout(pendingTimeout)
      pendingContextSyncTimeoutRef.current = null
    }
    pendingContextSyncRef.current = null
  }, [])

  const resetContextSync = useCallback(() => {
    clearPendingContextSync()
    lastContextKeyRef.current = ''
    lastImmediateContextKeyRef.current = ''
    hasSyncedContextRef.current = false
    contextSyncFailureCountRef.current = 0
    lastContextSyncAtRef.current = 0
    nextContextSyncRetryAtRef.current = 0
  }, [clearPendingContextSync])

  const resolveContextSyncDelay = useCallback((baseDelayMs: number) => {
    const retryDelayMs = Math.max(0, nextContextSyncRetryAtRef.current - window.performance.now())
    return Math.max(baseDelayMs, retryDelayMs)
  }, [])

  const isPendingContextSyncCurrent = useCallback(
    (pending: NonNullable<typeof pendingContextSyncRef.current>) =>
      pending.generation === contextSyncGenerationRef.current && input.sessionRef.current?.threadId === pending.threadId,
    [input.sessionRef],
  )

  const flushPendingContextSync = useCallback(() => {
    const pending = pendingContextSyncRef.current
    pendingContextSyncTimeoutRef.current = null
    if (!pending || !isPendingContextSyncCurrent(pending) || lastContextKeyRef.current === pending.key) {
      pendingContextSyncRef.current = null
      return
    }
    if (contextSyncInFlightRef.current) {
      return
    }
    pendingContextSyncRef.current = null
    lastContextSyncAtRef.current = window.performance.now()
    contextSyncInFlightRef.current = true
    void (async () => {
      try {
        await input.client.syncThreadContext(pending.threadId, pending.context)
        if (isPendingContextSyncCurrent(pending)) {
          lastContextKeyRef.current = pending.key
          lastImmediateContextKeyRef.current = pending.immediateKey
          hasSyncedContextRef.current = true
          contextSyncFailureCountRef.current = 0
          nextContextSyncRetryAtRef.current = 0
        }
      } catch (syncError) {
        logDebug('Failed to sync agent context update', { documentId: input.documentId, error: syncError })
        if (isPendingContextSyncCurrent(pending)) {
          const failureCount = contextSyncFailureCountRef.current + 1
          contextSyncFailureCountRef.current = failureCount
          const retryDelayMs = Math.min(
            AGENT_CONTEXT_SYNC_FAILURE_RETRY_MAX_MS,
            AGENT_CONTEXT_SYNC_FAILURE_RETRY_INITIAL_MS * 2 ** Math.max(0, failureCount - 1),
          )
          nextContextSyncRetryAtRef.current = window.performance.now() + retryDelayMs
          pendingContextSyncRef.current ??= pending
        }
      } finally {
        contextSyncInFlightRef.current = false
        if (pendingContextSyncRef.current !== null && pendingContextSyncTimeoutRef.current === null) {
          const elapsedSinceLastSync = window.performance.now() - lastContextSyncAtRef.current
          const delayMs = resolveContextSyncDelay(
            Math.max(AGENT_CONTEXT_SYNC_DEBOUNCE_MS, AGENT_CONTEXT_SYNC_MIN_INTERVAL_MS - elapsedSinceLastSync),
          )
          pendingContextSyncTimeoutRef.current = window.setTimeout(flushPendingContextSync, delayMs)
        }
      }
    })()
  }, [input.client, input.documentId, isPendingContextSyncCurrent, resolveContextSyncDelay])

  const scheduleContextSync = useCallback(() => {
    const activeSession = input.sessionRef.current
    if (!input.enabled || !input.snapshot || !activeSession) {
      clearPendingContextSync()
      return
    }
    const nextContext = input.getContextRef.current()
    const nextContextKey = `${activeSession.threadId}:${stringifyWorkbookAgentContextSyncKey(nextContext)}`
    const nextImmediateContextKey = `${activeSession.threadId}:${stringifyWorkbookAgentImmediateContextKey(nextContext)}`
    if (lastContextKeyRef.current === nextContextKey) {
      return
    }
    pendingContextSyncRef.current = {
      context: nextContext,
      generation: contextSyncGenerationRef.current,
      key: nextContextKey,
      immediateKey: nextImmediateContextKey,
      threadId: activeSession.threadId,
    }
    const shouldPrioritizeSync = !hasSyncedContextRef.current || lastImmediateContextKeyRef.current !== nextImmediateContextKey
    const elapsedSinceLastSync = window.performance.now() - lastContextSyncAtRef.current
    const delayMs = resolveContextSyncDelay(
      shouldPrioritizeSync
        ? AGENT_CONTEXT_SYNC_DEBOUNCE_MS
        : Math.max(AGENT_CONTEXT_SYNC_DEBOUNCE_MS, AGENT_CONTEXT_SYNC_MIN_INTERVAL_MS - elapsedSinceLastSync),
    )
    if (pendingContextSyncTimeoutRef.current !== null) {
      window.clearTimeout(pendingContextSyncTimeoutRef.current)
    }
    pendingContextSyncTimeoutRef.current = window.setTimeout(flushPendingContextSync, delayMs)
  }, [
    clearPendingContextSync,
    flushPendingContextSync,
    input.enabled,
    input.getContextRef,
    input.sessionRef,
    input.snapshot,
    resolveContextSyncDelay,
  ])

  return {
    clearPendingContextSync,
    resetContextSync,
    scheduleContextSync,
  }
}
