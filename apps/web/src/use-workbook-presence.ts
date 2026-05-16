import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { mutators, queries } from '@bilig/zero-sync'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import {
  normalizeWorkbookPresenceRows,
  selectActiveWorkbookCollaborators,
  WORKBOOK_PRESENCE_HEARTBEAT_MS,
  WORKBOOK_PRESENCE_SELECTION_PUBLISH_MS,
  WORKBOOK_PRESENCE_STALE_TICK_MS,
  type WorkbookCollaboratorPresence,
} from './workbook-presence-model.js'
import { logDebug } from './runtime-logger.js'

interface ZeroLiveView<T> {
  readonly data: T
  addListener(listener: (value: T) => void): () => void
  destroy(): void
}

interface ZeroPresenceSource {
  materialize(query: unknown): unknown
  mutate(mutation: unknown): unknown
}

type UpdatePresenceArgs = Parameters<typeof mutators.workbook.updatePresence>[0] & {
  presenceClientId?: string
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isZeroLiveView<T>(value: unknown): value is ZeroLiveView<T> {
  return isRecord(value) && 'data' in value && typeof value['addListener'] === 'function' && typeof value['destroy'] === 'function'
}

function observeZeroMutationResult(result: unknown): void {
  if (!isRecord(result)) {
    return
  }
  const observer = result['server'] ?? result['client']
  if (!(observer instanceof Promise)) {
    return
  }
  void (async () => {
    try {
      await observer
    } catch (error) {
      logDebug('Failed to persist workbook presence mutation', error)
    }
  })()
}

export function useWorkbookPresence(input: {
  readonly documentId: string
  readonly currentUserId: string
  readonly currentPresenceClientId: string
  readonly sessionId: string
  readonly selection: WorkerRuntimeSelection
  readonly sheetNames: readonly string[]
  readonly zero: ZeroPresenceSource
  readonly enabled: boolean
}): readonly WorkbookCollaboratorPresence[] {
  const { currentPresenceClientId, currentUserId, documentId, enabled, selection, sessionId, sheetNames, zero } = input
  const [presenceRows, setPresenceRows] = useState([] as readonly ReturnType<typeof normalizeWorkbookPresenceRows>[number][])
  const [now, setNow] = useState(() => Date.now())
  const latestSelectionRef = useRef(selection)
  const lastPresencePublishedAtRef = useRef(0)
  const pendingPresencePublishTimeoutRef = useRef<number | null>(null)

  latestSelectionRef.current = selection

  const clearPendingPresencePublish = useCallback(() => {
    if (pendingPresencePublishTimeoutRef.current === null) {
      return
    }
    window.clearTimeout(pendingPresencePublishTimeoutRef.current)
    pendingPresencePublishTimeoutRef.current = null
  }, [])

  const publishPresence = useCallback(
    (publishedAt = Date.now()) => {
      if (!enabled) {
        return
      }
      const presenceArgs: UpdatePresenceArgs = {
        documentId,
        sessionId,
        presenceClientId: currentPresenceClientId,
        sheetName: latestSelectionRef.current.sheetName,
        address: latestSelectionRef.current.address,
        selection: latestSelectionRef.current,
      }
      observeZeroMutationResult(zero.mutate(mutators.workbook.updatePresence(presenceArgs)))
      lastPresencePublishedAtRef.current = publishedAt
    },
    [currentPresenceClientId, documentId, enabled, sessionId, zero],
  )

  const publishPresenceNow = useCallback(() => {
    clearPendingPresencePublish()
    publishPresence(Date.now())
  }, [clearPendingPresencePublish, publishPresence])

  useEffect(() => {
    lastPresencePublishedAtRef.current = 0
    clearPendingPresencePublish()
  }, [clearPendingPresencePublish, currentPresenceClientId, documentId, sessionId])

  useEffect(() => {
    if (!enabled) {
      setPresenceRows([])
      clearPendingPresencePublish()
      return
    }
    const view = zero.materialize(queries.presenceCoarse.byWorkbook({ documentId }))
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error('Zero workbook presence query returned an invalid live view')
    }
    const publishRows = (value: unknown) => {
      setPresenceRows(normalizeWorkbookPresenceRows(value))
      setNow(Date.now())
    }
    publishRows(view.data)
    const cleanup = view.addListener((value) => {
      publishRows(value)
    })
    return () => {
      cleanup()
      view.destroy()
    }
  }, [clearPendingPresencePublish, documentId, enabled, zero])

  useEffect(() => {
    if (!enabled) {
      clearPendingPresencePublish()
      return
    }
    const intervalId = window.setInterval(() => {
      publishPresenceNow()
    }, WORKBOOK_PRESENCE_HEARTBEAT_MS)
    return () => {
      window.clearInterval(intervalId)
    }
  }, [clearPendingPresencePublish, enabled, publishPresenceNow])

  useEffect(() => {
    if (!enabled) {
      clearPendingPresencePublish()
      return
    }
    const publishedAt = lastPresencePublishedAtRef.current
    const currentTime = Date.now()
    const elapsedMs = currentTime - publishedAt
    if (publishedAt === 0 || elapsedMs >= WORKBOOK_PRESENCE_SELECTION_PUBLISH_MS) {
      publishPresenceNow()
      return
    }
    clearPendingPresencePublish()
    pendingPresencePublishTimeoutRef.current = window.setTimeout(() => {
      pendingPresencePublishTimeoutRef.current = null
      publishPresence(Date.now())
    }, WORKBOOK_PRESENCE_SELECTION_PUBLISH_MS - elapsedMs)
    return clearPendingPresencePublish
  }, [clearPendingPresencePublish, enabled, publishPresence, publishPresenceNow, selection.address, selection.sheetName])

  useEffect(() => {
    if (!enabled) {
      return
    }
    const intervalId = window.setInterval(() => {
      setNow(Date.now())
    }, WORKBOOK_PRESENCE_STALE_TICK_MS)
    return () => {
      window.clearInterval(intervalId)
    }
  }, [enabled])

  return useMemo(
    () =>
      selectActiveWorkbookCollaborators({
        rows: presenceRows,
        currentPresenceClientId,
        currentUserId,
        currentSessionId: sessionId,
        knownSheetNames: sheetNames,
        now,
      }),
    [currentPresenceClientId, currentUserId, now, presenceRows, sessionId, sheetNames],
  )
}
