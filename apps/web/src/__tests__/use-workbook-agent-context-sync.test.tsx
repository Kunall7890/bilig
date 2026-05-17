// @vitest-environment jsdom
import { act, useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import type { WorkbookAgentThreadSnapshot, WorkbookAgentUiContext } from '@bilig/contracts'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { useWorkbookAgentContextSync } from '../use-workbook-agent-context-sync.js'

function createContext(address: string): WorkbookAgentUiContext {
  return {
    selection: {
      address,
      sheetName: 'Sheet1',
    },
    viewport: {
      colEnd: 10,
      colStart: 0,
      rowEnd: 20,
      rowStart: 0,
    },
  }
}

function createSnapshot(): WorkbookAgentThreadSnapshot {
  return {
    activeTurnId: null,
    context: createContext('A1'),
    documentId: 'doc-1',
    entries: [],
    executionPolicy: 'autoApplyAll',
    executionRecords: [],
    lastError: null,
    reviewQueueItems: [],
    scope: 'private',
    status: 'idle',
    threadId: 'thr-1',
    workflowRuns: [],
  }
}

afterEach(() => {
  document.body.innerHTML = ''
  vi.useRealTimers()
})

describe('useWorkbookAgentContextSync', () => {
  it('backs off failed context sync retries while preserving the latest pending context', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()

    const syncThreadContext = vi.fn(async () => {
      if (syncThreadContext.mock.calls.length === 1) {
        throw new Error('temporary context failure')
      }
    })
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    function Harness() {
      const [address, setAddress] = useState('A1')
      const contextRef = useRef(createContext(address))
      const sessionRef = useRef({ threadId: 'thr-1' })
      const getContextRef = useRef(() => contextRef.current)
      const snapshot = useMemo(createSnapshot, [])
      contextRef.current = createContext(address)
      const { scheduleContextSync } = useWorkbookAgentContextSync({
        client: { syncThreadContext },
        documentId: 'doc-1',
        enabled: true,
        getContextRef,
        sessionRef,
        snapshot,
      })

      useEffect(() => {
        scheduleContextSync()
      }, [address, scheduleContextSync])

      return (
        <button data-testid="advance" type="button" onClick={() => setAddress((current) => `A${String(Number(current.slice(1)) + 1)}`)}>
          advance
        </button>
      )
    }

    try {
      await act(async () => {
        root.render(<Harness />)
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(200)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)
      expect(syncThreadContext).toHaveBeenLastCalledWith(
        'thr-1',
        expect.objectContaining({ selection: { sheetName: 'Sheet1', address: 'A1' } }),
      )

      await act(async () => {
        const advance = host.querySelector<HTMLButtonElement>("[data-testid='advance']")
        advance?.click()
        advance?.click()
        advance?.click()
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(1_700)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)

      await act(async () => {
        await vi.advanceTimersByTimeAsync(500)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(2)
      expect(syncThreadContext).toHaveBeenLastCalledWith(
        'thr-1',
        expect.objectContaining({ selection: { sheetName: 'Sheet1', address: 'A4' } }),
      )
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not keep pushing the first context sync back when the same context is scheduled repeatedly', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()

    const syncThreadContext = vi.fn(async () => {})
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    function Harness() {
      const [renderCount, setRenderCount] = useState(0)
      const contextRef = useRef(createContext('A1'))
      const sessionRef = useRef({ threadId: 'thr-1' })
      const getContextRef = useRef(() => contextRef.current)
      const snapshot = useMemo(createSnapshot, [])
      const { scheduleContextSync } = useWorkbookAgentContextSync({
        client: { syncThreadContext },
        documentId: 'doc-1',
        enabled: true,
        getContextRef,
        sessionRef,
        snapshot,
      })

      useEffect(() => {
        scheduleContextSync()
      })

      return (
        <button data-testid="rerender" type="button" onClick={() => setRenderCount((current) => current + 1)}>
          {renderCount}
        </button>
      )
    }

    try {
      await act(async () => {
        root.render(<Harness />)
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(100)
      })

      await act(async () => {
        const button = host.querySelector<HTMLButtonElement>("[data-testid='rerender']")
        button?.click()
        button?.click()
        button?.click()
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(60)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)
      expect(syncThreadContext).toHaveBeenLastCalledWith(
        'thr-1',
        expect.objectContaining({ selection: { sheetName: 'Sheet1', address: 'A1' } }),
      )
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })
})
