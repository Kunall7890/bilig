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

function createRenderedContext(address: string, capturedRevision: number): WorkbookAgentUiContext {
  return {
    ...createContext(address),
    rendered: {
      capturedAtUnixMs: capturedRevision * 100,
      capturedRevision,
      batchId: capturedRevision,
      selection: null,
      visibleRange: null,
    },
  }
}

function createSnapshot(status: WorkbookAgentThreadSnapshot['status'] = 'idle'): WorkbookAgentThreadSnapshot {
  return {
    activeTurnId: status === 'inProgress' ? 'turn-1' : null,
    context: createContext('A1'),
    documentId: 'doc-1',
    entries: [],
    executionPolicy: 'autoApplyAll',
    executionRecords: [],
    lastError: null,
    reviewQueueItems: [],
    scope: 'private',
    status,
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

  it('skips idle viewport-only context churn after the initial sync', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()

    const syncThreadContext = vi.fn(async () => {})
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    function Harness() {
      const [rowEnd, setRowEnd] = useState(20)
      const contextRef = useRef({
        ...createContext('A1'),
        viewport: {
          colEnd: 10,
          colStart: 0,
          rowEnd,
          rowStart: 0,
        },
      })
      const sessionRef = useRef({ threadId: 'thr-1' })
      const getContextRef = useRef(() => contextRef.current)
      const snapshot = useMemo(createSnapshot, [])
      contextRef.current = {
        ...createContext('A1'),
        viewport: {
          colEnd: 10,
          colStart: 0,
          rowEnd,
          rowStart: 0,
        },
      }
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
      }, [rowEnd, scheduleContextSync])

      return (
        <button data-testid="viewport" type="button" onClick={() => setRowEnd((current) => current + 1)}>
          {rowEnd}
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

      await act(async () => {
        host.querySelector<HTMLButtonElement>("[data-testid='viewport']")?.click()
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(200)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)

      await act(async () => {
        await vi.advanceTimersByTimeAsync(1_000)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('skips idle rendered proof revision churn after the initial sync', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()

    const syncThreadContext = vi.fn(async () => {})
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    function Harness() {
      const [capturedRevision, setCapturedRevision] = useState(7)
      const contextRef = useRef(createRenderedContext('A1', capturedRevision))
      const sessionRef = useRef({ threadId: 'thr-1' })
      const getContextRef = useRef(() => contextRef.current)
      const snapshot = useMemo(createSnapshot, [])
      contextRef.current = createRenderedContext('A1', capturedRevision)
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
      }, [capturedRevision, scheduleContextSync])

      return (
        <button data-testid="revision" type="button" onClick={() => setCapturedRevision((current) => current + 1)}>
          {capturedRevision}
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

      await act(async () => {
        host.querySelector<HTMLButtonElement>("[data-testid='revision']")?.click()
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(1_000)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)
      expect(syncThreadContext).toHaveBeenLastCalledWith(
        'thr-1',
        expect.objectContaining({ rendered: expect.objectContaining({ capturedRevision: 7 }) }),
      )
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('throttles passive rendered context churn during active assistant turns', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()

    const syncThreadContext = vi.fn(async () => {})
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    function Harness() {
      const [rowEnd, setRowEnd] = useState(20)
      const contextRef = useRef({
        ...createContext('A1'),
        viewport: {
          colEnd: 10,
          colStart: 0,
          rowEnd,
          rowStart: 0,
        },
      })
      const sessionRef = useRef({ threadId: 'thr-1' })
      const getContextRef = useRef(() => contextRef.current)
      const snapshot = useMemo(() => createSnapshot('inProgress'), [])
      contextRef.current = {
        ...createContext('A1'),
        viewport: {
          colEnd: 10,
          colStart: 0,
          rowEnd,
          rowStart: 0,
        },
      }
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
      }, [rowEnd, scheduleContextSync])

      return (
        <button data-testid="viewport" type="button" onClick={() => setRowEnd((current) => current + 1)}>
          {rowEnd}
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

      await act(async () => {
        host.querySelector<HTMLButtonElement>("[data-testid='viewport']")?.click()
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(800)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)

      await act(async () => {
        await vi.advanceTimersByTimeAsync(4_200)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(2)
      expect(syncThreadContext).toHaveBeenLastCalledWith(
        'thr-1',
        expect.objectContaining({ viewport: expect.objectContaining({ rowEnd: 21 }) }),
      )
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('coalesces rendered proof freshness during active assistant turns onto the passive sync cadence', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()

    const syncThreadContext = vi.fn(async () => {})
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    function Harness() {
      const [capturedRevision, setCapturedRevision] = useState(7)
      const contextRef = useRef(createRenderedContext('A1', capturedRevision))
      const sessionRef = useRef({ threadId: 'thr-1' })
      const getContextRef = useRef(() => contextRef.current)
      const snapshot = useMemo(() => createSnapshot('inProgress'), [])
      contextRef.current = createRenderedContext('A1', capturedRevision)
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
      }, [capturedRevision, scheduleContextSync])

      return (
        <button data-testid="revision" type="button" onClick={() => setCapturedRevision((current) => current + 1)}>
          {capturedRevision}
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

      const advanceRevision = async () => {
        await act(async () => {
          host.querySelector<HTMLButtonElement>("[data-testid='revision']")?.click()
        })
      }

      await advanceRevision()
      await advanceRevision()
      await advanceRevision()

      await act(async () => {
        await vi.advanceTimersByTimeAsync(4_800)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(1)

      await act(async () => {
        await vi.advanceTimersByTimeAsync(250)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(2)
      expect(syncThreadContext).toHaveBeenLastCalledWith(
        'thr-1',
        expect.objectContaining({ rendered: expect.objectContaining({ capturedRevision: 10 }) }),
      )
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('prioritizes selection context updates after an initial sync', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()

    const syncThreadContext = vi.fn(async () => {})
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
        <button data-testid="selection" type="button" onClick={() => setAddress('B2')}>
          {address}
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

      await act(async () => {
        host.querySelector<HTMLButtonElement>("[data-testid='selection']")?.click()
      })

      await act(async () => {
        await vi.advanceTimersByTimeAsync(160)
      })

      expect(syncThreadContext).toHaveBeenCalledTimes(2)
      expect(syncThreadContext).toHaveBeenLastCalledWith(
        'thr-1',
        expect.objectContaining({ selection: { sheetName: 'Sheet1', address: 'B2' } }),
      )
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })
})
