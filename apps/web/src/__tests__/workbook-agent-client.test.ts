// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createWorkbookAgentClient } from '../workbook-agent-client.js'

function createContext() {
  return {
    selection: {
      sheetName: 'Sheet1',
      address: 'A1',
    },
    viewport: {
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 5,
    },
  }
}

function createSnapshot(overrides: Record<string, unknown> = {}) {
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: createContext(),
    entries: [],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  }
}

function createThreadSummary(overrides: Record<string, unknown> = {}) {
  return {
    threadId: 'thr-1',
    scope: 'private',
    ownerUserId: 'alex@example.com',
    updatedAtUnixMs: 100,
    entryCount: 1,
    reviewQueueItemCount: 0,
    latestEntryText: null,
    ...overrides,
  }
}

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('workbook agent client', () => {
  it('builds thread urls and decodes successful responses', async () => {
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url
      if (url.endsWith('/chat/threads')) {
        return new Response(JSON.stringify([createThreadSummary()]), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      return new Response(JSON.stringify(createSnapshot()), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })
    vi.stubGlobal('fetch', fetchSpy)

    const client = createWorkbookAgentClient('doc-1')
    expect(client.threadEventsUrl('thr/1')).toBe('/v2/documents/doc-1/chat/threads/thr%2F1/events')
    await expect(client.loadThreadSummaries()).resolves.toEqual([expect.objectContaining({ threadId: 'thr-1' })])
    await expect(client.loadThreadSnapshot('thr-1')).resolves.toEqual(expect.objectContaining({ threadId: 'thr-1' }))
  })

  it('surfaces server-provided error messages', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(JSON.stringify({ message: 'Prompt rejected by server' }), {
            status: 422,
            headers: { 'content-type': 'application/json' },
          }),
      ),
    )

    const client = createWorkbookAgentClient('doc-1')
    await expect(client.sendPrompt('thr-1', 'Continue working', createContext())).rejects.toThrow('Prompt rejected by server')
  })
})
