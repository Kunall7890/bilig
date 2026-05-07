import { describe, expect, it, vi } from 'vitest'
import { CSV_CONTENT_TYPE, type AgentFrame } from '@bilig/agent-api'
import { prepareWorkbookLoad, routeAgentFrame } from './agent-routing.js'

describe('routeAgentFrame', () => {
  it('rejects non-request frames with a shared invalid-frame response', async () => {
    const response = await routeAgentFrame(
      {
        kind: 'response',
        response: { kind: 'ok', id: 'noop' },
      } satisfies AgentFrame,
      {},
      {
        invalidFrameMessage: 'requests only',
        errorCode: 'UNUSED',
        loadWorkbookFile: vi.fn(),
        openWorkbookSession: vi.fn(),
        closeWorkbookSession: vi.fn(),
        getMetrics: vi.fn(),
      },
    )

    expect(response).toEqual({
      kind: 'response',
      response: {
        kind: 'error',
        id: 'unknown',
        code: 'INVALID_AGENT_FRAME',
        message: 'requests only',
        retryable: false,
      },
    })
  })

  it('returns worksheet host unavailable without a worksheet handler', async () => {
    const response = await routeAgentFrame(
      {
        kind: 'request',
        request: {
          kind: 'readRange',
          id: 'read-1',
          sessionId: 'doc:replica',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
        },
      } satisfies AgentFrame,
      {},
      {
        invalidFrameMessage: 'requests only',
        errorCode: 'UNUSED',
        loadWorkbookFile: vi.fn(),
        openWorkbookSession: vi.fn(),
        closeWorkbookSession: vi.fn(),
        getMetrics: vi.fn(),
      },
    )

    expect(response).toEqual({
      kind: 'response',
      response: {
        kind: 'error',
        id: 'read-1',
        code: 'WORKSHEET_HOST_UNAVAILABLE',
        message: 'readRange requires a live worksheet executor, but none is configured for this server',
        retryable: true,
      },
    })
  })

  it('preserves full AgentFrame results from delegated handlers', async () => {
    const response = await routeAgentFrame(
      {
        kind: 'request',
        request: {
          kind: 'openWorkbookSession',
          id: 'open-1',
          documentId: 'doc-1',
          replicaId: 'replica-1',
        },
      } satisfies AgentFrame,
      {},
      {
        invalidFrameMessage: 'requests only',
        errorCode: 'UNUSED',
        loadWorkbookFile: vi.fn(),
        openWorkbookSession: async () =>
          ({
            kind: 'response',
            response: { kind: 'ok', id: 'open-1', sessionId: 'doc-1:replica-1' },
          }) satisfies AgentFrame,
        closeWorkbookSession: vi.fn(),
        getMetrics: vi.fn(),
      },
    )

    expect(response).toEqual({
      kind: 'response',
      response: { kind: 'ok', id: 'open-1', sessionId: 'doc-1:replica-1' },
    })
  })
})

describe('prepareWorkbookLoad', () => {
  it('normalizes workbook upload content type metadata before import', () => {
    const prepared = prepareWorkbookLoad(
      {
        kind: 'loadWorkbookFile',
        id: 'upload-1',
        replicaId: 'agent-local',
        openMode: 'create',
        fileName: 'upload.csv',
        contentType: ' TEXT/CSV; charset=utf-8 ',
        bytesBase64: Buffer.from('A,B\n1,2').toString('base64'),
      },
      {},
    )

    expect(prepared.imported.preview.contentType).toBe(CSV_CONTENT_TYPE)
    expect(prepared.imported.sheetNames).toEqual(['upload'])
    expect(prepared.documentId).toMatch(/^csv:/)
  })
})
