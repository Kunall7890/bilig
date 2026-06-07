import type { AgentFrame } from '@bilig/agent-api'
import { describe, expect, it, vi } from 'vitest'
import { CSV_CONTENT_TYPE, XLSX_CONTENT_TYPE } from '@bilig/agent-api'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import {
  createWorkbookLoadOptions,
  createCloseWorkbookSessionResponse,
  createOpenWorkbookSessionResponse,
  documentIdFromSessionId,
  handleWorkbookAgentFrame,
  loadWorkbookIntoRuntime,
} from './workbook-session-shared.js'

describe('workbook-session-shared', () => {
  it('derives a document id from a session id', () => {
    expect(documentIdFromSessionId('doc-1:replica-1')).toBe('doc-1')
    expect(documentIdFromSessionId('doc-2')).toBe('doc-2')
  })

  it('creates standard open and close session responses', () => {
    expect(createOpenWorkbookSessionResponse('open-1', 'doc-1:replica-1')).toEqual({
      kind: 'ok',
      id: 'open-1',
      sessionId: 'doc-1:replica-1',
    })
    expect(createCloseWorkbookSessionResponse('close-1')).toEqual({
      kind: 'ok',
      id: 'close-1',
    })
  })

  it('prepares workbook imports once and delegates registration and publish hooks', async () => {
    const registerPreparedSession = vi.fn()
    const publishImportedSnapshot = vi.fn()
    const encodedWorkbook = writeSimpleXlsxWorkbook({
      sheets: [{ name: 'Sheet1', cells: [{ address: 'A1', row: 0, col: 0, value: 'hello' }] }],
    })

    const response = await loadWorkbookIntoRuntime(
      {
        kind: 'loadWorkbookFile',
        id: 'load-1',
        replicaId: 'replica-1',
        fileName: 'tiny.xlsx',
        contentType: XLSX_CONTENT_TYPE,
        openMode: 'create',
        bytesBase64: Buffer.from(encodedWorkbook).toString('base64'),
      },
      {
        serverUrl: 'http://127.0.0.1:4321',
        browserAppBaseUrl: 'http://127.0.0.1:3000',
      },
      {
        registerPreparedSession,
        publishImportedSnapshot,
      },
    )

    expect(registerPreparedSession).toHaveBeenCalledTimes(1)
    expect(publishImportedSnapshot).toHaveBeenCalledTimes(1)
    expect(response).toEqual(
      expect.objectContaining({
        kind: 'workbookLoaded',
        id: 'load-1',
        sessionId: expect.stringContaining(':replica-1'),
        serverUrl: 'http://127.0.0.1:4321',
      }),
    )
  })

  it('supports csv imports through the shared workbook load path', async () => {
    const publishImportedSnapshot = vi.fn()

    const response = await loadWorkbookIntoRuntime(
      {
        kind: 'loadWorkbookFile',
        id: 'load-csv-1',
        replicaId: 'replica-1',
        fileName: 'tiny.csv',
        contentType: CSV_CONTENT_TYPE,
        openMode: 'create',
        bytesBase64: Buffer.from('Label,Value\nalpha,12').toString('base64'),
      },
      {
        serverUrl: 'http://127.0.0.1:4321',
      },
      {
        registerPreparedSession: vi.fn(),
        publishImportedSnapshot,
      },
    )

    expect(response).toEqual(
      expect.objectContaining({
        kind: 'workbookLoaded',
        id: 'load-csv-1',
        documentId: expect.stringMatching(/^csv:/),
        sheetNames: ['tiny'],
      }),
    )
    expect(publishImportedSnapshot).toHaveBeenCalledWith(
      expect.stringMatching(/^csv:/),
      expect.objectContaining({
        sheets: [
          expect.objectContaining({
            name: 'tiny',
          }),
        ],
      }),
      expect.any(Object),
    )
  })

  it('supports xlsb imports through the shared workbook load path', async () => {
    const publishImportedSnapshot = vi.fn()
    const encodedWorkbook = tinyXlsbWorkbookBytes()

    const response = await loadWorkbookIntoRuntime(
      {
        kind: 'loadWorkbookFile',
        id: 'load-xlsb-1',
        replicaId: 'replica-1',
        fileName: 'tiny.xlsb',
        contentType: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
        openMode: 'create',
        bytesBase64: Buffer.from(encodedWorkbook).toString('base64'),
      },
      {
        serverUrl: 'http://127.0.0.1:4321',
      },
      {
        registerPreparedSession: vi.fn(),
        publishImportedSnapshot,
      },
    )

    expect(response).toEqual(
      expect.objectContaining({
        kind: 'workbookLoaded',
        id: 'load-xlsb-1',
        documentId: expect.stringMatching(/^xlsb:/),
        sheetNames: ['Sheet1'],
      }),
    )
    expect(publishImportedSnapshot).toHaveBeenCalledWith(
      expect.stringMatching(/^xlsb:/),
      expect.objectContaining({
        sheets: [
          expect.objectContaining({
            name: 'Sheet1',
          }),
        ],
      }),
      expect.any(Object),
    )
  })

  it('normalizes open and close workbook session lifecycle responses', async () => {
    const openResponse = await handleWorkbookAgentFrame(
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
        invalidFrameMessage: 'bad frame',
        errorCode: 'TEST_ERROR',
        loadWorkbookFile: vi.fn(),
        openWorkbookSession: async () => 'doc-1:replica-1',
        closeWorkbookSession: async () => undefined,
        getMetrics: async (request) => ({ kind: 'metrics', id: request.id, value: { ok: true } }),
      },
    )

    expect(openResponse).toEqual({
      kind: 'response',
      response: {
        kind: 'ok',
        id: 'open-1',
        sessionId: 'doc-1:replica-1',
      },
    })

    const closeResponse = await handleWorkbookAgentFrame(
      {
        kind: 'request',
        request: {
          kind: 'closeWorkbookSession',
          id: 'close-1',
          sessionId: 'doc-1:replica-1',
        },
      } satisfies AgentFrame,
      {},
      {
        invalidFrameMessage: 'bad frame',
        errorCode: 'TEST_ERROR',
        loadWorkbookFile: vi.fn(),
        openWorkbookSession: async () => 'doc-1:replica-1',
        closeWorkbookSession: async () => undefined,
        getMetrics: async (request) => ({ kind: 'metrics', id: request.id, value: { ok: true } }),
      },
    )

    expect(closeResponse).toEqual({
      kind: 'response',
      response: {
        kind: 'ok',
        id: 'close-1',
      },
    })
  })

  it('builds workbook load options without undefined keys', () => {
    const options = createWorkbookLoadOptions(
      {
        browserAppBaseUrl: 'http://127.0.0.1:3000',
      },
      {
        registerPreparedSession: vi.fn(),
        publishImportedSnapshot: vi.fn(),
      },
    )

    expect(options).toEqual({
      browserAppBaseUrl: 'http://127.0.0.1:3000',
      registerPreparedSession: expect.any(Function),
      publishImportedSnapshot: expect.any(Function),
    })
  })
})

function tinyXlsbWorkbookBytes(): Uint8Array {
  return Buffer.from(
    [
      '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',
      'AAAA',
    ].join(''),
    'base64',
  )
}
