import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { createMemoryWorkbookLocalStoreFactory } from '@bilig/storage-browser'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { WorkbookWorkerRuntime } from '../worker-runtime.js'

function buildSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'bootstrap-doc' },
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        cells: [
          {
            address: 'A1',
            value: 42,
          },
        ],
      },
    ],
  }
}

describe('worker runtime authoritative bootstrap', () => {
  it('imports a bootstrap authoritative snapshot once and reuses projection state for persistence', async () => {
    const importSnapshot = vi.spyOn(SpreadsheetEngine.prototype, 'importSnapshot')
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryWorkbookLocalStoreFactory(),
    })

    await runtime.bootstrap({
      documentId: 'single-import-bootstrap-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    await runtime.installAuthoritativeSnapshot({
      snapshot: buildSnapshot(),
      authoritativeRevision: 3,
      mode: 'bootstrap',
    })

    expect(importSnapshot).toHaveBeenCalledTimes(1)
    expect(runtime.getAuthoritativeRevision()).toBe(3)
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 4,
          clientMutationId: null,
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'A1',
            value: 84,
          },
        },
      ],
      4,
    )

    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 84,
    })
  })
})
