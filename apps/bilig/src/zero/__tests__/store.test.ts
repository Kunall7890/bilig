import { describe, expect, it } from 'vitest'
import { loadWorkbookEventRecordsAfter, type QueryResultRow, type Queryable } from '../store.js'

class FakeQueryable implements Queryable {
  constructor(private readonly rows: readonly QueryResultRow[]) {}

  async query<T extends QueryResultRow = QueryResultRow>(): Promise<{ rows: T[] }> {
    return { rows: this.rows.filter((row): row is T => row !== null) }
  }
}

describe('zero store event replay', () => {
  it('drops workbook_event rows with invalid replay revisions', async () => {
    const validPayload = {
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'A1',
      value: 1,
    }
    const db = new FakeQueryable([
      {
        revision: -1,
        client_mutation_id: 'mutation-negative',
        txn_json: validPayload,
      },
      {
        revision: 0,
        client_mutation_id: 'mutation-zero',
        txn_json: validPayload,
      },
      {
        revision: String(Number.MAX_SAFE_INTEGER + 1),
        client_mutation_id: 'mutation-unsafe',
        txn_json: validPayload,
      },
      {
        revision: '4',
        client_mutation_id: 'mutation-4',
        txn_json: validPayload,
      },
    ])

    await expect(loadWorkbookEventRecordsAfter(db, 'book-1', 1)).resolves.toEqual([
      {
        revision: 4,
        clientMutationId: 'mutation-4',
        payload: validPayload,
      },
    ])
  })
})
