import { describe, expect, it } from 'vitest'
import { applyCellDiff, loadWorkbookEventRecordsAfter, type QueryResultRow, type Queryable } from '../store.js'
import type { CellSourceRow } from '../projection.js'

class FakeQueryable implements Queryable {
  constructor(private readonly rows: readonly QueryResultRow[]) {}

  async query<T extends QueryResultRow = QueryResultRow>(): Promise<{ rows: T[] }> {
    return { rows: this.rows.filter((row): row is T => row !== null) }
  }
}

class ConcurrentDetectingQueryable implements Queryable {
  readonly calls: string[] = []
  maxActiveQueries = 0
  private activeQueries = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string): Promise<{ rows: T[] }> {
    this.calls.push(text)
    this.activeQueries += 1
    this.maxActiveQueries = Math.max(this.maxActiveQueries, this.activeQueries)
    await new Promise((resolve) => setTimeout(resolve, 0))
    this.activeQueries -= 1
    return { rows: [] }
  }
}

function createCellSourceRow(address: string, value: unknown): CellSourceRow {
  const [, colLabel, rowLabel] = /^([A-Z]+)(\d+)$/u.exec(address) ?? []
  const colNum = colLabel ? colLabel.charCodeAt(0) - 'A'.charCodeAt(0) : 0
  const rowNum = rowLabel ? Number(rowLabel) - 1 : 0
  return {
    workbookId: 'book-1',
    sheetName: 'Sheet1',
    address,
    rowNum,
    colNum,
    inputValue: value,
    formula: null,
    format: null,
    styleId: null,
    explicitFormatId: null,
    sourceRevision: 2,
    updatedBy: 'user-1',
    updatedAt: '2026-05-16T00:00:00.000Z',
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

describe('zero projection writes', () => {
  it('serializes cell projection diff writes so transaction clients do not have in-flight batches', async () => {
    const db = new ConcurrentDetectingQueryable()

    await applyCellDiff(db, [createCellSourceRow('A1', 1)], [createCellSourceRow('B1', 2), createCellSourceRow('C1', 3)])

    expect(db.maxActiveQueries).toBe(1)
    expect(db.calls).toHaveLength(3)
    expect(db.calls[0]).toContain('DELETE FROM cells')
    expect(db.calls[1]).toContain('INSERT INTO cells')
    expect(db.calls[2]).toContain('INSERT INTO cells')
  })
})
