import { beforeEach, describe, expect, it, vi } from 'vitest'

const storeFns = vi.hoisted(() => ({
  applyAxisMetadataDiff: vi.fn(),
  applyCalculationSettings: vi.fn(),
  applyCellDiff: vi.fn(),
  applyDefinedNameDiff: vi.fn(),
  applyNumberFormatDiff: vi.fn(),
  applySheetDiff: vi.fn(),
  applyStyleDiff: vi.fn(),
  applyWorkbookMetadataDiff: vi.fn(),
  insertWorkbookHeaderIfMissing: vi.fn(),
  persistCellEvalRows: vi.fn(),
  repairWorkbookSheetIds: vi.fn(),
  upsertWorkbookHeader: vi.fn(),
}))

vi.mock('../store.js', () => ({
  applyAxisMetadataDiff: storeFns.applyAxisMetadataDiff,
  applyCalculationSettings: storeFns.applyCalculationSettings,
  applyCellDiff: storeFns.applyCellDiff,
  applyDefinedNameDiff: storeFns.applyDefinedNameDiff,
  applyNumberFormatDiff: storeFns.applyNumberFormatDiff,
  applySheetDiff: storeFns.applySheetDiff,
  applyStyleDiff: storeFns.applyStyleDiff,
  applyWorkbookMetadataDiff: storeFns.applyWorkbookMetadataDiff,
  insertWorkbookHeaderIfMissing: storeFns.insertWorkbookHeaderIfMissing,
  upsertWorkbookHeader: storeFns.upsertWorkbookHeader,
}))

vi.mock('../workbook-calculation-store.js', () => ({
  persistCellEvalRows: storeFns.persistCellEvalRows,
}))

vi.mock('../sheet-id-repair.js', () => ({
  repairWorkbookSheetIds: storeFns.repairWorkbookSheetIds,
}))

import {
  backfillCellEvalStyleJson,
  backfillWorkbookSourceProjectionVersion,
  dropLegacyZeroSyncSchemaObjects,
  enforceWorkbookSheetIdInvariant,
  enforceWorkbookEventClientMutationIdUniqueness,
  ensureWorkbookDocumentExists,
  repairWorkbookSheetIdsForMigration,
} from '../workbook-migration-store.js'
import type { QueryResultRow, Queryable } from '../store.js'

interface RecordedQuery {
  readonly text: string
  readonly values: readonly unknown[] | undefined
}

class FakeTransactionClient implements Queryable {
  readonly calls: RecordedQuery[] = []
  releaseCount = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    return { rows: [] }
  }

  release(): void {
    this.releaseCount += 1
  }
}

class FakeTransactionalQueryable implements Queryable {
  readonly calls: RecordedQuery[] = []
  readonly client = new FakeTransactionClient()
  connectCount = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    return { rows: [] as T[] }
  }

  async connect(): Promise<FakeTransactionClient> {
    this.connectCount += 1
    return this.client
  }
}

describe('workbook migration store', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  it('skips projection replacement when the workbook already exists', async () => {
    storeFns.insertWorkbookHeaderIfMissing.mockResolvedValueOnce(false)
    const query = vi.fn()
    const db: Queryable = { query }

    await ensureWorkbookDocumentExists(db, 'book-1', 'owner-1')

    expect(storeFns.insertWorkbookHeaderIfMissing).toHaveBeenCalledOnce()
    expect(storeFns.applySheetDiff).not.toHaveBeenCalled()
    expect(query).not.toHaveBeenCalled()
  })

  it('creates missing workbooks atomically when the queryable supports transactions', async () => {
    storeFns.insertWorkbookHeaderIfMissing.mockResolvedValueOnce(true)
    const db = new FakeTransactionalQueryable()

    await ensureWorkbookDocumentExists(db, 'book-1', 'owner-1')

    expect(db.connectCount).toBe(1)
    expect(db.calls).toEqual([])
    expect(db.client.releaseCount).toBe(1)
    expect(db.client.calls[0]?.text).toBe('BEGIN')
    expect(db.client.calls.at(-1)?.text).toBe('COMMIT')
    expect(storeFns.insertWorkbookHeaderIfMissing).toHaveBeenCalledWith(db.client, 'book-1', expect.any(Object), expect.any(Object), null)
    expect(storeFns.applySheetDiff).toHaveBeenCalledWith(db.client, [], expect.any(Array))
    expect(db.client.calls.some((call) => call.text.includes('DELETE FROM sheets'))).toBe(true)
  })

  it('drops the legacy zero-sync schema objects', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await dropLegacyZeroSyncSchemaObjects(db)

    expect(query.mock.calls).toEqual([
      [`DROP INDEX IF EXISTS sheet_style_ranges_workbook_sheet_idx`],
      [`DROP INDEX IF EXISTS sheet_format_ranges_workbook_sheet_idx`],
      [`DROP TABLE IF EXISTS sheet_style_ranges`],
      [`DROP TABLE IF EXISTS sheet_format_ranges`],
    ])
  })

  it('repairs sheet ids without blind sort-order backfills that can violate uniqueness', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await repairWorkbookSheetIdsForMigration(db)

    expect(storeFns.repairWorkbookSheetIds).toHaveBeenCalledWith(db)
    expect(query).not.toHaveBeenCalledWith(`UPDATE sheets SET sheet_id = sort_order + 1 WHERE sheet_id IS NULL`)
  })

  it('repairs and enforces the shared non-null positive sheet id invariant', async () => {
    const db = new FakeTransactionalQueryable()

    await enforceWorkbookSheetIdInvariant(db)

    expect(storeFns.repairWorkbookSheetIds).toHaveBeenCalledWith(db.client)
    expect(db.client.calls.map((call) => call.text)).toEqual([
      'BEGIN',
      `ALTER TABLE sheets ALTER COLUMN sheet_id SET NOT NULL`,
      expect.stringContaining('ADD CONSTRAINT sheets_sheet_id_positive_chk CHECK (sheet_id > 0)'),
      'COMMIT',
    ])
    expect(db.client.releaseCount).toBe(1)
  })

  it('returns early from projection backfill when no legacy workbook ids are found', async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({ rows: [{ relation: null }] })
      .mockResolvedValueOnce({ rows: [{ relation: null }] })
      .mockResolvedValueOnce({ rows: [] })
    const db: Queryable = { query }

    await backfillWorkbookSourceProjectionVersion(db)

    expect(query).toHaveBeenCalledTimes(3)
    expect(storeFns.upsertWorkbookHeader).not.toHaveBeenCalled()
    expect(storeFns.persistCellEvalRows).not.toHaveBeenCalled()
  })

  it('serializes projection rebuilds across legacy workbooks on one migration transaction client', async () => {
    let activeRebuilds = 0
    let maxActiveRebuilds = 0
    storeFns.applySheetDiff.mockImplementation(async () => {
      activeRebuilds += 1
      maxActiveRebuilds = Math.max(maxActiveRebuilds, activeRebuilds)
      await new Promise((resolve) => setTimeout(resolve, 0))
      activeRebuilds -= 1
    })
    const query = vi.fn(async (text: string) => {
      if (text.includes('source_projection_version')) {
        return { rows: [{ id: 'book-1' }, { id: 'book-2' }] }
      }
      if (text.includes('to_regclass')) {
        return { rows: [{ relation: null }] }
      }
      if (text.includes('WHERE id = ANY')) {
        return {
          rows: ['book-1', 'book-2'].map((id) => ({
            id,
            snapshot: null,
            replica_snapshot: null,
            head_revision: 0,
            calculated_revision: 0,
            owner_user_id: `${id}-owner`,
            updated_at: '2026-05-16T00:00:00.000Z',
          })),
        }
      }
      return { rows: [] }
    })
    const db: Queryable = { query }

    await backfillWorkbookSourceProjectionVersion(db)

    expect(maxActiveRebuilds).toBe(1)
    expect(storeFns.applySheetDiff).toHaveBeenCalledTimes(2)
    expect(storeFns.upsertWorkbookHeader).toHaveBeenCalledTimes(2)
  })

  it('rejects invalid workbook revision metadata during projection backfill', async () => {
    const query = vi.fn(async (text: string) => {
      if (text.includes('source_projection_version')) {
        return { rows: [{ id: 'book-bad' }] }
      }
      if (text.includes('to_regclass')) {
        return { rows: [{ relation: null }] }
      }
      if (text.includes('WHERE id = ANY')) {
        return {
          rows: [
            {
              id: 'book-bad',
              snapshot: null,
              replica_snapshot: null,
              head_revision: '-1',
              calculated_revision: '0',
              owner_user_id: 'owner-1',
              updated_at: '2026-05-16T00:00:00.000Z',
            },
          ],
        }
      }
      return { rows: [] }
    })
    const db: Queryable = { query }

    await expect(backfillWorkbookSourceProjectionVersion(db)).rejects.toThrow('Invalid workbook migration row for book-bad')

    expect(storeFns.upsertWorkbookHeader).not.toHaveBeenCalled()
    expect(storeFns.applySheetDiff).not.toHaveBeenCalled()
  })

  it('returns early from cell-eval backfill when no stale style_json rows are found', async () => {
    const query = vi.fn().mockResolvedValueOnce({ rows: [] })
    const db: Queryable = { query }

    await backfillCellEvalStyleJson(db)

    expect(query).toHaveBeenCalledOnce()
    expect(storeFns.persistCellEvalRows).not.toHaveBeenCalled()
  })

  it('creates the workbook event client mutation uniqueness index after verifying no duplicates exist', async () => {
    const query = vi.fn().mockResolvedValueOnce({ rows: [] }).mockResolvedValueOnce({ rows: [] })
    const db: Queryable = { query }

    await enforceWorkbookEventClientMutationIdUniqueness(db)

    expect(query.mock.calls[0]?.[0]).toContain('HAVING COUNT(*) > 1')
    expect(query.mock.calls[1]?.[0]).toContain('CREATE UNIQUE INDEX IF NOT EXISTS workbook_event_workbook_client_mutation_idx')
  })

  it('rejects workbook event client mutation uniqueness when existing duplicate ids are present', async () => {
    const query = vi.fn().mockResolvedValueOnce({
      rows: [
        {
          workbook_id: 'book-1',
          client_mutation_id: 'book-1:pending:4',
          duplicate_count: 2,
          first_revision: 7,
          last_revision: 9,
        },
      ],
    })
    const db: Queryable = { query }

    await expect(enforceWorkbookEventClientMutationIdUniqueness(db)).rejects.toThrow('book-1/book-1:pending:4 count=2 revisions=7-9')
    expect(query).toHaveBeenCalledOnce()
  })

  it('reports malformed duplicate workbook event revisions explicitly', async () => {
    const query = vi.fn().mockResolvedValueOnce({
      rows: [
        {
          workbook_id: 'book-1',
          client_mutation_id: 'book-1:pending:4',
          duplicate_count: 'bad',
          first_revision: '-1',
          last_revision: '9',
        },
      ],
    })
    const db: Queryable = { query }

    await expect(enforceWorkbookEventClientMutationIdUniqueness(db)).rejects.toThrow(
      'book-1/book-1:pending:4 count=<invalid> revisions=<invalid>-9',
    )
    expect(query).toHaveBeenCalledOnce()
  })
})
