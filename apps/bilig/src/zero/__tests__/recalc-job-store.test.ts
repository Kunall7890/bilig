import { beforeEach, describe, expect, it, vi } from 'vitest'
import { createEmptyWorkbookSnapshot } from '../store-support.js'

const storeFns = vi.hoisted(() => ({
  persistCellEvalDiff: vi.fn(),
  persistCellEvalIncremental: vi.fn(),
  persistWorkbookCheckpoint: vi.fn(),
  shouldPersistWorkbookCheckpointRevision: vi.fn((revision: number) => revision === 64),
}))

vi.mock('../store.js', () => ({
  shouldPersistWorkbookCheckpointRevision: storeFns.shouldPersistWorkbookCheckpointRevision,
}))

vi.mock('../workbook-calculation-store.js', () => ({
  persistCellEvalDiff: storeFns.persistCellEvalDiff,
  persistCellEvalIncremental: storeFns.persistCellEvalIncremental,
  persistWorkbookCheckpoint: storeFns.persistWorkbookCheckpoint,
}))

import { leaseNextRecalcJob, markRecalcJobCompleted, markRecalcJobFailed } from '../recalc-job-store.js'
import type { Queryable, QueryResultRow } from '../store.js'

function typedRows<T extends QueryResultRow>(rows: readonly QueryResultRow[]): T[] {
  return rows.filter((row): row is T => row !== null)
}

class FakeTransactionClient implements Queryable {
  readonly calls: { text: string; values: readonly unknown[] | undefined }[] = []
  releaseCount = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    if (text.includes('SELECT head_revision')) {
      return { rows: typedRows<T>([{ head_revision: 64 }]) }
    }
    return { rows: [] }
  }

  release(): void {
    this.releaseCount += 1
  }
}

class FakeTransactionalQueryable implements Queryable {
  readonly calls: { text: string; values: readonly unknown[] | undefined }[] = []
  readonly client = new FakeTransactionClient()
  connectCount = 0

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    if (text.includes('SELECT head_revision')) {
      return { rows: typedRows<T>([{ head_revision: 64 }]) }
    }
    return { rows: [] as T[] }
  }

  async connect(): Promise<FakeTransactionClient> {
    this.connectCount += 1
    return this.client
  }
}

describe('recalc job store', () => {
  beforeEach(() => {
    storeFns.persistCellEvalDiff.mockClear()
    storeFns.persistCellEvalIncremental.mockClear()
    storeFns.persistWorkbookCheckpoint.mockClear()
    storeFns.shouldPersistWorkbookCheckpointRevision.mockClear()
  })

  it('filters invalid dirty regions when leasing a job', async () => {
    const db: Queryable = {
      query: vi.fn().mockResolvedValue({
        rows: [
          {
            id: 'job-1',
            workbook_id: 'book-1',
            from_revision: '3',
            to_revision: '7',
            dirty_regions_json: [
              {
                sheetName: 'Sheet1',
                rowStart: 1,
                rowEnd: 2,
                colStart: 3,
                colEnd: 4,
              },
              { nope: true },
            ],
            attempts: '2',
          },
        ],
      }),
    }

    await expect(leaseNextRecalcJob(db, 'worker-1')).resolves.toEqual({
      id: 'job-1',
      workbookId: 'book-1',
      fromRevision: 3,
      toRevision: 7,
      dirtyRegions: [
        {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 2,
          colStart: 3,
          colEnd: 4,
        },
      ],
      attempts: 2,
    })
  })

  it('fails malformed leased jobs instead of coercing revisions to zero', async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({
        rows: [
          {
            id: 'job-bad',
            workbook_id: 'book-1',
            from_revision: '7',
            to_revision: '7',
            dirty_regions_json: null,
            attempts: '1',
          },
        ],
      })
      .mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await expect(leaseNextRecalcJob(db, 'worker-1')).resolves.toBeNull()

    expect(query).toHaveBeenLastCalledWith(expect.stringContaining("SET status = 'failed'"), ['job-bad', 'Malformed recalc job lease row'])
  })

  it('persists incremental results and checkpoints completed lease revisions', async () => {
    const db: Queryable = {
      query: vi
        .fn()
        .mockResolvedValueOnce({ rows: [{ head_revision: 64 }] })
        .mockResolvedValue({ rows: [] }),
    }
    const snapshot = createEmptyWorkbookSnapshot('book-1')

    await expect(
      markRecalcJobCompleted(
        db,
        {
          id: 'job-1',
          workbookId: 'book-1',
          fromRevision: 1,
          toRevision: 64,
          dirtyRegions: null,
          attempts: 1,
        },
        [],
        snapshot,
        null,
        true,
      ),
    ).resolves.toBe(true)

    expect(storeFns.persistCellEvalIncremental).toHaveBeenCalledWith(db, 'book-1', [])
    expect(storeFns.persistCellEvalDiff).not.toHaveBeenCalled()
    expect(storeFns.persistWorkbookCheckpoint).toHaveBeenCalledWith(db, 'book-1', 64, snapshot, null)
  })

  it('rolls back the whole recalc completion when checkpoint persistence fails', async () => {
    const db = new FakeTransactionalQueryable()
    const snapshot = createEmptyWorkbookSnapshot('book-1')
    storeFns.persistWorkbookCheckpoint.mockRejectedValueOnce(new Error('checkpoint failed'))

    await expect(
      markRecalcJobCompleted(
        db,
        {
          id: 'job-1',
          workbookId: 'book-1',
          fromRevision: 1,
          toRevision: 64,
          dirtyRegions: null,
          attempts: 1,
        },
        [],
        snapshot,
        null,
        true,
      ),
    ).rejects.toThrow('checkpoint failed')

    expect(db.connectCount).toBe(1)
    expect(db.calls).toEqual([])
    expect(db.client.releaseCount).toBe(1)
    expect(db.client.calls[0]?.text).toBe('BEGIN')
    expect(db.client.calls.at(-1)?.text).toBe('ROLLBACK')
    expect(db.client.calls.some((call) => call.text === 'COMMIT')).toBe(false)
    expect(storeFns.persistCellEvalIncremental).toHaveBeenCalledWith(db.client, 'book-1', [])
    expect(storeFns.persistWorkbookCheckpoint).toHaveBeenCalledWith(db.client, 'book-1', 64, snapshot, null)
  })

  it('rejects invalid workbook head revisions before persisting recalc output', async () => {
    const db: Queryable = {
      query: vi.fn().mockResolvedValueOnce({ rows: [{ head_revision: '-1' }] }),
    }

    await expect(
      markRecalcJobCompleted(
        db,
        {
          id: 'job-2',
          workbookId: 'book-1',
          fromRevision: 1,
          toRevision: 2,
          dirtyRegions: null,
          attempts: 1,
        },
        [],
        null,
        null,
        true,
      ),
    ).rejects.toThrow('Invalid workbook head revision while completing recalc job job-2')

    expect(storeFns.persistCellEvalIncremental).not.toHaveBeenCalled()
    expect(storeFns.persistCellEvalDiff).not.toHaveBeenCalled()
  })

  it('marks exhausted failures as failed instead of pending', async () => {
    const query = vi.fn().mockResolvedValue({ rows: [] })
    const db: Queryable = { query }

    await markRecalcJobFailed(
      db,
      {
        id: 'job-9',
        workbookId: 'book-1',
        fromRevision: 1,
        toRevision: 2,
        dirtyRegions: null,
        attempts: 3,
      },
      new Error('boom'),
    )

    expect(query).toHaveBeenCalledWith(expect.any(String), ['job-9', 'failed', expect.stringContaining('boom')])
  })
})
