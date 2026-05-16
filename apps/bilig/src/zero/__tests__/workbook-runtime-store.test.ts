import { describe, expect, it, vi } from 'vitest'
import { createEmptyWorkbookSnapshot } from '../store-support.js'

const storeFns = vi.hoisted(() => ({
  loadWorkbookEventRecordsAfter: vi.fn(),
}))

vi.mock('../store.js', async () => ({
  loadWorkbookEventRecordsAfter: storeFns.loadWorkbookEventRecordsAfter,
}))

import { acquireWorkbookMutationLock, loadWorkbookRuntimeMetadata, loadWorkbookState } from '../workbook-runtime-store.js'
import type { WorkbookRuntimeStoreConnection } from '../store.js'

describe('workbook runtime store', () => {
  it('returns inline workbook state without replay', async () => {
    const snapshot = createEmptyWorkbookSnapshot('book-1')
    const db: WorkbookRuntimeStoreConnection = {
      run: vi.fn(),
      query: vi
        .fn()
        .mockResolvedValueOnce({
          rows: [
            {
              owner_user_id: 'owner-1',
              head_revision: '4',
              calculated_revision: '3',
            },
          ],
        })
        .mockResolvedValueOnce({
          rows: [
            {
              snapshot,
              replica_snapshot: null,
            },
          ],
        }),
    }

    await expect(loadWorkbookState(db, 'book-1')).resolves.toEqual({
      snapshot,
      replicaSnapshot: null,
      headRevision: 4,
      calculatedRevision: 3,
      ownerUserId: 'owner-1',
    })
    expect(storeFns.loadWorkbookEventRecordsAfter).not.toHaveBeenCalled()
  })

  it('replays events from the latest checkpoint when inline state is absent', async () => {
    storeFns.loadWorkbookEventRecordsAfter.mockResolvedValueOnce([
      {
        revision: 2,
        payload: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 123,
        },
      },
    ])
    const checkpoint = createEmptyWorkbookSnapshot('book-2')
    const db: WorkbookRuntimeStoreConnection = {
      run: vi.fn(),
      query: vi
        .fn()
        .mockResolvedValueOnce({
          rows: [
            {
              owner_user_id: 'owner-2',
              head_revision: '2',
              calculated_revision: '1',
            },
          ],
        })
        .mockResolvedValueOnce({
          rows: [
            {
              snapshot: null,
              replica_snapshot: null,
            },
          ],
        })
        .mockResolvedValueOnce({
          rows: [
            {
              revision: '1',
              payload: checkpoint,
              replica_snapshot: null,
            },
          ],
        }),
    }

    const loaded = await loadWorkbookState(db, 'book-2')

    expect(loaded.headRevision).toBe(2)
    expect(loaded.snapshot.sheets[0]?.cells.length).toBeGreaterThan(0)
    expect(storeFns.loadWorkbookEventRecordsAfter).toHaveBeenCalledWith(db, 'book-2', 1)
  })

  it('loads metadata and acquires advisory locks', async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({
        rows: [
          {
            owner_user_id: 'owner-3',
            head_revision: '6',
            calculated_revision: '5',
          },
        ],
      })
      .mockResolvedValueOnce({ rows: [] })
    const db: WorkbookRuntimeStoreConnection = {
      run: vi.fn(),
      query,
    }

    await expect(loadWorkbookRuntimeMetadata(db, 'book-3')).resolves.toEqual({
      headRevision: 6,
      calculatedRevision: 5,
      ownerUserId: 'owner-3',
    })

    await acquireWorkbookMutationLock(db, 'book-3')

    expect(query).toHaveBeenLastCalledWith(`SELECT pg_advisory_xact_lock(hashtext($1))`, ['book-3'])
  })

  it('rejects impossible workbook revision metadata from Zero', async () => {
    const query = vi.fn().mockResolvedValueOnce({
      rows: [
        {
          owner_user_id: 'owner-4',
          head_revision: '3',
          calculated_revision: '4',
        },
      ],
    })
    const db: WorkbookRuntimeStoreConnection = {
      run: vi.fn(),
      query,
    }

    await expect(loadWorkbookRuntimeMetadata(db, 'book-4')).rejects.toThrow('Invalid Zero workbook revision order for book-4')
    expect(query).toHaveBeenCalledOnce()
  })

  it('does not trust malformed checkpoint revisions when replaying events', async () => {
    storeFns.loadWorkbookEventRecordsAfter.mockResolvedValueOnce([])
    const checkpoint = createEmptyWorkbookSnapshot('book-5')
    const db: WorkbookRuntimeStoreConnection = {
      run: vi.fn(),
      query: vi
        .fn()
        .mockResolvedValueOnce({
          rows: [
            {
              owner_user_id: 'owner-5',
              head_revision: '2',
              calculated_revision: '1',
            },
          ],
        })
        .mockResolvedValueOnce({
          rows: [
            {
              snapshot: null,
              replica_snapshot: null,
            },
          ],
        })
        .mockResolvedValueOnce({
          rows: [
            {
              revision: '-1',
              payload: checkpoint,
              replica_snapshot: null,
            },
          ],
        }),
    }

    await loadWorkbookState(db, 'book-5')

    expect(storeFns.loadWorkbookEventRecordsAfter).toHaveBeenCalledWith(db, 'book-5', 0)
  })
})
