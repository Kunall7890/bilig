import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import {
  createMemoryWorkbookLocalStoreFactory,
  type WorkbookLocalAuthoritativeBase,
  type WorkbookLocalAuthoritativeDelta,
  type WorkbookLocalMutationRecord,
  type WorkbookLocalProjectionOverlay,
} from '../index.js'

function createBase(value: number): WorkbookLocalAuthoritativeBase {
  return {
    sheets: [
      {
        sheetId: 1,
        name: 'Sheet1',
        sortOrder: 0,
        freezeRows: 0,
        freezeCols: 0,
      },
    ],
    cellInputs: [
      {
        sheetId: 1,
        sheetName: 'Sheet1',
        address: 'A1',
        rowNum: 0,
        colNum: 0,
        input: value,
        formula: undefined,
        format: undefined,
      },
    ],
    cellRenders: [
      {
        sheetId: 1,
        sheetName: 'Sheet1',
        address: 'A1',
        rowNum: 0,
        colNum: 0,
        value: { tag: ValueTag.Number, value },
        flags: 0,
        version: 1,
        styleId: undefined,
        numberFormatId: undefined,
      },
    ],
    rowAxisEntries: [],
    columnAxisEntries: [],
    styles: [],
  }
}

function createOverlay(value: number): WorkbookLocalProjectionOverlay {
  return {
    cells: [
      {
        sheetId: 1,
        sheetName: 'Sheet1',
        address: 'A1',
        rowNum: 0,
        colNum: 0,
        value: { tag: ValueTag.Number, value },
        flags: 0,
        version: 2,
        input: value,
        formula: undefined,
        format: undefined,
        styleId: undefined,
        numberFormatId: undefined,
      },
    ],
    rowAxisEntries: [],
    columnAxisEntries: [],
    styles: [],
  }
}

function createDelta(value: number): WorkbookLocalAuthoritativeDelta {
  return {
    replaceAll: true,
    replacedSheetIds: [],
    base: createBase(value),
  }
}

function createMutation(overrides: Partial<WorkbookLocalMutationRecord> = {}): WorkbookLocalMutationRecord {
  return {
    id: 'memory-doc:pending:1',
    localSeq: 1,
    baseRevision: 0,
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 17],
    enqueuedAtUnixMs: 100,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: 'local',
    ...overrides,
  }
}

describe('memory workbook local store', () => {
  it('persists runtime state and normalized projection data across reopen', async () => {
    const factory = createMemoryWorkbookLocalStoreFactory()
    const store = await factory.open('memory-doc')

    await store.persistProjectionState({
      state: {
        snapshot: { version: 1, workbook: { name: 'memory-doc' }, sheets: [] },
        replica: { replica: { id: 'seed', clock: 0 }, entityVersions: [], sheetDeleteVersions: [] },
        authoritativeRevision: 7,
        appliedPendingLocalSeq: 3,
      },
      authoritativeBase: createBase(11),
      projectionOverlay: createOverlay(17),
    })
    store.close()

    const reopened = await factory.open('memory-doc')
    await expect(reopened.loadBootstrapState()).resolves.toEqual({
      workbookName: 'memory-doc',
      sheetNames: ['Sheet1'],
      materializedCellCount: 1,
      authoritativeRevision: 7,
      appliedPendingLocalSeq: 3,
    })
    expect(await reopened.loadState()).toMatchObject({
      authoritativeRevision: 7,
      appliedPendingLocalSeq: 3,
    })
    expect(
      reopened.readViewportProjection('Sheet1', {
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      }),
    ).toMatchObject({
      cells: [
        {
          snapshot: {
            sheetName: 'Sheet1',
            address: 'A1',
            value: { tag: ValueTag.Number, value: 17 },
          },
        },
      ],
    })
  })

  it('keeps acked mutations in the journal while filtering them from the active list', async () => {
    const factory = createMemoryWorkbookLocalStoreFactory()
    const store = await factory.open('memory-doc')
    const local = createMutation()
    const acked = {
      ...local,
      submittedAtUnixMs: 120,
      lastAttemptedAtUnixMs: 120,
      ackedAtUnixMs: 180,
      attemptCount: 1,
      status: 'acked' as const,
    }

    await store.appendPendingMutation(local)
    await store.updatePendingMutation(acked)
    store.close()

    const reopened = await factory.open('memory-doc')
    await expect(reopened.listPendingMutations()).resolves.toEqual([])
    await expect(reopened.listMutationJournalEntries()).resolves.toEqual([acked])
  })

  it('acks absorbed pending mutations while ingesting authoritative deltas', async () => {
    const factory = createMemoryWorkbookLocalStoreFactory()
    const store = await factory.open('memory-doc')
    const failed = createMutation({
      submittedAtUnixMs: 120,
      lastAttemptedAtUnixMs: 130,
      failedAtUnixMs: 140,
      attemptCount: 2,
      failureMessage: 'transient apply failure',
      status: 'failed',
    })

    await store.appendPendingMutation(failed)
    await store.ingestAuthoritativeDelta({
      state: {
        snapshot: { version: 1, workbook: { name: 'memory-doc' }, sheets: [] },
        replica: { replica: { id: 'seed', clock: 1 }, entityVersions: [], sheetDeleteVersions: [] },
        authoritativeRevision: 8,
        appliedPendingLocalSeq: failed.localSeq,
      },
      authoritativeDelta: createDelta(19),
      projectionOverlay: createOverlay(19),
      removePendingMutationIds: [failed.id],
    })

    await expect(store.listPendingMutations()).resolves.toEqual([])
    await expect(store.listMutationJournalEntries()).resolves.toEqual([
      {
        ...failed,
        ackedAtUnixMs: expect.any(Number),
        failedAtUnixMs: null,
        failureMessage: null,
        status: 'acked',
      },
    ])
    await expect(store.loadState()).resolves.toMatchObject({
      authoritativeRevision: 8,
      appliedPendingLocalSeq: failed.localSeq,
    })
  })

  it('rejects unsafe runtime sequence numbers before writing local state', async () => {
    const factory = createMemoryWorkbookLocalStoreFactory()
    const store = await factory.open('memory-doc')

    await expect(
      store.persistProjectionState({
        state: {
          snapshot: { version: 1, workbook: { name: 'memory-doc' }, sheets: [] },
          replica: { replica: { id: 'seed', clock: 0 }, entityVersions: [], sheetDeleteVersions: [] },
          authoritativeRevision: Number.MAX_SAFE_INTEGER + 1,
          appliedPendingLocalSeq: 3,
        },
        authoritativeBase: createBase(11),
        projectionOverlay: createOverlay(17),
      }),
    ).rejects.toThrow('Invalid workbook runtime state')

    await expect(store.loadState()).resolves.toBeNull()
    expect(
      store.readViewportProjection('Sheet1', {
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      }),
    ).toBeNull()
  })

  it('rejects unsafe pending mutation counters before writing the journal', async () => {
    const factory = createMemoryWorkbookLocalStoreFactory()
    const store = await factory.open('memory-doc')

    await expect(store.appendPendingMutation(createMutation({ localSeq: Number.MAX_SAFE_INTEGER + 1 }))).rejects.toThrow(
      'Invalid workbook local mutation record',
    )
    await expect(store.appendPendingMutation(createMutation({ attemptCount: Number.NaN }))).rejects.toThrow(
      'Invalid workbook local mutation record',
    )

    await expect(store.listMutationJournalEntries()).resolves.toEqual([])
  })
})
