import { describe, expect, it } from 'vitest'
import type { WorkbookAxisEntryRecord } from '../workbook-metadata-types.js'
import {
  getAxisMetadataRecord,
  listAxisEntries,
  materializeAxisEntries,
  moveAxisEntries,
  snapshotAxisEntriesInRange,
  spliceAxisEntries,
  syncAxisMetadataBucket,
} from '../workbook-axis-records.js'

function createEntryFactory(prefix: string) {
  let nextId = 1
  return (): WorkbookAxisEntryRecord => ({
    id: `${prefix}-${nextId++}`,
    size: null,
    hidden: null,
  })
}

describe('workbook axis records', () => {
  it('materializes sparse entries and snapshots only defined records', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = []
    const createEntry = createEntryFactory('row')

    expect(materializeAxisEntries(entries, 1, 2, createEntry)).toEqual([
      { id: 'row-1', index: 1 },
      { id: 'row-2', index: 2 },
    ])

    entries[1]!.size = 120
    entries[2]!.hidden = true

    expect(listAxisEntries(entries)).toEqual([
      { id: 'row-1', index: 1, size: 120 },
      { id: 'row-2', index: 2, hidden: true },
    ])
    expect(snapshotAxisEntriesInRange(entries, 0, 4)).toEqual([
      { id: 'row-1', index: 1, size: 120 },
      { id: 'row-2', index: 2, hidden: true },
    ])
  })

  it('splices and moves sparse axis entries while preserving provided ids', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = []
    const createEntry = createEntryFactory('column')

    expect(spliceAxisEntries(entries, 0, 0, 2, createEntry, [{ id: 'column-existing', index: 1, size: 90, hidden: false }])).toEqual([])
    expect(listAxisEntries(entries)).toEqual([{ id: 'column-existing', index: 1, size: 90, hidden: false }])

    moveAxisEntries(entries, 1, 1, 0, createEntry)
    expect(listAxisEntries(entries)).toEqual([{ id: 'column-existing', index: 0, size: 90, hidden: false }])
  })

  it('keeps completely sparse default inserts sparse', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = []
    const createEntry = createEntryFactory('column')

    expect(spliceAxisEntries(entries, 1, 0, 2, createEntry)).toEqual([])

    expect(entries.length).toBe(0)
    expect(listAxisEntries(entries)).toEqual([])
  })

  it('splices generated axis entries without requiring provided snapshots', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = [{ id: 'column-existing', size: null, hidden: null }]
    const createEntry = createEntryFactory('column')

    expect(spliceAxisEntries(entries, 1, 0, 2, createEntry)).toEqual([])

    expect(listAxisEntries(entries)).toEqual([
      { id: 'column-existing', index: 0 },
      { id: 'column-1', index: 1 },
      { id: 'column-2', index: 2 },
    ])
  })

  it('uses the single-entry splice path for generated and provided inserts', () => {
    const generatedEntries: Array<WorkbookAxisEntryRecord | undefined> = [
      { id: 'column-existing', size: null, hidden: null },
      { id: 'column-tail', size: null, hidden: null },
    ]
    const providedEntries: Array<WorkbookAxisEntryRecord | undefined> = [
      { id: 'column-existing', size: null, hidden: null },
      { id: 'column-tail', size: null, hidden: null },
    ]
    const createEntry = createEntryFactory('column')

    expect(spliceAxisEntries(generatedEntries, 1, 0, 1, createEntry)).toEqual([])
    expect(spliceAxisEntries(providedEntries, 1, 0, 1, createEntry, [{ id: 'column-provided', index: 1 }])).toEqual([])

    expect(listAxisEntries(generatedEntries)).toEqual([
      { id: 'column-existing', index: 0 },
      { id: 'column-1', index: 1 },
      { id: 'column-tail', index: 2 },
    ])
    expect(listAxisEntries(providedEntries)).toEqual([
      { id: 'column-existing', index: 0 },
      { id: 'column-provided', index: 1 },
      { id: 'column-tail', index: 2 },
    ])
  })

  it('coalesces contiguous metadata ranges and rejects mixed ranges', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = [
      { id: 'row-1', size: 30, hidden: false, filtered: null },
      { id: 'row-2', size: 30, hidden: false, filtered: null },
      undefined,
      { id: 'row-3', size: 30, hidden: true, filtered: null },
    ]
    const bucket = new Map()

    syncAxisMetadataBucket(bucket, 'Sheet1', entries)

    expect([...bucket.values()].toSorted((left, right) => left.start - right.start)).toEqual([
      { sheetName: 'Sheet1', start: 0, count: 2, size: 30, hidden: false, filtered: null },
      { sheetName: 'Sheet1', start: 3, count: 1, size: 30, hidden: true, filtered: null },
    ])
    expect(getAxisMetadataRecord(entries, 'Sheet1', 0, 2)).toEqual({
      sheetName: 'Sheet1',
      start: 0,
      count: 2,
      size: 30,
      hidden: false,
      filtered: null,
    })
    expect(getAxisMetadataRecord(entries, 'Sheet1', 0, 4)).toBeUndefined()
  })
})
