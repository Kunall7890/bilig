import type { CellRangeRef } from '@bilig/protocol'
import { cloneFilterRecord, cloneSortKeyRecord, cloneSortRecord, filterKey, sortKey } from './workbook-metadata-records.js'
import { canonicalWorkbookFilterRangeOnSheet, canonicalWorkbookRangeOnSheet } from './workbook-metadata-service-helpers.js'
import type { WorkbookFilterRecord, WorkbookMetadataRecord, WorkbookSortKeyRecord, WorkbookSortRecord } from './workbook-metadata-types.js'

export function setWorkbookFilterRecord(
  metadata: WorkbookMetadataRecord,
  sheetName: string,
  range: WorkbookFilterRecord['range'],
): WorkbookFilterRecord {
  const storedRange = canonicalWorkbookFilterRangeOnSheet(sheetName, range)
  const record: WorkbookFilterRecord = { sheetName, range: storedRange }
  metadata.filters.set(filterKey(sheetName, storedRange), record)
  return cloneFilterRecord(record)
}

export function getWorkbookFilterRecord(
  metadata: WorkbookMetadataRecord,
  sheetName: string,
  range: CellRangeRef,
): WorkbookFilterRecord | undefined {
  const record = metadata.filters.get(filterKey(sheetName, range))
  return record ? cloneFilterRecord(record) : undefined
}

export function deleteWorkbookFilterRecord(metadata: WorkbookMetadataRecord, sheetName: string, range: CellRangeRef): boolean {
  return metadata.filters.delete(filterKey(sheetName, range))
}

export function listWorkbookFilterRecords(metadata: WorkbookMetadataRecord, sheetName: string): WorkbookFilterRecord[] {
  return [...metadata.filters.values()]
    .filter((record) => record.sheetName === sheetName)
    .toSorted((left, right) => filterKey(left.sheetName, left.range).localeCompare(filterKey(right.sheetName, right.range)))
    .map(cloneFilterRecord)
}

export function setWorkbookSortRecord(
  metadata: WorkbookMetadataRecord,
  sheetName: string,
  range: CellRangeRef,
  keys: readonly WorkbookSortKeyRecord[],
): WorkbookSortRecord {
  const storedRange = canonicalWorkbookRangeOnSheet(sheetName, range)
  const record: WorkbookSortRecord = {
    sheetName,
    range: storedRange,
    keys: keys.map(cloneSortKeyRecord),
  }
  metadata.sorts.set(sortKey(sheetName, storedRange), record)
  return cloneSortRecord(record)
}

export function getWorkbookSortRecord(
  metadata: WorkbookMetadataRecord,
  sheetName: string,
  range: CellRangeRef,
): WorkbookSortRecord | undefined {
  const record = metadata.sorts.get(sortKey(sheetName, range))
  return record ? cloneSortRecord(record) : undefined
}

export function deleteWorkbookSortRecord(metadata: WorkbookMetadataRecord, sheetName: string, range: CellRangeRef): boolean {
  return metadata.sorts.delete(sortKey(sheetName, range))
}

export function listWorkbookSortRecords(metadata: WorkbookMetadataRecord, sheetName: string): WorkbookSortRecord[] {
  return [...metadata.sorts.values()]
    .filter((record) => record.sheetName === sheetName)
    .toSorted((left, right) => sortKey(left.sheetName, left.range).localeCompare(sortKey(right.sheetName, right.range)))
    .map(cloneSortRecord)
}
