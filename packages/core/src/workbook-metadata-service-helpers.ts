import { Effect } from 'effect'
import type { CellRangeRef } from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import { cloneDataValidationRecord } from './workbook-metadata-records.js'
import type { WorkbookDataValidationRecord, WorkbookFilterRecord, WorkbookMergeRangeRecord } from './workbook-metadata-types.js'
import { canonicalWorkbookRangeRef } from './workbook-range-records.js'
import { WorkbookMetadataError } from './workbook-metadata-service-contract.js'

export function renameDataValidationSourceSheet(
  record: WorkbookDataValidationRecord,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDataValidationRecord {
  const cloned = cloneDataValidationRecord(record)
  if (cloned.rule.kind !== 'list' || !cloned.rule.source) {
    return cloned
  }
  switch (cloned.rule.source.kind) {
    case 'cell-ref':
    case 'range-ref':
      if (cloned.rule.source.sheetName === oldSheetName) {
        cloned.rule.source.sheetName = newSheetName
      }
      return cloned
    case 'named-range':
    case 'structured-ref':
      return cloned
  }
  return cloned
}

export function assertMergeRangesDoNotOverlap(ranges: readonly WorkbookMergeRangeRecord[]): void {
  const normalized = ranges.map(normalizeMergeRangeForOverlap).toSorted((left, right) => left.startRow - right.startRow)
  const active: NormalizedMergeRangeRecord[] = []
  for (const range of normalized) {
    for (let index = active.length - 1; index >= 0; index -= 1) {
      if (active[index]!.endRow < range.startRow) {
        active.splice(index, 1)
      }
    }
    if (active.some((entry) => mergeRangesOverlap(entry, range))) {
      throw new Error('Merged ranges cannot overlap')
    }
    active.push(range)
  }
}

export function canonicalWorkbookFilterRange(range: WorkbookFilterRecord['range']): WorkbookFilterRecord['range'] {
  const normalized = canonicalWorkbookRangeRef(range)
  const criteria = range.criteria?.length ? structuredClone(range.criteria) : undefined
  return criteria ? { ...normalized, criteria } : normalized
}

export function canonicalWorkbookRangeOnSheet(sheetName: string, range: CellRangeRef): CellRangeRef {
  return canonicalWorkbookRangeRef({ ...range, sheetName })
}

export function canonicalWorkbookFilterRangeOnSheet(
  sheetName: string,
  range: WorkbookFilterRecord['range'],
): WorkbookFilterRecord['range'] {
  return {
    ...canonicalWorkbookFilterRange({ ...range, sheetName }),
    sheetName,
  }
}

export function metadataEffect<Success>(message: string, run: () => Success): Effect.Effect<Success, WorkbookMetadataError> {
  return Effect.try({
    try: run,
    catch: (cause) =>
      new WorkbookMetadataError({
        message: metadataErrorMessage(message, cause),
        cause,
      }),
  })
}

export function normalizeMetadataKey(key: string): string {
  const trimmed = key.trim()
  if (trimmed.length === 0) {
    throw new Error('Workbook metadata keys must be non-empty')
  }
  return trimmed
}

interface NormalizedMergeRangeRecord {
  readonly record: WorkbookMergeRangeRecord
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

function normalizeMergeRangeForOverlap(record: WorkbookMergeRangeRecord): NormalizedMergeRangeRecord {
  const start = parseCellAddress(record.startAddress, record.sheetName)
  const end = parseCellAddress(record.endAddress, record.sheetName)
  return {
    record,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

function mergeRangesOverlap(left: NormalizedMergeRangeRecord, right: NormalizedMergeRangeRecord): boolean {
  return !(left.endRow < right.startRow || right.endRow < left.startRow || left.endCol < right.startCol || right.endCol < left.startCol)
}

function metadataErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}
