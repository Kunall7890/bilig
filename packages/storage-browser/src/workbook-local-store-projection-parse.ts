import type { SqlValue } from '@sqlite.org/sqlite-wasm'
import {
  ErrorCode,
  sanitizeCellStyleRecord,
  ValueTag,
  type CellSnapshot,
  type CellStyleRecord,
  type LiteralInput,
  type WorkbookAxisEntrySnapshot,
} from '@bilig/protocol'
import type { WorkbookLocalViewportCell } from './workbook-local-base.js'

export interface ViewportBounds {
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'boolean' || typeof value === 'string' || isFiniteNumber(value)
}

export function isValidViewportBounds(viewport: ViewportBounds): boolean {
  return (
    isSafeNonNegativeInteger(viewport.rowStart) &&
    isSafeNonNegativeInteger(viewport.rowEnd) &&
    isSafeNonNegativeInteger(viewport.colStart) &&
    isSafeNonNegativeInteger(viewport.colEnd) &&
    viewport.rowStart <= viewport.rowEnd &&
    viewport.colStart <= viewport.colEnd
  )
}

function isErrorCode(value: unknown): value is ErrorCode {
  return (
    value === ErrorCode.None ||
    value === ErrorCode.Div0 ||
    value === ErrorCode.Ref ||
    value === ErrorCode.Value ||
    value === ErrorCode.Name ||
    value === ErrorCode.NA ||
    value === ErrorCode.Cycle ||
    value === ErrorCode.Spill ||
    value === ErrorCode.Blocked
  )
}

function parseCellSnapshotValue(value: unknown): CellSnapshot['value'] | null {
  if (!isRecord(value)) {
    return null
  }
  switch (value['tag']) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return isFiniteNumber(value['value']) ? { tag: ValueTag.Number, value: value['value'] } : null
    case ValueTag.Boolean:
      return typeof value['value'] === 'boolean' ? { tag: ValueTag.Boolean, value: value['value'] } : null
    case ValueTag.String:
      return typeof value['value'] === 'string' && isSafeNonNegativeInteger(value['stringId'])
        ? {
            tag: ValueTag.String,
            value: value['value'],
            stringId: value['stringId'],
          }
        : null
    case ValueTag.Error:
      return isErrorCode(value['code']) ? { tag: ValueTag.Error, code: value['code'] } : null
    default:
      return null
  }
}

export function parseViewportCellFromRow(row: Record<string, SqlValue>): WorkbookLocalViewportCell | null {
  const address = row['address']
  const sheetName = row['sheetName']
  const rowNum = row['rowNum']
  const colNum = row['colNum']
  const valueJson = row['valueJson']
  const flags = row['flags']
  const version = row['version']
  if (
    typeof address !== 'string' ||
    typeof sheetName !== 'string' ||
    !isSafeNonNegativeInteger(rowNum) ||
    !isSafeNonNegativeInteger(colNum) ||
    typeof valueJson !== 'string' ||
    !isSafeNonNegativeInteger(flags) ||
    !isSafeNonNegativeInteger(version)
  ) {
    return null
  }
  try {
    const parsedValue = parseCellSnapshotValue(JSON.parse(valueJson) as unknown)
    if (!parsedValue) {
      return null
    }
    const snapshot: CellSnapshot = {
      sheetName,
      address,
      value: parsedValue,
      flags,
      version,
    }
    const inputJson = row['inputJson']
    if (typeof inputJson === 'string') {
      const parsedInput = JSON.parse(inputJson) as unknown
      if (isLiteralInput(parsedInput)) {
        snapshot.input = parsedInput
      }
    }
    if (typeof row['formula'] === 'string') {
      snapshot.formula = row['formula']
    }
    if (typeof row['format'] === 'string') {
      snapshot.format = row['format']
    }
    if (typeof row['styleId'] === 'string') {
      snapshot.styleId = row['styleId']
    }
    if (typeof row['numberFormatId'] === 'string') {
      snapshot.numberFormatId = row['numberFormatId']
    }
    return {
      row: rowNum,
      col: colNum,
      snapshot,
    }
  } catch {
    return null
  }
}

export function parseAxisEntrySnapshot(row: Record<string, SqlValue>): WorkbookAxisEntrySnapshot | null {
  const id = row['id']
  const entryIndex = row['entryIndex']
  if (typeof id !== 'string' || !isSafeNonNegativeInteger(entryIndex)) {
    return null
  }
  const entry: WorkbookAxisEntrySnapshot = {
    id,
    index: entryIndex,
  }
  if (isSafeNonNegativeInteger(row['size'])) {
    entry.size = row['size']
  }
  if (row['hidden'] === 0 || row['hidden'] === 1) {
    entry.hidden = row['hidden'] === 1
  } else if (typeof row['hidden'] === 'boolean') {
    entry.hidden = row['hidden']
  }
  return entry
}

export function parseCellStyleRecord(row: Record<string, SqlValue>): CellStyleRecord | null {
  const id = row['id']
  const recordJson = row['recordJson']
  if (typeof id !== 'string' || typeof recordJson !== 'string') {
    return null
  }
  try {
    const parsed = JSON.parse(recordJson) as unknown
    return sanitizeCellStyleRecord(id, parsed)
  } catch {
    return null
  }
}
