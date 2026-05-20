import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import { CellFlags } from '../cell-store.js'
import type { StringPool } from '../string-pool.js'
import type { WorkbookStore } from '../workbook-store.js'

const formulaCacheErrorCodeByText = new Map<string, ErrorCode>([
  ['#DIV/0!', ErrorCode.Div0],
  ['#REF!', ErrorCode.Ref],
  ['#VALUE!', ErrorCode.Value],
  ['#NAME?', ErrorCode.Name],
  ['#N/A', ErrorCode.NA],
  ['#SPILL!', ErrorCode.Spill],
  ['#BLOCKED!', ErrorCode.Blocked],
  ['#NUM!', ErrorCode.Num],
])

export function restoreLiteralCell(
  workbook: WorkbookStore,
  strings: StringPool,
  cellIndex: number,
  value: LiteralInput,
  stringIdCache?: Map<string, number>,
): void {
  const cellStore = workbook.cellStore
  const flags = cellStore.flags[cellIndex] ?? 0
  if (value === null) {
    cellStore.tags[cellIndex] = ValueTag.Empty
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = 0
    cellStore.flags[cellIndex] = flags | CellFlags.AuthoredBlank
  } else if (typeof value === 'number') {
    cellStore.tags[cellIndex] = ValueTag.Number
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = value
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
  } else if (typeof value === 'boolean') {
    cellStore.tags[cellIndex] = ValueTag.Boolean
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = 0
    cellStore.numbers[cellIndex] = value ? 1 : 0
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
  } else {
    let stringId = stringIdCache?.get(value)
    if (stringId === undefined) {
      stringId = strings.intern(value)
      stringIdCache?.set(value, stringId)
    }
    cellStore.tags[cellIndex] = ValueTag.String
    cellStore.errors[cellIndex] = ErrorCode.None
    cellStore.stringIds[cellIndex] = stringId
    cellStore.numbers[cellIndex] = 0
    if ((flags & CellFlags.AuthoredBlank) !== 0) {
      cellStore.flags[cellIndex] = flags & ~CellFlags.AuthoredBlank
    }
  }
  cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  cellStore.onSetValue?.(cellIndex)
}

export function literalToRestoredValue(input: LiteralInput, stringPool: StringPool, stringIdCache: Map<string, number>): CellValue {
  if (input === null) return { tag: ValueTag.Empty }
  if (typeof input === 'number') return { tag: ValueTag.Number, value: input }
  if (typeof input === 'boolean') return { tag: ValueTag.Boolean, value: input }
  let stringId = stringIdCache.get(input)
  if (stringId === undefined) {
    stringId = stringPool.intern(input)
    stringIdCache.set(input, stringId)
  }
  return { tag: ValueTag.String, value: input, stringId }
}

export function formulaCachedLiteralToRestoredValue(
  input: LiteralInput,
  stringPool: StringPool,
  stringIdCache: Map<string, number>,
): CellValue {
  if (typeof input === 'string') {
    const errorCode = formulaCacheErrorCodeByText.get(input)
    if (errorCode !== undefined) {
      return { tag: ValueTag.Error, code: errorCode }
    }
  }
  return literalToRestoredValue(input, stringPool, stringIdCache)
}
