import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import type { U32 } from '../runtime-state.js'

const TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY = '__biligTrackedPhysicalSheetId'
const TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY = '__biligTrackedPhysicalSortedSliceSplit'

export function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${String(value)}`)
}

export function normalizeRange(range: CellRangeRef): CellRangeRef & {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    ...range,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

export function rangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  const a = normalizeRange(left)
  const b = normalizeRange(right)
  return !(a.sheetName !== b.sheetName || a.endRow < b.startRow || b.endRow < a.startRow || a.endCol < b.startCol || b.endCol < a.startCol)
}

export function cellRange(sheetName: string, address: string): CellRangeRef {
  return {
    sheetName,
    startAddress: address,
    endAddress: address,
  }
}

export function throwProtectionBlocked(message: string): never {
  throw new Error(`Workbook protection blocks this change: ${message}`)
}

export function mergeChangedCellIndices(base: readonly number[] | U32, extras: readonly number[] | U32): U32 {
  if (base.length === 0) {
    return extras instanceof Uint32Array ? extras : Uint32Array.from(extras)
  }
  if (extras.length === 0) {
    return base instanceof Uint32Array ? base : Uint32Array.from(base)
  }
  if (base.length === 1 && extras.length === 1) {
    const baseCellIndex = base[0]!
    const extraCellIndex = extras[0]!
    return baseCellIndex === extraCellIndex ? Uint32Array.of(baseCellIndex) : Uint32Array.of(baseCellIndex, extraCellIndex)
  }
  const merged = new Set<number>()
  for (let index = 0; index < base.length; index += 1) {
    merged.add(base[index]!)
  }
  for (let index = 0; index < extras.length; index += 1) {
    merged.add(extras[index]!)
  }
  return Uint32Array.from(merged)
}

export function tagTrustedPhysicalTrackedChanges(changed: U32, sheetId: number, sortedSliceSplit: number): void {
  Reflect.set(changed, TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY, sheetId)
  Reflect.set(changed, TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY, sortedSliceSplit)
}

export function makeExistingNumericMutationResult(
  changedCellIndices: U32,
  explicitChangedCount: number,
): EngineExistingNumericCellMutationResult {
  return { changedCellIndices, explicitChangedCount }
}

export function makeCompactExistingNumericMutationResult(
  firstChangedCellIndex: number,
  secondChangedCellIndex: number | undefined,
  explicitChangedCount: number,
  secondChangedNumericValue?: number,
  secondChangedPosition?: { readonly row: number; readonly col: number },
): EngineExistingNumericCellMutationResult {
  return secondChangedCellIndex === undefined
    ? { firstChangedCellIndex, changedCellCount: 1, explicitChangedCount }
    : {
        firstChangedCellIndex,
        secondChangedCellIndex,
        changedCellCount: 2,
        explicitChangedCount,
        ...(secondChangedNumericValue === undefined ? {} : { secondChangedNumericValue }),
        ...(secondChangedPosition === undefined
          ? {}
          : { secondChangedRow: secondChangedPosition.row, secondChangedCol: secondChangedPosition.col }),
      }
}
