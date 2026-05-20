import { ValueTag, type CellValue } from '@bilig/protocol'
import { areCellValuesEqual } from '../../engine-value-utils.js'
import type { U32 } from '../runtime-state.js'

export function filterSkippedCachedFormulaCells(ordered: U32, orderedCount: number, skipped: ReadonlySet<number> | undefined): U32 {
  if (!skipped || skipped.size === 0) {
    return ordered.length === orderedCount ? ordered : ordered.subarray(0, orderedCount)
  }
  const filtered = new Uint32Array(orderedCount)
  let filteredCount = 0
  for (let index = 0; index < orderedCount; index += 1) {
    const cellIndex = ordered[index]
    if (cellIndex === undefined || skipped.has(cellIndex)) {
      continue
    }
    filtered[filteredCount] = cellIndex
    filteredCount += 1
  }
  return filtered.subarray(0, filteredCount)
}

export function areDifferentialCellValuesEqual(left: CellValue, right: CellValue): boolean {
  if (areCellValuesEqual(left, right)) {
    return true
  }
  if (left.tag !== ValueTag.Number || right.tag !== ValueTag.Number) {
    return false
  }
  if (!Number.isFinite(left.value) || !Number.isFinite(right.value)) {
    return false
  }
  const tolerance = 1e-12 * Math.max(1, Math.abs(left.value), Math.abs(right.value))
  return Math.abs(left.value - right.value) <= tolerance
}
