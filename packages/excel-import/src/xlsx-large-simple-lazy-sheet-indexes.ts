export interface LazySheetCellIndexRange {
  readonly start: number
  readonly length: number
}

export type LazySheetCellIndexes = Uint32Array | number | LazySheetCellIndexRange

export function isLazySheetCellIndexRange(indexes: LazySheetCellIndexes): indexes is LazySheetCellIndexRange {
  return typeof indexes !== 'number' && !(indexes instanceof Uint32Array)
}

export function lazySheetCellIndexCount(indexes: LazySheetCellIndexes): number {
  return typeof indexes === 'number' ? indexes : indexes.length
}

export function lazySheetCellArenaIndex(indexes: LazySheetCellIndexes, index: number): number {
  if (typeof indexes === 'number') {
    return index
  }
  if (isLazySheetCellIndexRange(indexes)) {
    return indexes.start + index
  }
  return indexes[index] ?? -1
}
