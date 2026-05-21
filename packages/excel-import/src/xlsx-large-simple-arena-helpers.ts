import { filledUint32Array } from './xlsx-large-simple-array-storage.js'

export const noPoolId = 0xffffffff
export const previewRowCount = 8
export const previewColumnCount = 6
export const previewCellCount = previewRowCount * previewColumnCount

export const maxSpreadsheetColumnCount = 16_384
export const initialSparseIntegerCapacity = 16
const minInt8 = -0x80
const maxInt8 = 0x7f
const minInt16 = -0x8000
const maxInt16 = 0x7fff
const minInt32 = -0x80000000
const maxInt32 = 0x7fffffff

export function isPreviewCell(row: number, column: number): boolean {
  return row >= 0 && row < previewRowCount && column >= 0 && column < previewColumnCount
}

export function previewIndex(row: number, column: number): number {
  return isPreviewCell(row, column) ? row * previewColumnCount + column : -1
}

export function packArenaCellAddress(row: number, column: number): number {
  return row * maxSpreadsheetColumnCount + column
}

export function canStoreLinearCoordinate(width: number, row: number, column: number): boolean {
  if (!Number.isSafeInteger(width) || width <= 0 || row < 0 || column < 0 || column >= width) {
    return false
  }
  const linearCellIndex = row * width + column
  return Number.isSafeInteger(linearCellIndex) && linearCellIndex >= 0 && linearCellIndex <= 0xffffffff
}

export function canStoreInt32Number(value: number): boolean {
  return Number.isInteger(value) && value >= minInt32 && value <= maxInt32 && !Object.is(value, -0)
}

export function canStoreInt8Number(value: number): boolean {
  return Number.isInteger(value) && value >= minInt8 && value <= maxInt8 && !Object.is(value, -0)
}

export function canStoreInt16Number(value: number): boolean {
  return Number.isInteger(value) && value >= minInt16 && value <= maxInt16 && !Object.is(value, -0)
}

export function binarySearchUint32(values: Uint32Array, target: number): number {
  return binarySearchUint32Prefix(values, values.length, target)
}

export function binarySearchUint32Prefix(values: Uint32Array, length: number, target: number): number {
  let low = 0
  let high = length - 1
  while (low <= high) {
    const mid = (low + high) >>> 1
    const value = values[mid] ?? 0
    if (value === target) {
      return mid
    }
    if (value < target) {
      low = mid + 1
    } else {
      high = mid - 1
    }
  }
  return -1
}

export interface ArenaStringIdStorage {
  readonly stringIds: Uint32Array<ArrayBuffer> | undefined
  readonly sparseStringCellIndexes: Uint32Array<ArrayBuffer> | undefined
  readonly sparseStringIds: Uint32Array<ArrayBuffer> | undefined
}

export function compactArenaStringIds(denseStringIds: Uint32Array<ArrayBuffer> | undefined, length: number): ArenaStringIdStorage | null {
  if (!denseStringIds || length === 0) {
    return null
  }
  let retainedCount = 0
  for (let index = 0; index < length; index += 1) {
    if ((denseStringIds[index] ?? noPoolId) !== noPoolId) {
      retainedCount += 1
    }
  }
  if (retainedCount === 0) {
    return { stringIds: undefined, sparseStringCellIndexes: undefined, sparseStringIds: undefined }
  }
  if (retainedCount * 2 >= length) {
    return null
  }
  const indexes = new Uint32Array(retainedCount)
  const ids = new Uint32Array(retainedCount)
  let outputIndex = 0
  for (let index = 0; index < length; index += 1) {
    const stringId = denseStringIds[index] ?? noPoolId
    if (stringId === noPoolId) {
      continue
    }
    indexes[outputIndex] = index
    ids[outputIndex] = stringId
    outputIndex += 1
  }
  return { stringIds: undefined, sparseStringCellIndexes: indexes, sparseStringIds: ids }
}

export function snapshotArenaStringIds(
  denseStringIds: Uint32Array<ArrayBuffer> | undefined,
  sparseIndexes: Uint32Array<ArrayBuffer> | undefined,
  sparseIds: Uint32Array<ArrayBuffer> | undefined,
  length: number,
): Uint32Array<ArrayBuffer> | undefined {
  if (denseStringIds) {
    return denseStringIds.subarray(0, length)
  }
  if (!sparseIndexes || !sparseIds || sparseIndexes.length === 0) {
    return undefined
  }
  const output = filledUint32Array(length, noPoolId)
  for (let index = 0; index < sparseIndexes.length; index += 1) {
    const cellIndex = sparseIndexes[index] ?? -1
    if (cellIndex >= 0 && cellIndex < output.length) {
      output[cellIndex] = sparseIds[index] ?? noPoolId
    }
  }
  return output
}

export function collectArenaIndexesWithCount(count: number, length: number, includeIndex: (index: number) => boolean): Uint32Array {
  const output = new Uint32Array(count)
  let outputIndex = 0
  for (let index = 0; index < length; index += 1) {
    if (!includeIndex(index)) {
      continue
    }
    output[outputIndex] = index
    outputIndex += 1
  }
  return output
}
