export const noPoolId = 0xffffffff
export const previewRowCount = 8
export const previewColumnCount = 6
export const previewCellCount = previewRowCount * previewColumnCount

export const maxSpreadsheetColumnCount = 16_384
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

export function binarySearchUint32(values: Uint32Array, target: number): number {
  let low = 0
  let high = values.length - 1
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
