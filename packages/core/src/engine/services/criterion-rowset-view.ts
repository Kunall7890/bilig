export interface CriterionRowSetView {
  readonly absoluteRows: Uint32Array
  readonly startIndex: number
  readonly endIndex: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly cardinality: number
  readonly forEachOffset: (fn: (offset: number) => void) => void
}

function lowerBound(rows: Uint32Array, value: number): number {
  let low = 0
  let high = rows.length
  while (low < high) {
    const mid = (low + high) >>> 1
    if (rows[mid]! < value) {
      low = mid + 1
    } else {
      high = mid
    }
  }
  return low
}

export function sliceAbsoluteRowsToRangeView(rows: Uint32Array, rowStart: number, rowEnd: number): CriterionRowSetView {
  const startIndex = lowerBound(rows, rowStart)
  const endIndex = lowerBound(rows, rowEnd + 1)
  const cardinality = Math.max(0, endIndex - startIndex)
  return {
    absoluteRows: rows,
    startIndex,
    endIndex,
    rowStart,
    rowEnd,
    cardinality,
    forEachOffset(fn) {
      for (let index = startIndex; index < endIndex; index += 1) {
        fn(rows[index]! - rowStart)
      }
    },
  }
}
