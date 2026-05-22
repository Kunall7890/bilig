export interface LinearRowMajorCoordinateCompaction {
  readonly rowRunIndexes: Uint32Array<ArrayBuffer>
  readonly rowRunRows: Uint32Array<ArrayBuffer>
  readonly rowRunStartColumns?: Uint16Array<ArrayBuffer>
  readonly columnPattern?: Uint16Array<ArrayBuffer>
  readonly columns?: Uint16Array<ArrayBuffer>
  readonly runCount: number
}

export function compactLinearRowMajorCoordinates(
  linearCellIndexes: Uint32Array,
  length: number,
  width: number,
): LinearRowMajorCoordinateCompaction | null {
  if (length <= 0 || width <= 0 || linearCellIndexes.length < length) {
    return null
  }
  let runCount = 0
  let previousRow = -1
  let previousColumn = -1
  let nextContiguousColumn = -1
  let columnsAreContiguousByRun = true
  for (let index = 0; index < length; index += 1) {
    const linearCellIndex = linearCellIndexes[index] ?? 0
    const row = Math.floor(linearCellIndex / width)
    const column = linearCellIndex % width
    if (row < previousRow || (row === previousRow && column <= previousColumn)) {
      return null
    }
    if (row !== previousRow) {
      runCount += 1
      previousRow = row
      nextContiguousColumn = column + 1
    } else if (column === nextContiguousColumn) {
      nextContiguousColumn += 1
    } else {
      columnsAreContiguousByRun = false
    }
    previousColumn = column
  }
  if (runCount * 2 >= length) {
    return null
  }
  const rowRunIndexes: Uint32Array<ArrayBuffer> = new Uint32Array(runCount)
  const rowRunRows: Uint32Array<ArrayBuffer> = new Uint32Array(runCount)
  const rowRunStartColumns: Uint16Array<ArrayBuffer> | undefined = columnsAreContiguousByRun ? new Uint16Array(runCount) : undefined
  let outputIndex = 0
  previousRow = -1
  for (let index = 0; index < length; index += 1) {
    const linearCellIndex = linearCellIndexes[index] ?? 0
    const row = Math.floor(linearCellIndex / width)
    const column = linearCellIndex % width
    if (row === previousRow) {
      continue
    }
    rowRunIndexes[outputIndex] = index
    rowRunRows[outputIndex] = row
    if (rowRunStartColumns) {
      rowRunStartColumns[outputIndex] = column
    }
    outputIndex += 1
    previousRow = row
  }
  const columnPattern = columnsAreContiguousByRun
    ? undefined
    : readRepeatedColumnPattern(linearCellIndexes, length, width, rowRunIndexes, runCount)
  const columns = columnsAreContiguousByRun || columnPattern ? undefined : materializeColumns(linearCellIndexes, length, width)
  return {
    rowRunIndexes,
    rowRunRows,
    ...(rowRunStartColumns ? { rowRunStartColumns } : {}),
    ...(columnPattern ? { columnPattern } : {}),
    ...(columns ? { columns } : {}),
    runCount,
  }
}

function materializeColumns(linearCellIndexes: Uint32Array, length: number, width: number): Uint16Array<ArrayBuffer> {
  const columns: Uint16Array<ArrayBuffer> = new Uint16Array(length)
  for (let index = 0; index < length; index += 1) {
    columns[index] = (linearCellIndexes[index] ?? 0) % width
  }
  return columns
}

function readRepeatedColumnPattern(
  linearCellIndexes: Uint32Array,
  length: number,
  width: number,
  rowRunIndexes: Uint32Array,
  runCount: number,
): Uint16Array<ArrayBuffer> | undefined {
  if (runCount === 0) {
    return undefined
  }
  const patternLength = (runCount > 1 ? (rowRunIndexes[1] ?? length) : length) - (rowRunIndexes[0] ?? 0)
  if (patternLength <= 0) {
    return undefined
  }
  const pattern: Uint16Array<ArrayBuffer> = new Uint16Array(patternLength)
  for (let offset = 0; offset < patternLength; offset += 1) {
    pattern[offset] = (linearCellIndexes[offset] ?? 0) % width
  }
  for (let run = 1; run < runCount; run += 1) {
    const start = rowRunIndexes[run] ?? 0
    const end = run + 1 < runCount ? (rowRunIndexes[run + 1] ?? length) : length
    if (end - start !== patternLength) {
      return undefined
    }
    for (let offset = 0; offset < patternLength; offset += 1) {
      if ((linearCellIndexes[start + offset] ?? 0) % width !== pattern[offset]) {
        return undefined
      }
    }
  }
  return pattern
}
