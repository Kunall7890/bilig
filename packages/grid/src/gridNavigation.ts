import { formatAddress } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS, ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { Item, Rectangle } from './gridTypes.js'

export type GridNavigationDirection = 'up' | 'down' | 'left' | 'right'

export interface GridKeyNavigationResolver {
  resolveCurrentRegion(this: void, cell: Item): Rectangle | null
  resolveDataEdge(this: void, cell: Item, direction: GridNavigationDirection): Item | null
}

interface NavigationBounds {
  readonly maxCol: number
  readonly maxRow: number
  readonly minCol: number
  readonly minRow: number
}

interface NavigationIndex {
  readonly bounds: NavigationBounds | null
  readonly colsByRow: ReadonlyMap<number, readonly number[]>
  readonly rowsByCol: ReadonlyMap<number, readonly number[]>
}

function lowerBound(values: readonly number[], target: number): number {
  let low = 0
  let high = values.length
  while (low < high) {
    const mid = low + Math.floor((high - low) / 2)
    if ((values[mid] ?? 0) < target) {
      low = mid + 1
    } else {
      high = mid
    }
  }
  return low
}

function cellHasNavigableContent(snapshot: CellSnapshot): boolean {
  return (
    snapshot.value.tag !== ValueTag.Empty || snapshot.formula !== undefined || (snapshot.input !== undefined && snapshot.input !== null)
  )
}

function pushIndexValue(index: Map<number, number[]>, key: number, value: number): void {
  const values = index.get(key)
  if (values) {
    values.push(value)
    return
  }
  index.set(key, [value])
}

function buildNavigationIndex(engine: GridEngineLike, sheetName: string): NavigationIndex {
  const sheet = engine.workbook.getSheet(sheetName)
  const colsByRow = new Map<number, number[]>()
  const rowsByCol = new Map<number, number[]>()
  let minCol = Number.POSITIVE_INFINITY
  let minRow = Number.POSITIVE_INFINITY
  let maxCol = -1
  let maxRow = -1

  sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
    const snapshot = engine.getCell(sheetName, formatAddress(row, col))
    if (!cellHasNavigableContent(snapshot)) {
      return
    }
    pushIndexValue(colsByRow, row, col)
    pushIndexValue(rowsByCol, col, row)
    minCol = Math.min(minCol, col)
    minRow = Math.min(minRow, row)
    maxCol = Math.max(maxCol, col)
    maxRow = Math.max(maxRow, row)
  })

  colsByRow.forEach((values) => values.sort((a, b) => a - b))
  rowsByCol.forEach((values) => values.sort((a, b) => a - b))

  return {
    bounds:
      maxCol >= 0
        ? {
            maxCol,
            maxRow,
            minCol,
            minRow,
          }
        : null,
    colsByRow,
    rowsByCol,
  }
}

function resolvePositiveDataEdge(values: readonly number[], position: number, limit: number): number {
  const index = lowerBound(values, position)
  if (values[index] === position) {
    const nextIndex = index + 1
    if (values[nextIndex] === position + 1) {
      let cursor = nextIndex
      while (values[cursor + 1] === (values[cursor] ?? 0) + 1) {
        cursor += 1
      }
      return values[cursor] ?? limit
    }
    return values[nextIndex] ?? limit
  }
  return values[index] ?? limit
}

function resolveNegativeDataEdge(values: readonly number[], position: number, limit: number): number {
  const index = lowerBound(values, position)
  if (values[index] === position) {
    const previousIndex = index - 1
    if (values[previousIndex] === position - 1) {
      let cursor = previousIndex
      while (values[cursor - 1] === (values[cursor] ?? 0) - 1) {
        cursor -= 1
      }
      return values[cursor] ?? limit
    }
    return values[previousIndex] ?? limit
  }
  return values[index - 1] ?? limit
}

function listHasValueBetween(values: readonly number[] | undefined, start: number, end: number): boolean {
  if (!values || values.length === 0) {
    return false
  }
  const index = lowerBound(values, start)
  const value = values[index]
  return value !== undefined && value <= end
}

function rowHasContentBetween(index: NavigationIndex, row: number, startCol: number, endCol: number): boolean {
  return listHasValueBetween(index.colsByRow.get(row), startCol, endCol)
}

function colHasContentBetween(index: NavigationIndex, col: number, startRow: number, endRow: number): boolean {
  return listHasValueBetween(index.rowsByCol.get(col), startRow, endRow)
}

function rectangleHasContent(index: NavigationIndex, range: Rectangle): boolean {
  const endRow = range.y + range.height - 1
  const endCol = range.x + range.width - 1
  for (const [row, cols] of index.colsByRow) {
    if (row >= range.y && row <= endRow && listHasValueBetween(cols, range.x, endCol)) {
      return true
    }
  }
  return false
}

function resolveDataEdgeFromIndex(index: NavigationIndex, cell: Item, direction: GridNavigationDirection): Item | null {
  const [col, row] = cell
  switch (direction) {
    case 'up':
      return [col, resolveNegativeDataEdge(index.rowsByCol.get(col) ?? [], row, 0)]
    case 'down':
      return [col, resolvePositiveDataEdge(index.rowsByCol.get(col) ?? [], row, MAX_ROWS - 1)]
    case 'left':
      return [resolveNegativeDataEdge(index.colsByRow.get(row) ?? [], col, 0), row]
    case 'right':
      return [resolvePositiveDataEdge(index.colsByRow.get(row) ?? [], col, MAX_COLS - 1), row]
  }
}

function resolveCurrentRegionFromIndex(index: NavigationIndex, cell: Item): Rectangle | null {
  const bounds = index.bounds
  if (!bounds) {
    return null
  }

  let startCol = cell[0]
  let endCol = cell[0]
  let startRow = cell[1]
  let endRow = cell[1]
  let expanded = true

  while (expanded) {
    expanded = false
    while (startCol > bounds.minCol && colHasContentBetween(index, startCol - 1, startRow, endRow)) {
      startCol -= 1
      expanded = true
    }
    while (endCol < bounds.maxCol && colHasContentBetween(index, endCol + 1, startRow, endRow)) {
      endCol += 1
      expanded = true
    }
    while (startRow > bounds.minRow && rowHasContentBetween(index, startRow - 1, startCol, endCol)) {
      startRow -= 1
      expanded = true
    }
    while (endRow < bounds.maxRow && rowHasContentBetween(index, endRow + 1, startCol, endCol)) {
      endRow += 1
      expanded = true
    }
  }

  const range = {
    x: startCol,
    y: startRow,
    width: endCol - startCol + 1,
    height: endRow - startRow + 1,
  }
  return rectangleHasContent(index, range) ? range : null
}

export function createGridNavigationResolver(input: {
  readonly engine: GridEngineLike
  readonly sheetName: string
}): GridKeyNavigationResolver {
  let index: NavigationIndex | null = null
  const getIndex = (): NavigationIndex => {
    index ??= buildNavigationIndex(input.engine, input.sheetName)
    return index
  }

  return {
    resolveCurrentRegion(cell) {
      return resolveCurrentRegionFromIndex(getIndex(), cell)
    },
    resolveDataEdge(cell, direction) {
      return resolveDataEdgeFromIndex(getIndex(), cell, direction)
    },
  }
}
