import type { CellRangeRef, LiteralInput, WorkbookSnapshot, WorkbookSortSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'

export interface SpreadsheetEngineSortRangeOptions {
  readonly header?: boolean
}

interface SortBounds {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

interface SortCell {
  readonly formula?: string
  readonly format?: string
  readonly value?: LiteralInput
}

interface SortRow {
  readonly row: number
  readonly originalIndex: number
}

export function buildSortRangeOps(
  snapshot: WorkbookSnapshot,
  sheetName: string,
  range: CellRangeRef,
  keys: readonly WorkbookSortSnapshot['keys'][number][],
  options: SpreadsheetEngineSortRangeOptions = {},
): EngineOp[] {
  if (keys.length === 0) {
    throw new Error('sortRange requires at least one sort key')
  }

  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  if (!sheet) {
    throw new Error(`Unknown sheet: ${sheetName}`)
  }

  const bounds = sortBounds(sheetName, range)
  const bodyStartRow = options.header === true ? bounds.startRow + 1 : bounds.startRow
  if (bodyStartRow > bounds.endRow) {
    return []
  }

  const cellsByAddress = new Map<string, SortCell>()
  for (const cell of sheet.cells) {
    cellsByAddress.set(cell.address, {
      ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
      ...(cell.value !== undefined ? { value: cell.value } : {}),
      ...(cell.format !== undefined ? { format: cell.format } : {}),
    })
  }

  const keyColumns = keys.map((key) => parseCellAddress(key.keyAddress, sheetName).col)
  const rows = Array.from({ length: bounds.endRow - bodyStartRow + 1 }, (_, index): SortRow => {
    return { row: bodyStartRow + index, originalIndex: index }
  })
  const sortedRows = rows.toSorted((left, right) => compareSortRows(cellsByAddress, keyColumns, keys, left, right))

  const ops: EngineOp[] = []
  for (let targetOffset = 0; targetOffset < sortedRows.length; targetOffset += 1) {
    const sourceRow = sortedRows[targetOffset]!.row
    const targetRow = bodyStartRow + targetOffset
    for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
      const source = cellsByAddress.get(formatAddress(sourceRow, col))
      const targetAddress = formatAddress(targetRow, col)
      const existing = cellsByAddress.get(targetAddress)
      appendCellReplacementOps(ops, sheetName, targetAddress, existing, source)
    }
  }

  const metadataRange: CellRangeRef = {
    sheetName,
    startAddress: formatAddress(bodyStartRow, bounds.startCol),
    endAddress: formatAddress(bounds.endRow, bounds.endCol),
  }
  ops.push({
    kind: 'setSort',
    sheetName,
    range: metadataRange,
    keys: keys.map((key) => ({ keyAddress: key.keyAddress, direction: key.direction })),
  })
  return ops
}

function sortBounds(sheetName: string, range: CellRangeRef): SortBounds {
  const start = parseCellAddress(range.startAddress, sheetName)
  const end = parseCellAddress(range.endAddress, sheetName)
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

function compareSortRows(
  cellsByAddress: ReadonlyMap<string, SortCell>,
  keyColumns: readonly number[],
  keys: readonly WorkbookSortSnapshot['keys'][number][],
  left: SortRow,
  right: SortRow,
): number {
  for (let index = 0; index < keys.length; index += 1) {
    const key = keys[index]!
    const col = keyColumns[index]!
    const leftValue = cellsByAddress.get(formatAddress(left.row, col))?.value
    const rightValue = cellsByAddress.get(formatAddress(right.row, col))?.value
    const compared = compareSortValues(leftValue, rightValue, key.direction)
    if (compared !== 0) {
      return compared
    }
  }
  return left.originalIndex - right.originalIndex
}

function compareSortValues(left: LiteralInput | undefined, right: LiteralInput | undefined, direction: 'asc' | 'desc'): number {
  const leftBlank = left === undefined || left === null
  const rightBlank = right === undefined || right === null
  if (leftBlank || rightBlank) {
    if (leftBlank === rightBlank) {
      return 0
    }
    return leftBlank ? 1 : -1
  }

  const leftRank = sortValueRank(left)
  const rightRank = sortValueRank(right)
  const compared =
    leftRank !== rightRank
      ? leftRank - rightRank
      : typeof left === 'number' && typeof right === 'number'
        ? left - right
        : typeof left === 'boolean' && typeof right === 'boolean'
          ? Number(left) - Number(right)
          : String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: 'base' })
  return direction === 'desc' ? -compared : compared
}

function sortValueRank(value: LiteralInput | undefined): number {
  if (value === undefined || value === null) {
    return 3
  }
  if (typeof value === 'number') {
    return 0
  }
  if (typeof value === 'string') {
    return 1
  }
  if (typeof value === 'boolean') {
    return 2
  }
  return 3
}

function appendCellReplacementOps(
  ops: EngineOp[],
  sheetName: string,
  address: string,
  existing: SortCell | undefined,
  source: SortCell | undefined,
): void {
  if (!source) {
    if (existing) {
      ops.push({ kind: 'clearCell', sheetName, address })
      if (existing.format !== undefined) {
        ops.push({ kind: 'setCellFormat', sheetName, address, format: null })
      }
    }
    return
  }

  if (source.formula !== undefined) {
    if (existing?.formula !== source.formula) {
      ops.push({ kind: 'setCellFormula', sheetName, address, formula: source.formula })
    }
  } else if (source.value !== undefined) {
    if (existing?.formula !== undefined || existing?.value !== source.value) {
      ops.push({ kind: 'setCellValue', sheetName, address, value: source.value })
    }
  } else if (existing?.formula !== undefined || existing?.value !== undefined) {
    ops.push({ kind: 'clearCell', sheetName, address })
  }

  if (existing?.format !== source.format) {
    ops.push({ kind: 'setCellFormat', sheetName, address, format: source.format ?? null })
  }
}
