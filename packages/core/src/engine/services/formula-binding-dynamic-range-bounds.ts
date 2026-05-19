import { formatRangeAddress, parseRangeAddress, type FormulaNode, type RangeRefNode } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { SheetRecord } from '../../workbook-sheet-record.js'
import type { WorkbookStore } from '../../workbook-store.js'

const INDEXED_RESIDENT_AXIS_BOUND_LIMIT = 4096

export interface DynamicRangeBounds {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export function normalizeRangeDependency(node: RangeRefNode): string | undefined {
  if (node.sheetEndName !== undefined) {
    return undefined
  }
  try {
    return formatRangeAddress(
      parseRangeAddress(node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`),
    )
  } catch {
    return undefined
  }
}

export function rangeBounds(node: FormulaNode | undefined, ownerSheetName: string): DynamicRangeBounds | undefined {
  if (!node || node.kind !== 'RangeRef' || node.sheetEndName !== undefined) {
    return undefined
  }
  try {
    const parsed = parseRangeAddress(`${node.start}:${node.end}`, node.sheetName ?? ownerSheetName)
    if (parsed.kind === 'cells') {
      return {
        sheetName: parsed.sheetName ?? ownerSheetName,
        rowStart: Math.min(parsed.start.row, parsed.end.row),
        rowEnd: Math.max(parsed.start.row, parsed.end.row),
        colStart: Math.min(parsed.start.col, parsed.end.col),
        colEnd: Math.max(parsed.start.col, parsed.end.col),
      }
    }
    if (parsed.kind === 'rows') {
      return {
        sheetName: parsed.sheetName ?? ownerSheetName,
        rowStart: Math.min(parsed.start.row, parsed.end.row),
        rowEnd: Math.max(parsed.start.row, parsed.end.row),
        colStart: 0,
        colEnd: MAX_COLS - 1,
      }
    }
    return {
      sheetName: parsed.sheetName ?? ownerSheetName,
      rowStart: 0,
      rowEnd: MAX_ROWS - 1,
      colStart: Math.min(parsed.start.col, parsed.end.col),
      colEnd: Math.max(parsed.start.col, parsed.end.col),
    }
  } catch {
    return undefined
  }
}

function findMaxResidentRowInColumns(sheet: SheetRecord, bounds: DynamicRangeBounds): number {
  let maxResidentRow = -1
  if (bounds.colEnd - bounds.colStart + 1 <= INDEXED_RESIDENT_AXIS_BOUND_LIMIT) {
    for (let col = bounds.colStart; col <= bounds.colEnd; col += 1) {
      sheet.logical.forEachVisibleColumnCellEntry(col, (_cellIndex, row) => {
        if (row >= bounds.rowStart && row <= bounds.rowEnd && row > maxResidentRow) {
          maxResidentRow = row
        }
      })
    }
    return maxResidentRow
  }
  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (col >= bounds.colStart && col <= bounds.colEnd && row >= bounds.rowStart && row <= bounds.rowEnd && row > maxResidentRow) {
      maxResidentRow = row
    }
  })
  return maxResidentRow
}

function findMaxResidentColInRows(sheet: SheetRecord, bounds: DynamicRangeBounds): number {
  let maxResidentCol = -1
  if (bounds.rowEnd - bounds.rowStart + 1 <= INDEXED_RESIDENT_AXIS_BOUND_LIMIT) {
    for (let row = bounds.rowStart; row <= bounds.rowEnd; row += 1) {
      sheet.logical.forEachVisibleRowCellEntry(row, (_cellIndex, col) => {
        if (col >= bounds.colStart && col <= bounds.colEnd && col > maxResidentCol) {
          maxResidentCol = col
        }
      })
    }
    return maxResidentCol
  }
  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (row >= bounds.rowStart && row <= bounds.rowEnd && col >= bounds.colStart && col <= bounds.colEnd && col > maxResidentCol) {
      maxResidentCol = col
    }
  })
  return maxResidentCol
}

export function residentRangeShape(args: {
  readonly workbook: WorkbookStore
  readonly ownerSheetName: string
  readonly range: RangeRefNode
}): { readonly rows: number; readonly cols: number } | undefined {
  if (args.range.sheetEndName !== undefined) {
    return undefined
  }
  const bounds = rangeBounds(args.range, args.ownerSheetName)
  if (!bounds) {
    return undefined
  }
  if (args.range.refKind === 'cells') {
    return {
      rows: bounds.rowEnd - bounds.rowStart + 1,
      cols: bounds.colEnd - bounds.colStart + 1,
    }
  }
  const sheet = args.workbook.getSheet(bounds.sheetName)
  if (!sheet) {
    return undefined
  }
  if (args.range.refKind === 'cols') {
    const maxResidentRow = findMaxResidentRowInColumns(sheet, bounds)
    return {
      rows: maxResidentRow + 1,
      cols: bounds.colEnd - bounds.colStart + 1,
    }
  }

  const maxResidentCol = findMaxResidentColInRows(sheet, bounds)
  return {
    rows: bounds.rowEnd - bounds.rowStart + 1,
    cols: maxResidentCol + 1,
  }
}
