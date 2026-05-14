import { formatRangeAddress, parseRangeAddress, type FormulaNode, type RangeRefNode } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { WorkbookStore } from '../../workbook-store.js'

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
    let maxResidentRow = -1
    sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
      if (col >= bounds.colStart && col <= bounds.colEnd && row > maxResidentRow) {
        maxResidentRow = row
      }
    })
    return {
      rows: maxResidentRow + 1,
      cols: bounds.colEnd - bounds.colStart + 1,
    }
  }

  let maxResidentCol = -1
  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (row >= bounds.rowStart && row <= bounds.rowEnd && col > maxResidentCol) {
      maxResidentCol = col
    }
  })
  return {
    rows: bounds.rowEnd - bounds.rowStart + 1,
    cols: maxResidentCol + 1,
  }
}
