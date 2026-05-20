import type { CellStyleRecord, SheetStyleRangeSnapshot } from '@bilig/protocol'
import { internImportedStyle } from './xlsx-import-cell-styles.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'

export function buildLargeSimpleStyleRanges(
  sheetName: string,
  cellScan: ImportedWorksheetCellScan,
  stylesByIndex: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>,
  styleCatalog: Map<string, CellStyleRecord>,
): SheetStyleRangeSnapshot[] {
  const ranges: SheetStyleRangeSnapshot[] = []
  let active: { row: number; startColumn: number; endColumn: number; styleId: string } | undefined
  const styleIdsByIndex = new Map<number, string>()
  cellScan.styleIndexes.forEach((row, column, styleIndex) => {
    let styleId = styleIdsByIndex.get(styleIndex)
    if (!styleId) {
      const style = stylesByIndex.get(styleIndex)
      if (!style) {
        return
      }
      styleId = internImportedStyle(style, styleCatalog)
      styleIdsByIndex.set(styleIndex, styleId)
    }
    if (active && active.row === row && active.endColumn + 1 === column && active.styleId === styleId) {
      active.endColumn = column
      return
    }
    if (active) {
      ranges.push(styleRunToRange(sheetName, active))
    }
    active = { row, startColumn: column, endColumn: column, styleId }
  })
  if (active) {
    ranges.push(styleRunToRange(sheetName, active))
  }
  return ranges
}

export function buildLargeSimpleStyleRangesForCells(
  sheetName: string,
  cells: readonly { readonly row: number; readonly column: number; readonly styleId: string }[],
): SheetStyleRangeSnapshot[] {
  const ranges: SheetStyleRangeSnapshot[] = []
  let active: { row: number; startColumn: number; endColumn: number; styleId: string } | undefined
  for (const cell of cells) {
    if (active && active.row === cell.row && active.endColumn + 1 === cell.column && active.styleId === cell.styleId) {
      active.endColumn = cell.column
      continue
    }
    if (active) {
      ranges.push(styleRunToRange(sheetName, active))
    }
    active = { row: cell.row, startColumn: cell.column, endColumn: cell.column, styleId: cell.styleId }
  }
  if (active) {
    ranges.push(styleRunToRange(sheetName, active))
  }
  return ranges
}

function styleRunToRange(
  sheetName: string,
  run: { readonly row: number; readonly startColumn: number; readonly endColumn: number; readonly styleId: string },
): SheetStyleRangeSnapshot {
  return {
    range: {
      sheetName,
      startAddress: encodeCellAddress(run.row, run.startColumn),
      endAddress: encodeCellAddress(run.row, run.endColumn),
    },
    styleId: run.styleId,
  }
}

function encodeCellAddress(row: number, column: number): string {
  let value = column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(row + 1)}`
}
