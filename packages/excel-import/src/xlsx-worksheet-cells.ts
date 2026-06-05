import { decodeCellAddress, encodeCellAddress } from '@bilig/xlsx'
import type { SheetJsWorkSheet } from './xlsx-sheetjs-types.js'

export interface WorksheetCellEntry {
  address: string
  cell: Record<string, unknown>
  row: number
  column: number
}

export interface WorksheetCellRecord {
  address: string
  cell: Record<string, unknown>
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function isWorksheetCellAddress(value: string): boolean {
  return /^[A-Z]{1,3}[1-9][0-9]*$/u.test(value)
}

function denseWorksheetRows(sheet: SheetJsWorkSheet): unknown[] | null {
  const denseRows = (sheet as Record<string, unknown>)['!data']
  return Array.isArray(denseRows) ? denseRows : null
}

function denseWorksheetCell(row: unknown, column: number): Record<string, unknown> | null {
  if (!Array.isArray(row)) {
    return null
  }
  const cell = row[column]
  return isRecord(cell) ? cell : null
}

function compareWorksheetCellAddresses(left: string, right: string): number {
  const leftCell = decodeCellAddress(left)
  const rightCell = decodeCellAddress(right)
  return leftCell.r - rightCell.r || leftCell.c - rightCell.c || left.localeCompare(right)
}

function sortedSparseWorksheetCellAddresses(sheet: SheetJsWorkSheet): string[] {
  return Object.keys(sheet).filter(isWorksheetCellAddress).toSorted(compareWorksheetCellAddresses)
}

export function worksheetCellAt(sheet: SheetJsWorkSheet, row: number, column: number): Record<string, unknown> | null {
  const denseRows = denseWorksheetRows(sheet)
  if (denseRows) {
    return denseWorksheetCell(denseRows[row], column)
  }
  const value = sheet[encodeCellAddress({ r: row, c: column })]
  return isRecord(value) ? value : null
}

export function* worksheetCellRecords(sheet: SheetJsWorkSheet): Generator<WorksheetCellRecord> {
  const denseRows = denseWorksheetRows(sheet)
  if (denseRows) {
    for (const [rowIndex, row] of denseRows.entries()) {
      if (!Array.isArray(row)) {
        continue
      }
      for (const [columnIndex, cell] of row.entries()) {
        if (!isRecord(cell)) {
          continue
        }
        yield {
          address: encodeCellAddress({ r: rowIndex, c: columnIndex }),
          cell,
        }
      }
    }
    return
  }

  for (const address of sortedSparseWorksheetCellAddresses(sheet)) {
    const value: unknown = sheet[address]
    if (!isRecord(value)) {
      continue
    }
    yield { address, cell: value }
  }
}

export function* worksheetCellEntries(sheet: SheetJsWorkSheet): Generator<WorksheetCellEntry> {
  const denseRows = denseWorksheetRows(sheet)
  if (denseRows) {
    for (const [rowIndex, row] of denseRows.entries()) {
      if (!Array.isArray(row)) {
        continue
      }
      for (const [columnIndex, cell] of row.entries()) {
        if (!isRecord(cell)) {
          continue
        }
        yield {
          address: encodeCellAddress({ r: rowIndex, c: columnIndex }),
          cell,
          row: rowIndex,
          column: columnIndex,
        }
      }
    }
    return
  }

  for (const address of sortedSparseWorksheetCellAddresses(sheet)) {
    const value: unknown = sheet[address]
    if (!isRecord(value)) {
      continue
    }
    const decoded = decodeCellAddress(address)
    yield {
      address,
      cell: value,
      row: decoded.r,
      column: decoded.c,
    }
  }
}

export function* worksheetCellEntriesAtAddresses(sheet: SheetJsWorkSheet, addresses: Iterable<string>): Generator<WorksheetCellEntry> {
  for (const address of addresses) {
    const decoded = decodeCellAddress(address)
    const cell = worksheetCellAt(sheet, decoded.r, decoded.c)
    if (!cell) {
      continue
    }
    yield {
      address,
      cell,
      row: decoded.r,
      column: decoded.c,
    }
  }
}
