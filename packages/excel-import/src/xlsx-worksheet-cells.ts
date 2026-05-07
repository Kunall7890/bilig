import * as XLSX from 'xlsx'

export interface WorksheetCellEntry {
  address: string
  cell: Record<string, unknown>
  row: number
  column: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isWorksheetCellAddress(value: string): boolean {
  return /^[A-Z]{1,3}[1-9][0-9]*$/u.test(value)
}

function denseWorksheetRows(sheet: XLSX.WorkSheet): unknown[] | null {
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

export function worksheetCellAt(sheet: XLSX.WorkSheet, row: number, column: number): Record<string, unknown> | null {
  const denseRows = denseWorksheetRows(sheet)
  if (denseRows) {
    return denseWorksheetCell(denseRows[row], column)
  }
  const value = sheet[XLSX.utils.encode_cell({ r: row, c: column })]
  return isRecord(value) ? value : null
}

export function* worksheetCellEntries(sheet: XLSX.WorkSheet): Generator<WorksheetCellEntry> {
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
          address: XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex }),
          cell,
          row: rowIndex,
          column: columnIndex,
        }
      }
    }
    return
  }

  for (const address in sheet) {
    const value: unknown = sheet[address]
    if (!isWorksheetCellAddress(address) || !isRecord(value)) {
      continue
    }
    const decoded = XLSX.utils.decode_cell(address)
    yield {
      address,
      cell: value,
      row: decoded.r,
      column: decoded.c,
    }
  }
}
