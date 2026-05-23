import { ErrorCode, MAX_ROWS, type WorkbookTableSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress, type FormulaNode } from '@bilig/formula'
import type { WorkbookStore } from '../../workbook-store.js'
import type { StructuredReferenceResolutionOptions } from './formula-evaluation-service-types.js'

function normalizedTableColumnName(value: string): string {
  return value.trim().replace(/\s+/gu, ' ').toUpperCase()
}

function findStructuredReferenceTable(
  workbook: WorkbookStore,
  tableName: string,
  options: StructuredReferenceResolutionOptions | undefined,
): WorkbookTableSnapshot | undefined {
  if (tableName.length > 0) {
    return workbook.getTable(tableName)
  }
  if (!options?.ownerSheetName || !options.ownerAddress) {
    return undefined
  }
  const owner = parseCellAddress(options.ownerAddress, options.ownerSheetName)
  return workbook.listTables().find((table) => {
    if (table.sheetName !== options.ownerSheetName) {
      return false
    }
    const start = parseCellAddress(table.startAddress, table.sheetName)
    const end = parseCellAddress(table.endAddress, table.sheetName)
    return owner.row >= start.row && owner.row <= end.row && owner.col >= start.col && owner.col <= end.col
  })
}

export function resolveStructuredReferenceNow(
  workbook: WorkbookStore,
  tableName: string,
  columnName: string,
  options?: StructuredReferenceResolutionOptions,
): FormulaNode | undefined {
  const table = findStructuredReferenceTable(workbook, tableName, options)
  if (!table) {
    return undefined
  }
  const columnIndex = table.columnNames.findIndex((name) => normalizedTableColumnName(name) === normalizedTableColumnName(columnName))
  const endColumnIndex =
    options?.endColumnName === undefined
      ? columnIndex
      : table.columnNames.findIndex((name) => normalizedTableColumnName(name) === normalizedTableColumnName(options.endColumnName ?? ''))
  if (columnIndex === -1 || endColumnIndex === -1) {
    return undefined
  }
  const start = parseCellAddress(table.startAddress, table.sheetName)
  const end = parseCellAddress(table.endAddress, table.sheetName)
  const dataStartRow = start.row + (table.headerRow ? 1 : 0)
  const dataEndRow = Math.max(dataStartRow, end.row - (table.totalsRow ? 1 : 0))
  let startRow = dataStartRow
  let endRow = dataEndRow
  switch (options?.section) {
    case 'all':
      startRow = start.row
      endRow = end.row
      break
    case 'headers':
      if (!table.headerRow) {
        return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
      }
      startRow = start.row
      endRow = start.row
      break
    case 'this-row': {
      if (!options.ownerSheetName || !options.ownerAddress) {
        return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
      }
      const owner = parseCellAddress(options.ownerAddress, options.ownerSheetName)
      if (owner.row < dataStartRow || owner.row > dataEndRow) {
        return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
      }
      startRow = owner.row
      endRow = owner.row
      break
    }
    case 'totals':
      if (!table.totalsRow) {
        return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
      }
      startRow = end.row
      endRow = end.row
      break
    case 'data':
    case undefined:
      break
  }
  if (startRow >= MAX_ROWS || endRow >= MAX_ROWS) {
    return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
  }
  const startColumn = start.col + Math.min(columnIndex, endColumnIndex)
  const endColumn = start.col + Math.max(columnIndex, endColumnIndex)
  return {
    kind: 'RangeRef',
    refKind: 'cells',
    sheetName: table.sheetName,
    start: formatAddress(startRow, startColumn),
    end: formatAddress(endRow, endColumn),
  }
}
