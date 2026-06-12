import type { FormulaNode } from '@bilig/formula'

import { decodeCellAddress } from './address.js'
import { normalizeStructuredReferenceColumnName } from './streaming-native-text.js'
import type { NativeTable } from './streaming-native-recalc.js'

export function compileReferenceColumn(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): number | null {
  if (node.kind === 'CellRef') {
    if (node.sheetName && node.sheetName !== sheetName) {
      return null
    }
    const address = decodeCellAddress(node.ref.replaceAll('$', ''))
    return address.r === row ? address.c : null
  }
  if (node.kind !== 'StructuredRef' || node.endColumnName || (node.section && node.section !== 'this-row')) {
    return null
  }
  const table = findCurrentRowTable(node.tableName, sheetName, row, formulaColumn, tablesBySheet)
  if (!table) {
    return null
  }
  const columnIndex = table.columns.findIndex(
    (column) => column.toLocaleLowerCase('en-US') === normalizeStructuredReferenceColumnName(node.columnName).toLocaleLowerCase('en-US'),
  )
  return columnIndex >= 0 ? table.range.s.c + columnIndex : null
}

function findCurrentRowTable(
  tableName: string,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): NativeTable | null {
  const tables = tablesBySheet.get(sheetName) ?? []
  const matching = tables.filter((table) => {
    const nameMatches =
      tableName.length === 0 ||
      table.name.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US') ||
      table.displayName.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US')
    return nameMatches && rowIsInTableDataBody(table, row)
  })
  return matching.find((table) => formulaColumn >= table.range.s.c && formulaColumn <= table.range.e.c) ?? matching[0] ?? null
}

function rowIsInTableDataBody(table: NativeTable, row: number): boolean {
  const start = table.range.s.r + table.headerRowCount
  const end = table.range.e.r - table.totalsRowCount
  return row >= start && row <= end
}
