import { parseCellAddress } from '@bilig/formula'
import type { LiteralInput } from '@bilig/protocol'
import type { WorkbookTableRecord } from '../../workbook-store.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'
import { rewriteFormulaSourceForRenamedStructuredReference } from './structure-structured-ref-rewrite.js'

interface TableHeaderRenameCounts {
  readonly formulaChangedCount: number
  readonly topologyChanged: boolean
}

export function applyTableHeaderRenameForSetCellValue(args: {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly value: LiteralInput
  readonly formulaChangedCount: number
  readonly topologyChanged: boolean
}): TableHeaderRenameCounts {
  const nextColumnName = tableColumnNameForLiteral(args.value)
  if (nextColumnName === undefined) {
    return { formulaChangedCount: args.formulaChangedCount, topologyChanged: args.topologyChanged }
  }

  const header = findTableHeaderCell(args.serviceArgs.state.workbook.listTables(), args.sheetName, args.row, args.col)
  if (!header) {
    return { formulaChangedCount: args.formulaChangedCount, topologyChanged: args.topologyChanged }
  }

  const previousColumnName = header.table.columnNames[header.columnIndex] ?? ''
  if (normalizeTableColumnName(previousColumnName) === normalizeTableColumnName(nextColumnName)) {
    return { formulaChangedCount: args.formulaChangedCount, topologyChanged: args.topologyChanged }
  }

  const nextColumnNames = header.table.columnNames.slice()
  nextColumnNames[header.columnIndex] = nextColumnName
  const nextColumns = header.table.columns?.map((column, index) =>
    index === header.columnIndex ? { ...column, name: nextColumnName } : column,
  )
  args.serviceArgs.state.workbook.setTable({
    ...header.table,
    columnNames: nextColumnNames,
    ...(nextColumns ? { columns: nextColumns } : {}),
  })

  let formulaChangedCount = args.formulaChangedCount
  let topologyChanged = args.topologyChanged
  for (const cellIndex of args.serviceArgs.collectFormulaCellsForTables([header.table.name])) {
    const formula = args.serviceArgs.state.formulas.get(cellIndex)
    if (!formula) {
      continue
    }
    const source = rewriteFormulaSourceForRenamedStructuredReference(formula.source, {
      tableName: header.table.name,
      oldColumnName: previousColumnName,
      newColumnName: nextColumnName,
    })
    if (!source || source === formula.source) {
      continue
    }
    const ownerSheetName = args.serviceArgs.state.workbook.getSheetNameById(args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]!)
    if (!ownerSheetName) {
      continue
    }
    const changedTopology =
      args.serviceArgs.rewriteFormulaSourcePreservingBinding?.(cellIndex, ownerSheetName, source) === true
        ? false
        : args.serviceArgs.bindFormula(cellIndex, ownerSheetName, source)
    formulaChangedCount = args.serviceArgs.markFormulaChanged(cellIndex, formulaChangedCount)
    topologyChanged = topologyChanged || changedTopology
  }

  return { formulaChangedCount, topologyChanged }
}

export function isTableHeaderCell(tables: readonly WorkbookTableRecord[], sheetName: string, row: number, col: number): boolean {
  return findTableHeaderCell(tables, sheetName, row, col) !== undefined
}

export function findTableHeaderCell(
  tables: readonly WorkbookTableRecord[],
  sheetName: string,
  row: number,
  col: number,
): { readonly table: WorkbookTableRecord; readonly columnIndex: number } | undefined {
  for (const table of tables) {
    if (table.sheetName !== sheetName || !table.headerRow) {
      continue
    }
    const start = parseCellAddress(table.startAddress, table.sheetName)
    const end = parseCellAddress(table.endAddress, table.sheetName)
    const headerRow = Math.min(start.row, end.row)
    const startCol = Math.min(start.col, end.col)
    const endCol = Math.max(start.col, end.col)
    if (row === headerRow && col >= startCol && col <= endCol) {
      return { table, columnIndex: col - startCol }
    }
  }
  return undefined
}

function normalizeTableColumnName(name: string): string {
  return name.trim().toUpperCase()
}

function tableColumnNameForLiteral(value: LiteralInput): string | undefined {
  if (value === null) {
    return undefined
  }
  const name = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : String(value)
  const trimmed = name.trim()
  return trimmed.length === 0 ? undefined : trimmed
}
