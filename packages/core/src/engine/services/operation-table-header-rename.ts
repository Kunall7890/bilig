import { parseCellAddress } from '@bilig/formula'
import type { LiteralInput } from '@bilig/protocol'
import { normalizeDefinedName, type WorkbookTableRecord } from '../../workbook-store.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'
import { rewriteFormulaSourceForRenamedStructuredReference, type RenamedTableColumnReference } from './structure-structured-ref-rewrite.js'

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
  const renamedReference = {
    tableName: header.table.name,
    oldColumnName: previousColumnName,
    newColumnName: nextColumnName,
  }
  const changedDefinedNames = rewriteDefinedNamesForRenamedTableColumn(args.serviceArgs, renamedReference)
  if (changedDefinedNames.size > 0) {
    const reboundCount = formulaChangedCount
    formulaChangedCount = args.serviceArgs.rebindDefinedNameDependents([...changedDefinedNames], formulaChangedCount)
    topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
  }
  for (const cellIndex of args.serviceArgs.collectFormulaCellsForTables([header.table.name])) {
    const formula = args.serviceArgs.state.formulas.get(cellIndex)
    if (!formula) {
      continue
    }
    const source = rewriteFormulaSourceForRenamedStructuredReference(formula.source, renamedReference)
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

function rewriteDefinedNamesForRenamedTableColumn(
  args: CreateEngineOperationServiceArgs,
  renamedReference: RenamedTableColumnReference,
): Set<string> {
  const changedNames = new Set<string>()
  for (const record of args.state.workbook.listDefinedNames()) {
    const value = record.value
    if (typeof value === 'string' && value.startsWith('=')) {
      const source = rewriteFormulaSourceForRenamedStructuredReference(value.slice(1), renamedReference)
      if (!source || `=${source}` === value) {
        continue
      }
      args.state.workbook.setDefinedName(record.name, `=${source}`, record.scopeSheetName)
      changedNames.add(normalizeDefinedName(record.name))
      continue
    }
    if (typeof value !== 'object' || value === null) {
      continue
    }
    switch (value.kind) {
      case 'formula': {
        const formula = value.formula.startsWith('=') ? value.formula.slice(1) : value.formula
        const source = rewriteFormulaSourceForRenamedStructuredReference(formula, renamedReference)
        if (!source) {
          break
        }
        const nextFormula = value.formula.startsWith('=') ? `=${source}` : source
        if (nextFormula === value.formula) {
          break
        }
        args.state.workbook.setDefinedName(record.name, { ...value, formula: nextFormula }, record.scopeSheetName)
        changedNames.add(normalizeDefinedName(record.name))
        break
      }
      case 'structured-ref':
        if (!structuredReferenceMatches(renamedReference, value.tableName, value.columnName)) {
          break
        }
        args.state.workbook.setDefinedName(record.name, { ...value, columnName: renamedReference.newColumnName }, record.scopeSheetName)
        changedNames.add(normalizeDefinedName(record.name))
        break
      case 'cell-ref':
      case 'range-ref':
      case 'scalar':
        break
    }
  }
  return changedNames
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

function structuredReferenceMatches(reference: RenamedTableColumnReference, tableName: string, columnName: string): boolean {
  return (
    normalizeTableColumnName(reference.tableName) === normalizeTableColumnName(tableName) &&
    normalizeTableColumnName(reference.oldColumnName) === normalizeTableColumnName(columnName)
  )
}

function tableColumnNameForLiteral(value: LiteralInput): string | undefined {
  if (value === null) {
    return undefined
  }
  const name = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : String(value)
  const trimmed = name.trim()
  return trimmed.length === 0 ? undefined : trimmed
}
