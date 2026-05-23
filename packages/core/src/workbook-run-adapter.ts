import { ValueTag, type CellRangeRef, type CellValue, type LiteralInput, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  formula as workbookFormula,
  type EngineOp,
  type WorkbookActionCommand,
  type WorkbookActionPlan,
  type WorkbookCheckResult,
  type WorkbookColumnRef,
  type WorkbookNameRef,
  type WorkbookRangeRef,
  type WorkbookRef,
  type WorkbookRowsRef,
  type WorkbookRunAdapter,
  type WorkbookRunReadback,
  type WorkbookTableRef,
  type WorkbookUndoRef,
} from '@bilig/workbook'
import type { SpreadsheetEngine } from './engine.js'
import { buildFormatClearOps, buildFormatPatchOps, buildStylePatchOps } from './engine-range-format-ops.js'
import type { WorkbookTableRecord } from './workbook-store.js'

export interface WorkbookRunEngineAdapterOptions {
  readonly captureUndo?: boolean
  readonly potentialNewCells?: number
}

let undoRefCounter = 0

interface RangeBounds {
  readonly startRow: number
  readonly startCol: number
  readonly endRow: number
  readonly endCol: number
  readonly height: number
  readonly width: number
}

interface CellTarget {
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
}

interface ResolvedRows {
  readonly table: WorkbookTableRecord
  readonly rows: readonly number[]
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function createUndoRef(plan: WorkbookActionPlan, ops: readonly EngineOp[] | null): WorkbookUndoRef | undefined {
  if (ops === null || ops.length === 0) {
    return undefined
  }
  undoRefCounter += 1
  return {
    id: `${plan.modelName}.${plan.actionName}.undo.${String(undoRefCounter)}`,
    ops,
  }
}

function literalFromCellValue(value: CellValue): LiteralInput | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return undefined
  }
}

function checked(check: WorkbookCheckResult, status: WorkbookCheckResult['status']): WorkbookCheckResult {
  return {
    ...check,
    status,
  }
}

function sheetExists(engine: SpreadsheetEngine, sheetName: string): boolean {
  return engine.exportSnapshot().sheets.some((sheet) => sheet.name === sheetName)
}

function rangeExists(engine: SpreadsheetEngine, range: CellRangeRef): boolean {
  if (!sheetExists(engine, range.sheetName)) {
    return false
  }
  try {
    rangeBounds(range)
    return true
  } catch {
    return false
  }
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function rangeBounds(range: CellRangeRef): RangeBounds {
  const start = parseCellAddress(range.startAddress)
  const end = parseCellAddress(range.endAddress)
  if (end.row < start.row || end.col < start.col) {
    throw new Error(`Invalid range ${range.sheetName}!${range.startAddress}:${range.endAddress}`)
  }
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
    height: end.row - start.row + 1,
    width: end.col - start.col + 1,
  }
}

function rangeSource(range: CellRangeRef): string {
  const sheet = quoteSheetName(range.sheetName)
  return range.startAddress === range.endAddress ? `${sheet}!${range.startAddress}` : `${sheet}!${range.startAddress}:${range.endAddress}`
}

function cellSource(sheetName: string, row: number, col: number): string {
  return `${quoteSheetName(sheetName)}!${formatAddress(row, col)}`
}

function cellsFromRange(range: CellRangeRef): readonly CellTarget[] {
  const bounds = rangeBounds(range)
  const cells: CellTarget[] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
      cells.push({
        sheetName: range.sheetName,
        address: formatAddress(row, col),
        row,
        col,
      })
    }
  }
  return cells
}

function tableMatches(ref: WorkbookTableRef, table: WorkbookTableRecord): boolean {
  if (ref.name !== undefined && table.name !== ref.name) {
    return false
  }
  if (ref.sheetName !== undefined && table.sheetName !== ref.sheetName) {
    return false
  }
  if (ref.headers !== undefined && !ref.headers.every((header) => table.columnNames.includes(header))) {
    return false
  }
  return true
}

function findTable(engine: SpreadsheetEngine, ref: WorkbookTableRef): WorkbookTableRecord | undefined {
  if (ref.name !== undefined) {
    const table = engine.getTable(ref.name)
    return table !== undefined && tableMatches(ref, table) ? table : undefined
  }
  const matches = engine.getTables().filter((table) => tableMatches(ref, table))
  if (matches.length > 1) {
    throw new Error(
      `Ambiguous table selector ${ref.label} matched ${matches.map((table) => table.name).join(', ')}; add a table name or sheet constraint`,
    )
  }
  return matches[0]
}

function tableRange(table: WorkbookTableRecord): CellRangeRef {
  return {
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
  }
}

function tableDataBounds(table: WorkbookTableRecord): RangeBounds | null {
  const bounds = rangeBounds(tableRange(table))
  const startRow = table.headerRow ? bounds.startRow + 1 : bounds.startRow
  const endRow = table.totalsRow ? bounds.endRow - 1 : bounds.endRow
  if (endRow < startRow) {
    return null
  }
  return {
    startRow,
    startCol: bounds.startCol,
    endRow,
    endCol: bounds.endCol,
    height: endRow - startRow + 1,
    width: bounds.width,
  }
}

function columnRange(table: WorkbookTableRecord, columnName: string): CellRangeRef | null {
  const columnIndex = table.columnNames.indexOf(columnName)
  if (columnIndex < 0) {
    return null
  }
  const start = parseCellAddress(table.startAddress)
  const end = parseCellAddress(table.endAddress)
  const column = start.col + columnIndex
  if (column > end.col) {
    return null
  }
  const startRow = table.headerRow ? start.row + 1 : start.row
  const endRow = table.totalsRow ? end.row - 1 : end.row
  if (endRow < startRow) {
    return null
  }
  return {
    sheetName: table.sheetName,
    startAddress: formatAddress(startRow, column),
    endAddress: formatAddress(endRow, column),
  }
}

function orderedCompare(actual: LiteralInput, expected: LiteralInput): number | null {
  if (typeof actual === 'number' && typeof expected === 'number') {
    return actual - expected
  }
  if (typeof actual === 'string' && typeof expected === 'string') {
    return actual.localeCompare(expected)
  }
  return null
}

function rowValueMatches(actual: LiteralInput | undefined, op: WorkbookRowsRef['where']['op'], expected: LiteralInput): boolean {
  if (actual === undefined) {
    return false
  }
  switch (op) {
    case 'eq':
      return Object.is(actual, expected)
    case 'neq':
      return !Object.is(actual, expected)
    case 'contains':
      return typeof actual === 'string' && typeof expected === 'string' && actual.includes(expected)
    case 'startsWith':
      return typeof actual === 'string' && typeof expected === 'string' && actual.startsWith(expected)
    case 'gt': {
      const compared = orderedCompare(actual, expected)
      return compared !== null && compared > 0
    }
    case 'gte': {
      const compared = orderedCompare(actual, expected)
      return compared !== null && compared >= 0
    }
    case 'lt': {
      const compared = orderedCompare(actual, expected)
      return compared !== null && compared < 0
    }
    case 'lte': {
      const compared = orderedCompare(actual, expected)
      return compared !== null && compared <= 0
    }
  }
}

function resolveRows(engine: SpreadsheetEngine, ref: WorkbookRowsRef): ResolvedRows | null {
  if (ref.table === undefined) {
    return null
  }
  const table = findTable(engine, ref.table)
  if (table === undefined) {
    return null
  }
  const dataBounds = tableDataBounds(table)
  if (dataBounds === null) {
    return {
      table,
      rows: [],
    }
  }
  const predicateIndex = table.columnNames.indexOf(ref.where.column)
  if (predicateIndex < 0) {
    return null
  }
  const predicateCol = dataBounds.startCol + predicateIndex
  const rows: number[] = []
  for (let row = dataBounds.startRow; row <= dataBounds.endRow; row += 1) {
    const value = literalFromCellValue(engine.getCellValue(table.sheetName, formatAddress(row, predicateCol)))
    if (rowValueMatches(value, ref.where.op, ref.where.value)) {
      rows.push(row)
    }
  }
  return {
    table,
    rows,
  }
}

function rowRangesFromRowsRef(engine: SpreadsheetEngine, ref: WorkbookRowsRef): readonly CellRangeRef[] | null {
  const resolved = resolveRows(engine, ref)
  if (resolved === null) {
    return null
  }
  const dataBounds = tableDataBounds(resolved.table)
  if (dataBounds === null) {
    return []
  }
  return resolved.rows.map((row) => ({
    sheetName: resolved.table.sheetName,
    startAddress: formatAddress(row, dataBounds.startCol),
    endAddress: formatAddress(row, dataBounds.endCol),
  }))
}

function rangesFromColumnRef(engine: SpreadsheetEngine, ref: WorkbookColumnRef): readonly CellRangeRef[] | null {
  if (ref.rows === undefined) {
    const table = findTable(engine, ref.table)
    const range = table === undefined ? null : columnRange(table, ref.name)
    return range === null ? null : [range]
  }
  const resolved = resolveRows(engine, ref.rows)
  if (resolved === null) {
    return null
  }
  const dataBounds = tableDataBounds(resolved.table)
  if (dataBounds === null) {
    return []
  }
  const columnIndex = resolved.table.columnNames.indexOf(ref.name)
  if (columnIndex < 0) {
    return null
  }
  const col = dataBounds.startCol + columnIndex
  return resolved.rows.map((row) => ({
    sheetName: resolved.table.sheetName,
    startAddress: formatAddress(row, col),
    endAddress: formatAddress(row, col),
  }))
}

function rangeFromDefinedNameValue(value: WorkbookDefinedNameValueSnapshot): CellRangeRef | null {
  if (typeof value !== 'object' || value === null) {
    return null
  }
  if (value.kind === 'cell-ref') {
    return {
      sheetName: value.sheetName,
      startAddress: value.address,
      endAddress: value.address,
    }
  }
  if (value.kind === 'range-ref') {
    return {
      sheetName: value.sheetName,
      startAddress: value.startAddress,
      endAddress: value.endAddress,
    }
  }
  return null
}

function rangesFromRef(engine: SpreadsheetEngine, ref: WorkbookRef): readonly CellRangeRef[] | null {
  switch (ref.kind) {
    case 'range':
      return [ref.range]
    case 'name': {
      const definedName = engine.getDefinedName(ref.name)
      const range = definedName === undefined ? null : rangeFromDefinedNameValue(definedName.value)
      return range === null ? null : [range]
    }
    case 'table': {
      const table = findTable(engine, ref)
      return table === undefined ? null : [tableRange(table)]
    }
    case 'column':
      return rangesFromColumnRef(engine, ref)
    case 'rows':
      return rowRangesFromRowsRef(engine, ref)
  }
}

function cellsFromRef(engine: SpreadsheetEngine, ref: WorkbookRef): readonly CellTarget[] | null {
  const ranges = rangesFromRef(engine, ref)
  return ranges === null ? null : ranges.flatMap((range) => cellsFromRange(range))
}

function refExists(engine: SpreadsheetEngine, ref: WorkbookRef): boolean {
  switch (ref.kind) {
    case 'range':
      return rangeExists(engine, ref.range)
    case 'name':
      return engine.getDefinedName(ref.name) !== undefined
    case 'table':
      return findTable(engine, ref) !== undefined
    case 'column': {
      if (ref.rows === undefined) {
        const table = findTable(engine, ref.table)
        return table !== undefined && table.columnNames.includes(ref.name)
      }
      const cells = cellsFromRef(engine, ref)
      return cells !== null && cells.length > 0
    }
    case 'rows': {
      const cells = cellsFromRef(engine, ref)
      return cells !== null && cells.length > 0
    }
  }
}

function rangeHasFormulaErrors(engine: SpreadsheetEngine, range: CellRangeRef): boolean {
  if (!rangeExists(engine, range)) {
    return true
  }
  const values = engine.getRangeValues(range)
  return values.some((row) => row.some((value) => value.tag === ValueTag.Error))
}

function verifyNoFormulaErrors(engine: SpreadsheetEngine, ref: WorkbookRef): 'passed' | 'failed' | 'planned' {
  const ranges = rangesFromRef(engine, ref)
  if (ranges === null) {
    return 'failed'
  }
  return ranges.some((range) => rangeHasFormulaErrors(engine, range)) ? 'failed' : 'passed'
}

function readbackForRange(engine: SpreadsheetEngine, target: WorkbookRangeRef): WorkbookRunReadback | null {
  const range = target.range
  if (range.startAddress !== range.endAddress || !rangeExists(engine, range)) {
    return null
  }
  const cell = engine.getCell(range.sheetName, range.startAddress)
  const value = literalFromCellValue(cell.value)
  return {
    target,
    ...(value !== undefined ? { value } : {}),
    formula: cell.formula ?? null,
  }
}

function readbackForName(engine: SpreadsheetEngine, target: WorkbookNameRef): WorkbookRunReadback | null {
  const definedName = engine.getDefinedName(target.name)
  if (definedName === undefined) {
    return null
  }
  const range = rangeFromDefinedNameValue(definedName.value)
  if (range === null || range.startAddress !== range.endAddress || !rangeExists(engine, range)) {
    return null
  }
  const cell = engine.getCell(range.sheetName, range.startAddress)
  const value = literalFromCellValue(cell.value)
  return {
    target,
    ...(value !== undefined ? { value } : {}),
    formula: cell.formula ?? null,
  }
}

function readbackForColumn(engine: SpreadsheetEngine, target: WorkbookColumnRef): WorkbookRunReadback | null {
  const cells = cellsFromRef(engine, target)
  if (cells === null || cells.length !== 1) {
    return null
  }
  const targetCell = cells[0]!
  const cell = engine.getCell(targetCell.sheetName, targetCell.address)
  const value = literalFromCellValue(cell.value)
  return {
    target,
    ...(value !== undefined ? { value } : {}),
    formula: cell.formula ?? null,
  }
}

function readbackForTarget(engine: SpreadsheetEngine, target: WorkbookRef): WorkbookRunReadback | null {
  switch (target.kind) {
    case 'range':
      return readbackForRange(engine, target)
    case 'name':
      return readbackForName(engine, target)
    case 'column':
      return readbackForColumn(engine, target)
    case 'table':
    case 'rows':
      return null
  }
}

function inputReplacementForCell(
  engine: SpreadsheetEngine,
  input: WorkbookRef,
  targetCells: readonly CellTarget[],
  cellIndex: number,
): string {
  const inputCells = cellsFromRef(engine, input)
  if (inputCells === null) {
    throw new Error(`Cannot resolve formula input ${input.label}`)
  }
  if (inputCells.length === targetCells.length) {
    const inputCell = inputCells[cellIndex]
    if (inputCell === undefined) {
      throw new Error(`Cannot align formula input ${input.label}`)
    }
    return cellSource(inputCell.sheetName, inputCell.row, inputCell.col)
  }
  if (inputCells.length === 1) {
    const inputCell = inputCells[0]!
    return cellSource(inputCell.sheetName, inputCell.row, inputCell.col)
  }
  const inputRanges = rangesFromRef(engine, input)
  if (inputRanges !== null && inputRanges.length === 1) {
    return rangeSource(inputRanges[0]!)
  }
  throw new Error(`Cannot align formula input ${input.label} with ${targetCells.length} target cell(s)`)
}

function formulaForCell(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  targetCells: readonly CellTarget[],
  cellIndex: number,
): string {
  let source = command.formula
  command.inputs.forEach((input) => {
    const token = workbookFormula.source(workbookFormula.ref(input))
    const replacement = inputReplacementForCell(engine, input, targetCells, cellIndex)
    source = source.replaceAll(token, replacement)
  })
  return source
}

function commandTargetRanges(
  engine: SpreadsheetEngine,
  command: Exclude<WorkbookActionCommand, { readonly kind: 'op' }>,
): readonly CellRangeRef[] {
  const ranges = rangesFromRef(engine, command.target)
  if (ranges === null) {
    throw new Error(`Cannot resolve ${command.target.label} for ${command.kind}`)
  }
  return ranges
}

function commandTargetCells(
  engine: SpreadsheetEngine,
  command: Exclude<WorkbookActionCommand, { readonly kind: 'op' }>,
): readonly CellTarget[] {
  return commandTargetRanges(engine, command).flatMap((range) => cellsFromRange(range))
}

function materializeCommandOps(engine: SpreadsheetEngine, command: WorkbookActionCommand): readonly EngineOp[] {
  if (command.kind === 'op') {
    return [structuredClone(command.op)]
  }

  switch (command.kind) {
    case 'writeFormula': {
      const cells = commandTargetCells(engine, command)
      return cells.map((cell, index) => ({
        kind: 'setCellFormula',
        sheetName: cell.sheetName,
        address: cell.address,
        formula: formulaForCell(engine, command, cells, index),
      }))
    }
    case 'writeValue':
      return commandTargetCells(engine, command).map((cell) => ({
        kind: 'setCellValue',
        sheetName: cell.sheetName,
        address: cell.address,
        value: command.value,
      }))
    case 'clear':
      return commandTargetCells(engine, command).map((cell) => ({
        kind: 'clearCell',
        sheetName: cell.sheetName,
        address: cell.address,
      }))
    case 'format':
      if (command.style === undefined && command.numberFormat === undefined) {
        return []
      }
      {
        const ranges = commandTargetRanges(engine, command)
        const ops: EngineOp[] = []
        for (const range of ranges) {
          if (command.style !== undefined) {
            ops.push(...buildStylePatchOps(engine.workbook, range, command.style))
          }
          if (command.numberFormat !== undefined) {
            ops.push(
              ...(command.numberFormat === null
                ? buildFormatClearOps(engine.workbook, range)
                : buildFormatPatchOps(engine.workbook, range, command.numberFormat)),
            )
          }
        }
        return ops
      }
  }
}

function materializePlanOps(engine: SpreadsheetEngine, plan: WorkbookActionPlan): readonly EngineOp[] {
  if (plan.commands.length === 0) {
    return plan.ops
  }
  return plan.commands.flatMap((command) => materializeCommandOps(engine, command))
}

function verifyCheck(engine: SpreadsheetEngine, check: WorkbookCheckResult): WorkbookCheckResult {
  if (check.status === 'failed' || check.target === undefined) {
    return check
  }
  if (check.kind === 'exists') {
    return checked(check, refExists(engine, check.target) ? 'passed' : 'failed')
  }
  if (check.kind === 'noFormulaErrors') {
    return checked(check, verifyNoFormulaErrors(engine, check.target))
  }
  return check
}

export function createWorkbookRunAdapter(engine: SpreadsheetEngine, options: WorkbookRunEngineAdapterOptions = {}): WorkbookRunAdapter {
  return {
    apply(plan: WorkbookActionPlan) {
      try {
        const ops = materializePlanOps(engine, plan)
        const undoOps = engine.applyOps(ops, {
          captureUndo: options.captureUndo ?? true,
          ...(options.potentialNewCells !== undefined ? { potentialNewCells: options.potentialNewCells } : {}),
          source: 'local',
          trusted: true,
        })
        const undo = createUndoRef(plan, undoOps)
        return {
          status: 'applied',
          previewOps: ops,
          appliedOps: ops,
          proof: {
            source: '@bilig/core',
            opCount: ops.length,
          },
          ...(undo !== undefined ? { undo } : {}),
        }
      } catch (error) {
        return {
          status: 'failed',
          errors: [
            {
              code: 'apply_failed',
              message: errorMessage(error),
            },
          ],
        }
      }
    },
    read(targets) {
      return targets.flatMap((target) => {
        const readback = readbackForTarget(engine, target)
        return readback === null ? [] : [readback]
      })
    },
    verifyChecks(checks) {
      return checks.map((check) => verifyCheck(engine, check))
    },
  }
}
