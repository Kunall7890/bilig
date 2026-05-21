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

function rangeOps(range: CellRangeRef, create: (address: string, rowOffset: number, colOffset: number) => EngineOp): readonly EngineOp[] {
  const bounds = rangeBounds(range)
  const ops: EngineOp[] = []
  for (let rowOffset = 0; rowOffset < bounds.height; rowOffset += 1) {
    for (let colOffset = 0; colOffset < bounds.width; colOffset += 1) {
      ops.push(create(formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset), rowOffset, colOffset))
    }
  }
  return ops
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

function rangeFromRef(engine: SpreadsheetEngine, ref: WorkbookRef): CellRangeRef | null {
  switch (ref.kind) {
    case 'range':
      return ref.range
    case 'name': {
      const definedName = engine.getDefinedName(ref.name)
      return definedName === undefined ? null : rangeFromDefinedNameValue(definedName.value)
    }
    case 'table': {
      const table = findTable(engine, ref)
      return table === undefined ? null : tableRange(table)
    }
    case 'column': {
      const table = findTable(engine, ref.table)
      return table === undefined ? null : columnRange(table, ref.name)
    }
    case 'rows':
      return null
  }
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
      const table = findTable(engine, ref.table)
      return table !== undefined && table.columnNames.includes(ref.name)
    }
    case 'rows':
      return rowsOwnerExists(engine, ref)
  }
}

function rowsOwnerExists(engine: SpreadsheetEngine, ref: WorkbookRowsRef): boolean {
  if (ref.table !== undefined) {
    return findTable(engine, ref.table) !== undefined
  }
  return ref.sheetName !== undefined && sheetExists(engine, ref.sheetName)
}

function rangeHasFormulaErrors(engine: SpreadsheetEngine, range: CellRangeRef): boolean {
  if (!rangeExists(engine, range)) {
    return true
  }
  const values = engine.getRangeValues(range)
  return values.some((row) => row.some((value) => value.tag === ValueTag.Error))
}

function verifyNoFormulaErrors(engine: SpreadsheetEngine, ref: WorkbookRef): 'passed' | 'failed' | 'planned' {
  const range = rangeFromRef(engine, ref)
  if (range === null) {
    return 'failed'
  }
  return rangeHasFormulaErrors(engine, range) ? 'failed' : 'passed'
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
  const table = findTable(engine, target.table)
  const range = table === undefined ? null : columnRange(table, target.name)
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
  targetBounds: RangeBounds,
  rowOffset: number,
  colOffset: number,
): string {
  const inputRange = rangeFromRef(engine, input)
  if (inputRange === null) {
    throw new Error(`Cannot resolve formula input ${input.label}`)
  }
  const inputBounds = rangeBounds(inputRange)
  if (inputBounds.height === targetBounds.height && inputBounds.width === targetBounds.width) {
    return cellSource(inputRange.sheetName, inputBounds.startRow + rowOffset, inputBounds.startCol + colOffset)
  }
  return rangeSource(inputRange)
}

function formulaForCell(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  targetRange: CellRangeRef,
  rowOffset: number,
  colOffset: number,
): string {
  const targetBounds = rangeBounds(targetRange)
  let source = command.formula
  command.inputs.forEach((input) => {
    const token = workbookFormula.source(workbookFormula.ref(input))
    const replacement = inputReplacementForCell(engine, input, targetBounds, rowOffset, colOffset)
    source = source.replaceAll(token, replacement)
  })
  return source
}

function commandTargetRange(engine: SpreadsheetEngine, command: Exclude<WorkbookActionCommand, { readonly kind: 'op' }>): CellRangeRef {
  const range = rangeFromRef(engine, command.target)
  if (range === null) {
    throw new Error(`Cannot resolve ${command.target.label} for ${command.kind}`)
  }
  return range
}

function materializeCommandOps(engine: SpreadsheetEngine, command: WorkbookActionCommand): readonly EngineOp[] {
  if (command.kind === 'op') {
    return [structuredClone(command.op)]
  }

  const range = commandTargetRange(engine, command)
  switch (command.kind) {
    case 'writeFormula':
      return rangeOps(range, (address, rowOffset, colOffset) => ({
        kind: 'setCellFormula',
        sheetName: range.sheetName,
        address,
        formula: formulaForCell(engine, command, range, rowOffset, colOffset),
      }))
    case 'writeValue':
      return rangeOps(range, (address) => ({
        kind: 'setCellValue',
        sheetName: range.sheetName,
        address,
        value: command.value,
      }))
    case 'clear':
      return rangeOps(range, (address) => ({
        kind: 'clearCell',
        sheetName: range.sheetName,
        address,
      }))
    case 'format':
      if (command.style === undefined && command.numberFormat === undefined) {
        return []
      }
      {
        const ops: EngineOp[] = []
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
          ...(undo !== undefined ? { undo } : {}),
        }
      } catch (error) {
        return {
          status: 'failed',
          errors: [
            {
              code: 'engine_apply_failed',
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
