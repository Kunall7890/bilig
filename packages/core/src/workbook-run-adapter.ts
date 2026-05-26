import { ValueTag, type CellRangeRef, type CellValue, type LiteralInput, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  workbookActionCommandDigest,
  workbookPlanId,
  materializeFormulaLabels,
  normalizeWorkbookActionInput,
  type EngineOp,
  type WorkbookActionInput,
  type WorkbookActionCommand,
  type WorkbookActionPlan,
  type WorkbookRunApplyCommandReceipt,
  type WorkbookCheckResult,
  type WorkbookColumnRef,
  type WorkbookCommandResolvedRefs,
  type WorkbookNameRef,
  type WorkbookRangeRef,
  type WorkbookRef,
  type WorkbookResolvedRefData,
  type WorkbookResolvedRefValue,
  type WorkbookRowsRef,
  type WorkbookRunAdapter,
  type WorkbookFormulaLabelReplacement,
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
  readonly baseRevision?: number
  readonly revision?: number
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

function workbookRefKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
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

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function resolveApplyRevisionProof(
  options: WorkbookRunEngineAdapterOptions,
  opCount: number,
): { readonly baseRevision: number; readonly revision: number } {
  const baseRevision = options.baseRevision ?? 0
  if (!isSafeNonNegativeInteger(baseRevision)) {
    throw new Error('Workbook run adapter baseRevision must be a non-negative safe integer')
  }
  const revision = options.revision ?? (opCount > 0 ? baseRevision + 1 : baseRevision)
  if (!isSafeNonNegativeInteger(revision) || revision < baseRevision) {
    throw new Error('Workbook run adapter revision must be a non-negative safe integer at or after baseRevision')
  }
  return {
    baseRevision,
    revision,
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

function checked(check: WorkbookCheckResult, status: WorkbookCheckResult['status'], proof?: WorkbookActionInput): WorkbookCheckResult {
  return {
    ...check,
    status,
    ...(proof !== undefined ? { proof: normalizeWorkbookActionInput(proof) } : {}),
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

function formulaLabelProofForTarget(
  engine: SpreadsheetEngine,
  plan: WorkbookActionPlan,
  target: WorkbookRef,
): readonly WorkbookFormulaLabelReplacement[] | undefined {
  const targetCells = cellsFromRef(engine, target)
  if (targetCells === null || targetCells.length !== 1) {
    return undefined
  }
  const replacements = new Map<string, string>()
  plan.checks.forEach((check) => {
    if (check.target === undefined || workbookRefKey(check.target) !== workbookRefKey(target)) {
      return
    }
    if (check.expectation?.kind !== 'formulaEquals') {
      return
    }
    check.expectation.labels.forEach((label) => {
      replacements.set(label.name, inputReplacementForCell(engine, label.ref, targetCells, 0))
    })
  })
  if (replacements.size === 0) {
    return undefined
  }
  return Object.freeze([...replacements].map(([name, source]) => Object.freeze({ name, source })))
}

function readbackForRange(engine: SpreadsheetEngine, target: WorkbookRangeRef, plan: WorkbookActionPlan): WorkbookRunReadback | null {
  const range = target.range
  if (range.startAddress !== range.endAddress || !rangeExists(engine, range)) {
    return null
  }
  const cell = engine.getCell(range.sheetName, range.startAddress)
  const value = literalFromCellValue(cell.value)
  const formulaLabels = formulaLabelProofForTarget(engine, plan, target)
  return {
    target,
    ...(value !== undefined ? { value } : {}),
    formula: cell.formula ?? null,
    ...(formulaLabels !== undefined ? { formulaLabels } : {}),
  }
}

function readbackForName(engine: SpreadsheetEngine, target: WorkbookNameRef, plan: WorkbookActionPlan): WorkbookRunReadback | null {
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
  const formulaLabels = formulaLabelProofForTarget(engine, plan, target)
  return {
    target,
    ...(value !== undefined ? { value } : {}),
    formula: cell.formula ?? null,
    ...(formulaLabels !== undefined ? { formulaLabels } : {}),
  }
}

function readbackForColumn(engine: SpreadsheetEngine, target: WorkbookColumnRef, plan: WorkbookActionPlan): WorkbookRunReadback | null {
  const cells = cellsFromRef(engine, target)
  if (cells === null || cells.length !== 1) {
    return null
  }
  const targetCell = cells[0]!
  const cell = engine.getCell(targetCell.sheetName, targetCell.address)
  const value = literalFromCellValue(cell.value)
  const formulaLabels = formulaLabelProofForTarget(engine, plan, target)
  return {
    target,
    ...(value !== undefined ? { value } : {}),
    formula: cell.formula ?? null,
    ...(formulaLabels !== undefined ? { formulaLabels } : {}),
  }
}

function readbackForTarget(engine: SpreadsheetEngine, target: WorkbookRef, plan: WorkbookActionPlan): WorkbookRunReadback | null {
  switch (target.kind) {
    case 'range':
      return readbackForRange(engine, target, plan)
    case 'name':
      return readbackForName(engine, target, plan)
    case 'column':
      return readbackForColumn(engine, target, plan)
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
  const source = command.formula
  if (command.inputs.length > 0 && command.labels.length === 0) {
    throw new Error(`Formula command for ${command.target.label} has inputs but no labels`)
  }
  return materializeFormulaLabels(
    source,
    command.labels.map((label) => ({
      name: label.name,
      source: inputReplacementForCell(engine, label.ref, targetCells, cellIndex),
    })),
  )
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

function formulaLabelProofForCommandReceipt(
  engine: SpreadsheetEngine,
  command: WorkbookActionCommand,
): readonly WorkbookFormulaLabelReplacement[] | undefined {
  if (command.kind !== 'writeFormula' || command.labels.length === 0) {
    return undefined
  }
  const targetCells = commandTargetCells(engine, command)
  if (targetCells.length !== 1) {
    return undefined
  }
  return Object.freeze(
    command.labels.map((label) =>
      Object.freeze({
        name: label.name,
        source: inputReplacementForCell(engine, label.ref, targetCells, 0),
      }),
    ),
  )
}

function rangeResolvedRefData(range: CellRangeRef): WorkbookResolvedRefData {
  return {
    kind: 'range',
    id: `range_${range.sheetName}_${range.startAddress}_${range.endAddress}`.replaceAll(/[^A-Za-z0-9_]+/g, '_'),
    label: rangeSource(range),
    range,
  }
}

function resolvedRefValueForRef(engine: SpreadsheetEngine, ref: WorkbookRef): WorkbookResolvedRefValue | undefined {
  const ranges = rangesFromRef(engine, ref)
  if (ranges === null) {
    return undefined
  }
  const refs = ranges.map(rangeResolvedRefData)
  return refs.length === 1 ? refs[0] : refs
}

function resolvedRefsForCommand(engine: SpreadsheetEngine, command: WorkbookActionCommand): WorkbookCommandResolvedRefs | undefined {
  const refs: { target?: WorkbookResolvedRefValue; inputs?: WorkbookResolvedRefValue[] } = {}
  if (command.target !== undefined) {
    const target = resolvedRefValueForRef(engine, command.target)
    if (target !== undefined) {
      refs.target = target
    }
  }
  if (command.kind === 'writeFormula' && command.inputs.length > 0) {
    const inputs = command.inputs.map((input) => {
      const resolved = resolvedRefValueForRef(engine, input)
      return resolved
    })
    if (inputs.every((input): input is WorkbookResolvedRefValue => input !== undefined)) {
      refs.inputs = inputs
    }
  }
  return Object.keys(refs).length === 0 ? undefined : refs
}

function noopEffectForCommand(command: WorkbookActionCommand): WorkbookActionInput {
  switch (command.kind) {
    case 'writeValue':
      return normalizeWorkbookActionInput({ kind: command.kind, value: command.value })
    case 'writeFormula':
      return normalizeWorkbookActionInput({ kind: command.kind, formula: command.formula })
    case 'clear':
      return normalizeWorkbookActionInput({ kind: command.kind, cleared: true })
    case 'format':
      return normalizeWorkbookActionInput({
        kind: command.kind,
        ...(command.style !== undefined ? { style: command.style } : {}),
        ...(command.numberFormat !== undefined ? { numberFormat: command.numberFormat } : {}),
      })
    case 'op':
      return normalizeWorkbookActionInput({ kind: command.kind, opKind: command.op.kind, op: command.op })
  }
}

function materializePlanCommandReceipts(engine: SpreadsheetEngine, plan: WorkbookActionPlan): readonly WorkbookRunApplyCommandReceipt[] {
  if (plan.commands.length === 0) {
    return []
  }
  return plan.commands.map((command, commandIndex) => {
    const ops = materializeCommandOps(engine, command)
    const resolvedRefs = resolvedRefsForCommand(engine, command)
    const formulaLabels = formulaLabelProofForCommandReceipt(engine, command)
    const commandDigest = workbookActionCommandDigest(command)
    return {
      commandIndex,
      commandKind: command.kind,
      commandDigest,
      previewOps: ops,
      appliedOps: ops,
      ...(ops.length === 0
        ? {
            noop: {
              reason: 'already_satisfied',
              proof: {
                source: '@bilig/core',
                evidence: 'materialized_zero_ops',
                commandKind: command.kind,
                commandDigest,
                opCount: 0,
                effect: noopEffectForCommand(command),
              },
            },
          }
        : {}),
      proof: {
        source: '@bilig/core',
        opCount: ops.length,
      },
      ...(resolvedRefs !== undefined ? { resolvedRefs } : {}),
      ...(formulaLabels !== undefined ? { formulaLabels } : {}),
    }
  })
}

function materializeCommandReceiptOps(receipts: readonly WorkbookRunApplyCommandReceipt[]): readonly EngineOp[] {
  const ops: EngineOp[] = []
  for (const receipt of receipts) {
    for (const op of receipt.appliedOps) {
      ops.push(op)
    }
  }
  return ops
}

function materializePlanOps(engine: SpreadsheetEngine, plan: WorkbookActionPlan): readonly EngineOp[] {
  if (plan.commands.length === 0) {
    return plan.ops
  }
  return materializeCommandReceiptOps(materializePlanCommandReceipts(engine, plan))
}

function verifyCheck(engine: SpreadsheetEngine, check: WorkbookCheckResult): WorkbookCheckResult {
  if (check.status === 'failed' || check.target === undefined) {
    return check
  }
  if (check.kind === 'exists') {
    const exists = refExists(engine, check.target)
    return checked(check, exists ? 'passed' : 'failed', {
      source: '@bilig/core',
      check: 'exists',
      target: check.target.label,
      exists,
    })
  }
  if (check.kind === 'noFormulaErrors') {
    const status = verifyNoFormulaErrors(engine, check.target)
    return checked(check, status, {
      source: '@bilig/core',
      check: 'noFormulaErrors',
      target: check.target.label,
      passed: status === 'passed',
    })
  }
  return check
}

export function createWorkbookRunAdapter(engine: SpreadsheetEngine, options: WorkbookRunEngineAdapterOptions = {}): WorkbookRunAdapter {
  return {
    apply(plan: WorkbookActionPlan) {
      try {
        const commandReceipts = materializePlanCommandReceipts(engine, plan)
        const ops = commandReceipts.length === 0 ? materializePlanOps(engine, plan) : materializeCommandReceiptOps(commandReceipts)
        const revisionProof = resolveApplyRevisionProof(options, ops.length)
        const undoOps = engine.applyOps(ops, {
          captureUndo: options.captureUndo ?? true,
          ...(options.potentialNewCells !== undefined ? { potentialNewCells: options.potentialNewCells } : {}),
          source: 'local',
          trusted: true,
        })
        const undo = createUndoRef(plan, undoOps)
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: revisionProof.baseRevision,
          revision: revisionProof.revision,
          previewOps: ops,
          appliedOps: ops,
          ...(commandReceipts.length > 0 ? { commandReceipts } : {}),
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
    read(targets, plan) {
      return targets.flatMap((target) => {
        const readback = readbackForTarget(engine, target, plan)
        return readback === null ? [] : [readback]
      })
    },
    verifyChecks(checks) {
      return checks.map((check) => verifyCheck(engine, check))
    },
  }
}
