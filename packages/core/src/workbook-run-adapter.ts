import { ValueTag, type CellRangeRef, type CellValue, type LiteralInput, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  formula as workbookFormula,
  type EngineOp,
  type WorkbookActionCommand,
  type WorkbookActionPlan,
  type WorkbookCheckProof,
  type WorkbookCheckResult,
  type WorkbookColumnRef,
  type WorkbookRef,
  type WorkbookRowsRef,
  type WorkbookCellReadback,
  type WorkbookRunAdapter,
  type WorkbookRunReadback,
  type WorkbookTableRef,
  type WorkbookUndoRef,
  describeRuntimeRequirements,
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

function runtimeProof(message: string, data: Record<string, LiteralInput>): WorkbookCheckProof {
  return {
    kind: 'runtime',
    message,
    data,
  }
}

function checked(check: WorkbookCheckResult, status: WorkbookCheckResult['status'], proof?: WorkbookCheckProof): WorkbookCheckResult {
  return {
    ...check,
    status,
    ...(proof !== undefined ? { proof } : {}),
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

function readbackForRanges(engine: SpreadsheetEngine, target: WorkbookRef, ranges: readonly CellRangeRef[]): WorkbookRunReadback {
  const values: LiteralInput[][] = []
  const formulas: (string | null)[][] = []
  const cells: WorkbookCellReadback[] = []
  let hasValueError = false

  for (const range of ranges) {
    const bounds = rangeBounds(range)
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const valueRow: LiteralInput[] = []
      const formulaRow: (string | null)[] = []
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        const address = formatAddress(row, col)
        const cell = engine.getCell(range.sheetName, address)
        const value = literalFromCellValue(cell.value)
        if (cell.value.tag === ValueTag.Error) {
          hasValueError = true
        }
        valueRow.push(value ?? null)
        formulaRow.push(cell.formula ?? null)
        cells.push({
          sheetName: range.sheetName,
          address,
          ...(value !== undefined ? { value } : {}),
          formula: cell.formula ?? null,
        })
      }
      values.push(valueRow)
      formulas.push(formulaRow)
    }
  }

  const single = cells.length === 1 ? cells[0] : undefined
  return {
    target,
    ...(!hasValueError && single?.value !== undefined ? { value: single.value } : {}),
    ...(single !== undefined ? { formula: single.formula ?? null } : {}),
    ...(!hasValueError ? { values } : {}),
    formulas,
    cells,
  }
}

function readbackForTarget(engine: SpreadsheetEngine, target: WorkbookRef): WorkbookRunReadback | null {
  const ranges = rangesFromRef(engine, target)
  if (ranges === null || !ranges.every((range) => rangeExists(engine, range))) {
    return null
  }
  return readbackForRanges(engine, target, ranges)
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

interface FormulaReplacement {
  readonly token: string
  readonly replacement: string
}

function formulaQuoteEnd(source: string, startIndex: number, quote: '"' | "'"): number {
  let index = startIndex + 1
  while (index < source.length) {
    const current = source[index]
    if (current === quote) {
      if (source[index + 1] === quote) {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return source.length
}

function isFormulaIdentifierChar(value: string | undefined): boolean {
  return value !== undefined && /[A-Za-z0-9_.$]/.test(value)
}

function isFormulaTokenBoundary(source: string, index: number, token: string): boolean {
  const previous = source[index - 1]
  const next = source[index + token.length]
  return !isFormulaIdentifierChar(previous) && previous !== ']' && !isFormulaIdentifierChar(next) && next !== '['
}

function materializeFormulaSource(source: string, replacements: readonly FormulaReplacement[]): string {
  let output = ''
  let index = 0
  const orderedReplacements = replacements.toSorted((left, right) => right.token.length - left.token.length)

  while (index < source.length) {
    const replacement = orderedReplacements.find(
      (candidate) => source.startsWith(candidate.token, index) && isFormulaTokenBoundary(source, index, candidate.token),
    )
    if (replacement !== undefined) {
      output += replacement.replacement
      index += replacement.token.length
      continue
    }

    const current = source[index]
    if (current === '"' || current === "'") {
      const endIndex = formulaQuoteEnd(source, index, current)
      output += source.slice(index, endIndex)
      index = endIndex
      continue
    }

    output += current
    index += 1
  }

  return output
}

function formulaForCell(
  engine: SpreadsheetEngine,
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  targetCells: readonly CellTarget[],
  cellIndex: number,
): string {
  const replacements = command.inputs.map((input) => ({
    token: workbookFormula.source(workbookFormula.ref(input)),
    replacement: inputReplacementForCell(engine, input, targetCells, cellIndex),
  }))
  return materializeFormulaSource(command.formula, replacements)
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

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalValue)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalValue(entry)]),
    )
  }
  return value
}

function opKey(op: EngineOp): string {
  return JSON.stringify(canonicalValue(op))
}

function directCommandOps(command: WorkbookActionCommand): readonly EngineOp[] {
  if (command.kind === 'op') {
    return [structuredClone(command.op)]
  }
  if (command.target.kind !== 'range' || command.target.range.startAddress !== command.target.range.endAddress) {
    return []
  }
  const cell = {
    sheetName: command.target.range.sheetName,
    address: command.target.range.startAddress,
  }
  switch (command.kind) {
    case 'writeFormula':
      return [{ kind: 'setCellFormula', ...cell, formula: command.formula }]
    case 'writeValue':
      return [{ kind: 'setCellValue', ...cell, value: command.value }]
    case 'clear':
      return [{ kind: 'clearCell', ...cell }]
    case 'format':
      return command.numberFormat === undefined ? [] : [{ kind: 'setCellFormat', ...cell, format: command.numberFormat }]
  }
}

function materializePlanOps(engine: SpreadsheetEngine, plan: WorkbookActionPlan): readonly EngineOp[] {
  if (plan.commands.length === 0) {
    return plan.ops
  }
  const commandOps = plan.commands.flatMap((command) => materializeCommandOps(engine, command))
  const commandOpKeys = new Set([...commandOps, ...plan.commands.flatMap(directCommandOps)].map(opKey))
  const additionalOps = plan.ops.filter((op) => !commandOpKeys.has(opKey(op)))
  return [...commandOps, ...additionalOps]
}

function verifyCheck(engine: SpreadsheetEngine, check: WorkbookCheckResult): WorkbookCheckResult {
  if (check.status === 'failed' || check.target === undefined) {
    return check
  }
  if (check.kind === 'exists') {
    const exists = refExists(engine, check.target)
    return checked(
      check,
      exists ? 'passed' : 'failed',
      runtimeProof(exists ? 'Runtime confirmed the reference exists' : 'Runtime could not resolve the reference', {
        exists,
        target: check.target.label,
      }),
    )
  }
  if (check.kind === 'noFormulaErrors') {
    const status = verifyNoFormulaErrors(engine, check.target)
    return checked(
      check,
      status,
      status === 'planned'
        ? undefined
        : runtimeProof(
            status === 'passed' ? 'Runtime confirmed no formula errors' : 'Runtime found formula errors or could not resolve formulas',
            {
              passed: status === 'passed',
              target: check.target.label,
            },
          ),
    )
  }
  return check
}

export function createWorkbookRunAdapter(engine: SpreadsheetEngine, options: WorkbookRunEngineAdapterOptions = {}): WorkbookRunAdapter {
  return {
    preview(plan: WorkbookActionPlan) {
      return {
        modelName: plan.modelName,
        actionName: plan.actionName,
        requirements: describeRuntimeRequirements(plan).requirements,
        materializedOps: materializePlanOps(engine, plan),
      }
    },
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
