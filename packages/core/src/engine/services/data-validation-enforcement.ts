import { ValueTag, type CellRangeRef, type LiteralInput, type WorkbookDataValidationRuleSnapshot } from '@bilig/protocol'
import { formatAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import type {
  EngineCellMutationRef,
  EngineExistingLiteralCellMutationRef,
  EngineExistingNumericCellMutationRef,
} from '../../cell-mutations-at.js'
import type { WorkbookStore, WorkbookTableRecord } from '../../workbook-store.js'
import { normalizeRange } from './operation-change-helpers.js'

export interface DataValidationEnforcementState {
  readonly workbook: WorkbookStore
  readonly strings: {
    get(id: number): string
  }
}

export function assertLocalDataValidationsForEngineOps(state: DataValidationEnforcementState, ops: readonly EngineOp[]): void {
  if (!state.workbook.hasDataValidations()) {
    return
  }
  for (const op of ops) {
    if (op.kind === 'setCellValue') {
      const bounds = normalizeRange({ sheetName: op.sheetName, startAddress: op.address, endAddress: op.address })
      assertLocalDataValidationForCellValue(state, op.sheetName, bounds.startRow, bounds.startCol, op.value)
    }
  }
}

export function assertLocalDataValidationsForCellMutationRefs(
  state: DataValidationEnforcementState,
  refs: readonly EngineCellMutationRef[],
): void {
  if (!state.workbook.hasDataValidations()) {
    return
  }
  for (const ref of refs) {
    if (ref.mutation.kind !== 'setCellValue') {
      continue
    }
    const sheetName = state.workbook.getSheetById(ref.sheetId)?.name
    if (sheetName === undefined) {
      throw new Error(`Unknown sheet id: ${String(ref.sheetId)}`)
    }
    assertLocalDataValidationForCellValue(state, sheetName, ref.mutation.row, ref.mutation.col, ref.mutation.value)
  }
}

export function assertLocalDataValidationForExistingNumericCellMutation(
  state: DataValidationEnforcementState,
  request: EngineExistingNumericCellMutationRef,
): void {
  if (!state.workbook.hasDataValidations()) {
    return
  }
  const sheetName = state.workbook.getSheetById(request.sheetId)?.name
  if (sheetName === undefined) {
    throw new Error(`Unknown sheet id: ${String(request.sheetId)}`)
  }
  assertLocalDataValidationForCellValue(state, sheetName, request.row, request.col, request.value)
}

export function assertLocalDataValidationForExistingLiteralCellMutation(
  state: DataValidationEnforcementState,
  request: EngineExistingLiteralCellMutationRef,
): void {
  if (!state.workbook.hasDataValidations()) {
    return
  }
  const sheetName = state.workbook.getSheetById(request.sheetId)?.name
  if (sheetName === undefined) {
    throw new Error(`Unknown sheet id: ${String(request.sheetId)}`)
  }
  assertLocalDataValidationForCellValue(state, sheetName, request.row, request.col, request.value)
}

function assertLocalDataValidationForCellValue(
  state: DataValidationEnforcementState,
  sheetName: string,
  row: number,
  col: number,
  value: LiteralInput,
): void {
  for (const validation of state.workbook.listDataValidations(sheetName)) {
    if (!rangeContainsCell(validation.range, sheetName, row, col)) {
      continue
    }
    if (isBlankLiteral(value) && validation.allowBlank !== false) {
      continue
    }
    if (ruleAllowsLiteral(state, validation.range.sheetName, validation.rule, value)) {
      continue
    }
    const address = formatAddress(row, col)
    throw new Error(
      `Excel data validation rejects ${sheetName}!${address} value ${formatLiteralForError(value)} for ${validation.rule.kind} rule`,
    )
  }
}

function rangeContainsCell(range: CellRangeRef, sheetName: string, row: number, col: number): boolean {
  if (range.sheetName !== sheetName) {
    return false
  }
  const bounds = normalizeRange(range)
  return row >= bounds.startRow && row <= bounds.endRow && col >= bounds.startCol && col <= bounds.endCol
}

function ruleAllowsLiteral(
  state: DataValidationEnforcementState,
  targetSheetName: string,
  rule: WorkbookDataValidationRuleSnapshot,
  value: LiteralInput,
): boolean {
  switch (rule.kind) {
    case 'any':
      return true
    case 'checkbox':
      return literalsEqual(value, rule.checkedValue ?? true) || literalsEqual(value, rule.uncheckedValue ?? false)
    case 'list':
      return listRuleAllowsLiteral(state, targetSheetName, rule, value)
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength':
      return scalarRuleAllowsLiteral(rule, value)
  }
}

function listRuleAllowsLiteral(
  state: DataValidationEnforcementState,
  targetSheetName: string,
  rule: Extract<WorkbookDataValidationRuleSnapshot, { kind: 'list' }>,
  value: LiteralInput,
): boolean {
  const values = [...(rule.values ?? [])]
  if (rule.source !== undefined) {
    values.push(...resolveListSourceValues(state, targetSheetName, rule.source))
  }
  const key = listComparisonKey(value)
  return values.some((candidate) => listComparisonKey(candidate) === key)
}

function scalarRuleAllowsLiteral(
  rule: Extract<WorkbookDataValidationRuleSnapshot, { kind: 'whole' | 'decimal' | 'date' | 'time' | 'textLength' }>,
  value: LiteralInput,
): boolean {
  const scalar = validationScalarValue(rule.kind, value)
  if (scalar === null) {
    return false
  }
  if (rule.kind === 'whole' && !Number.isInteger(scalar)) {
    return false
  }
  const first = validationScalarRuleValue(rule.kind, rule.values[0] ?? null)
  if (first === null) {
    return false
  }
  if (rule.operator === 'between' || rule.operator === 'notBetween') {
    const second = validationScalarRuleValue(rule.kind, rule.values[1] ?? null)
    if (second === null) {
      return false
    }
    const lower = Math.min(first, second)
    const upper = Math.max(first, second)
    const isBetween = scalar >= lower && scalar <= upper
    return rule.operator === 'between' ? isBetween : !isBetween
  }
  switch (rule.operator) {
    case 'equal':
      return scalar === first
    case 'notEqual':
      return scalar !== first
    case 'greaterThan':
      return scalar > first
    case 'greaterThanOrEqual':
      return scalar >= first
    case 'lessThan':
      return scalar < first
    case 'lessThanOrEqual':
      return scalar <= first
  }
}

function resolveListSourceValues(
  state: DataValidationEnforcementState,
  targetSheetName: string,
  source: Extract<WorkbookDataValidationRuleSnapshot, { kind: 'list' }>['source'],
): LiteralInput[] {
  if (source === undefined) {
    return []
  }
  switch (source.kind) {
    case 'cell-ref':
      return readRangeValues(state, {
        sheetName: source.sheetName,
        startAddress: source.address,
        endAddress: source.address,
      })
    case 'range-ref':
      return readRangeValues(state, source)
    case 'structured-ref':
      return readStructuredReferenceValues(state, source.tableName, source.columnName)
    case 'named-range': {
      const definedName = state.workbook.getDefinedName(source.name, targetSheetName)
      if (definedName === undefined) {
        throw new Error(`Excel data validation list source named range not found: ${source.name}`)
      }
      return readDefinedNameValues(state, targetSheetName, source.name, definedName.value)
    }
  }
}

function readDefinedNameValues(
  state: DataValidationEnforcementState,
  targetSheetName: string,
  name: string,
  value: LiteralInput | { kind: string },
): LiteralInput[] {
  if (typeof value !== 'object' || value === null) {
    return [value]
  }
  switch (value.kind) {
    case 'scalar':
      return ['value' in value && isLiteralInput(value.value) ? value.value : null]
    case 'cell-ref':
      return 'sheetName' in value && 'address' in value && typeof value.sheetName === 'string' && typeof value.address === 'string'
        ? readRangeValues(state, { sheetName: value.sheetName, startAddress: value.address, endAddress: value.address })
        : []
    case 'range-ref':
      return 'sheetName' in value &&
        'startAddress' in value &&
        'endAddress' in value &&
        typeof value.sheetName === 'string' &&
        typeof value.startAddress === 'string' &&
        typeof value.endAddress === 'string'
        ? readRangeValues(state, { sheetName: value.sheetName, startAddress: value.startAddress, endAddress: value.endAddress })
        : []
    case 'structured-ref':
      return 'tableName' in value && 'columnName' in value && typeof value.tableName === 'string' && typeof value.columnName === 'string'
        ? readStructuredReferenceValues(state, value.tableName, value.columnName)
        : []
    case 'formula':
      throw new Error(`Excel data validation list source named range is formula-backed and cannot be enforced: ${name}`)
    default:
      throw new Error(`Excel data validation list source named range has unsupported value in ${targetSheetName}: ${name}`)
  }
}

function readStructuredReferenceValues(state: DataValidationEnforcementState, tableName: string, columnName: string): LiteralInput[] {
  const table = state.workbook.getTable(tableName)
  if (table === undefined) {
    throw new Error(`Excel data validation list source table not found: ${tableName}`)
  }
  const columnIndex = table.columnNames.indexOf(columnName)
  if (columnIndex < 0) {
    throw new Error(`Excel data validation list source table column not found: ${tableName}[${columnName}]`)
  }
  return readTableColumnValues(state, table, columnIndex)
}

function readTableColumnValues(state: DataValidationEnforcementState, table: WorkbookTableRecord, columnIndex: number): LiteralInput[] {
  const bounds = normalizeRange({
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
  })
  const rowStart = table.headerRow ? bounds.startRow + 1 : bounds.startRow
  const rowEnd = table.totalsRow ? bounds.endRow - 1 : bounds.endRow
  const col = bounds.startCol + columnIndex
  if (rowStart > rowEnd || col > bounds.endCol) {
    return []
  }
  return readRangeValues(state, {
    sheetName: table.sheetName,
    startAddress: formatAddress(rowStart, col),
    endAddress: formatAddress(rowEnd, col),
  })
}

function readRangeValues(state: DataValidationEnforcementState, range: CellRangeRef): LiteralInput[] {
  const sheet = state.workbook.getSheet(range.sheetName)
  if (sheet === undefined) {
    throw new Error(`Excel data validation list source sheet not found: ${range.sheetName}`)
  }
  const bounds = normalizeRange(range)
  const values: LiteralInput[] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
      const cellIndex = state.workbook.getCellIndex(range.sheetName, formatAddress(row, col))
      const literal = cellIndex === undefined ? null : readCellLiteral(state, cellIndex)
      if (literal !== undefined) {
        values.push(literal)
      }
    }
  }
  return values
}

function readCellLiteral(state: DataValidationEnforcementState, cellIndex: number): LiteralInput | undefined {
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] ?? ValueTag.Empty) as ValueTag
  switch (tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return cellStore.numbers[cellIndex] ?? 0
    case ValueTag.Boolean:
      return (cellStore.numbers[cellIndex] ?? 0) !== 0
    case ValueTag.String: {
      const stringId = cellStore.stringIds[cellIndex] ?? 0
      return stringId === 0 ? '' : state.strings.get(stringId)
    }
    case ValueTag.Error:
      return undefined
  }
  return undefined
}

function validationScalarValue(kind: 'whole' | 'decimal' | 'date' | 'time' | 'textLength', value: LiteralInput): number | null {
  if (kind === 'textLength') {
    return literalDisplayText(value).length
  }
  if (kind === 'date') {
    return literalDateSerial(value)
  }
  if (kind === 'time') {
    return literalTimeSerial(value)
  }
  return literalNumber(value)
}

function validationScalarRuleValue(kind: 'whole' | 'decimal' | 'date' | 'time' | 'textLength', value: LiteralInput): number | null {
  return kind === 'textLength' ? literalNumber(value) : validationScalarValue(kind, value)
}

function literalNumber(value: LiteralInput): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null
  }
  if (typeof value === 'boolean') {
    return value ? 1 : 0
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (trimmed.length === 0) {
      return null
    }
    const parsed = Number(trimmed)
    return Number.isFinite(parsed) ? parsed : null
  }
  return null
}

function literalDateSerial(value: LiteralInput): number | null {
  const numeric = literalNumber(value)
  if (numeric !== null) {
    return numeric
  }
  if (typeof value !== 'string') {
    return null
  }
  const parsed = new Date(value)
  const timestamp = parsed.getTime()
  if (!Number.isFinite(timestamp)) {
    return null
  }
  const midnightUtc = Date.UTC(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate())
  return (midnightUtc - Date.UTC(1899, 11, 30)) / 86_400_000
}

function literalTimeSerial(value: LiteralInput): number | null {
  const numeric = literalNumber(value)
  if (numeric !== null) {
    return numeric
  }
  if (typeof value !== 'string') {
    return null
  }
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?$/u.exec(value.trim())
  if (match === null) {
    return null
  }
  const hours = Number(match[1])
  const minutes = Number(match[2])
  const seconds = match[3] === undefined ? 0 : Number(match[3])
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59 || seconds < 0 || seconds >= 60) {
    return null
  }
  return (hours * 3600 + minutes * 60 + seconds) / 86_400
}

function isBlankLiteral(value: LiteralInput): boolean {
  return value === null || value === ''
}

function listComparisonKey(value: LiteralInput): string {
  return literalDisplayText(value).trim().toUpperCase()
}

function literalDisplayText(value: LiteralInput): string {
  if (value === null) {
    return ''
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return String(value)
}

function literalsEqual(left: LiteralInput, right: LiteralInput): boolean {
  return left === right
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function formatLiteralForError(value: LiteralInput): string {
  if (value === null) {
    return '<blank>'
  }
  if (typeof value === 'string') {
    return JSON.stringify(value)
  }
  return String(value)
}
