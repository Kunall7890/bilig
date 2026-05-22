import { formatAddress, parseCellAddress } from '@bilig/formula'
import { isLiteralInput, type CellRangeRef, type LiteralInput } from '@bilig/protocol'

export type WorkbookRefKind = 'range' | 'name' | 'table' | 'column' | 'rows'

export interface WorkbookBaseRef {
  readonly kind: WorkbookRefKind
  readonly id: string
  readonly label: string
}

export interface WorkbookRangeRef extends WorkbookBaseRef {
  readonly kind: 'range'
  readonly range: CellRangeRef
}

export interface WorkbookNameRef extends WorkbookBaseRef {
  readonly kind: 'name'
  readonly name: string
}

export interface WorkbookTableRef extends WorkbookBaseRef {
  readonly kind: 'table'
  readonly name?: string
  readonly sheetName?: string
  readonly headers?: readonly string[]
  readonly column: (name: string) => WorkbookColumnRef
}

export interface WorkbookColumnRef extends WorkbookBaseRef {
  readonly kind: 'column'
  readonly table: WorkbookTableRef
  readonly rows?: WorkbookRowsRef
  readonly name: string
}

export type WorkbookRowOperator = 'eq' | 'neq' | 'contains' | 'startsWith' | 'gt' | 'gte' | 'lt' | 'lte'

const WORKBOOK_ROW_OPERATORS: readonly WorkbookRowOperator[] = ['eq', 'neq', 'contains', 'startsWith', 'gt', 'gte', 'lt', 'lte']
const WORKBOOK_ROW_OPERATOR_SET = new Set<string>(WORKBOOK_ROW_OPERATORS)
const REF_ID_SEPARATOR = '_p_'

function hasOwnString<Key extends string>(value: object, key: Key): value is Record<Key, string> {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && typeof descriptor.value === 'string'
}

function hasOwnFunction<Key extends string>(value: object, key: Key): value is Record<Key, (...args: never[]) => unknown> {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && typeof descriptor.value === 'function'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function selectorObject(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`Workbook ${label} must be an object`)
  }
  return value
}

export interface WorkbookRowsRef extends WorkbookBaseRef {
  readonly kind: 'rows'
  readonly table: WorkbookTableRef
  readonly where: {
    readonly column: string
    readonly op: WorkbookRowOperator
    readonly value: LiteralInput
  }
  readonly column: (name: string) => WorkbookColumnRef
}

export type WorkbookRef = WorkbookRangeRef | WorkbookNameRef | WorkbookTableRef | WorkbookColumnRef | WorkbookRowsRef

const WORKBOOK_REF_KINDS = new Set<string>(['range', 'name', 'table', 'column', 'rows'])

function hasValidBaseRef(value: object): value is WorkbookBaseRef {
  return hasOwnString(value, 'kind') && WORKBOOK_REF_KINDS.has(value.kind) && hasOwnString(value, 'id') && hasOwnString(value, 'label')
}

function hasOptionalString(value: object, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor === undefined || typeof descriptor.value === 'string'
}

function hasOptionalStringArray(value: object, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor === undefined || (Array.isArray(descriptor.value) && descriptor.value.every((entry) => typeof entry === 'string'))
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    typeof value === 'object' &&
    value !== null &&
    hasOwnString(value, 'sheetName') &&
    hasOwnString(value, 'startAddress') &&
    hasOwnString(value, 'endAddress')
  )
}

function isWorkbookTableRef(value: unknown): value is WorkbookTableRef {
  return (
    typeof value === 'object' &&
    value !== null &&
    hasValidBaseRef(value) &&
    value.kind === 'table' &&
    hasOptionalString(value, 'name') &&
    hasOptionalString(value, 'sheetName') &&
    hasOptionalStringArray(value, 'headers') &&
    hasOwnFunction(value, 'column')
  )
}

function isWorkbookRowsRef(value: unknown): value is WorkbookRowsRef {
  if (typeof value !== 'object' || value === null || !hasValidBaseRef(value) || value.kind !== 'rows' || !hasOwnFunction(value, 'column')) {
    return false
  }
  const where = Object.getOwnPropertyDescriptor(value, 'where')?.value
  if (
    typeof where !== 'object' ||
    where === null ||
    !hasOwnString(where, 'column') ||
    !isWorkbookRowOperator(Object.getOwnPropertyDescriptor(where, 'op')?.value) ||
    !isLiteralInput(Object.getOwnPropertyDescriptor(where, 'value')?.value)
  ) {
    return false
  }
  const table = Object.getOwnPropertyDescriptor(value, 'table')?.value
  return isWorkbookTableRef(table)
}

function isWorkbookColumnRef(value: unknown): value is WorkbookColumnRef {
  if (typeof value !== 'object' || value === null || !hasValidBaseRef(value) || value.kind !== 'column' || !hasOwnString(value, 'name')) {
    return false
  }
  const table = Object.getOwnPropertyDescriptor(value, 'table')?.value
  const rows = Object.getOwnPropertyDescriptor(value, 'rows')?.value
  return isWorkbookTableRef(table) && (rows === undefined || isWorkbookRowsRef(rows))
}

export function isWorkbookRef(value: unknown): value is WorkbookRef {
  if (typeof value !== 'object' || value === null || !hasValidBaseRef(value)) {
    return false
  }
  switch (value.kind) {
    case 'range':
      return isCellRangeRef(Object.getOwnPropertyDescriptor(value, 'range')?.value)
    case 'name':
      return hasOwnString(value, 'name')
    case 'table':
      return isWorkbookTableRef(value)
    case 'column':
      return isWorkbookColumnRef(value)
    case 'rows':
      return isWorkbookRowsRef(value)
  }
}

export function collectWorkbookRefs(value: unknown): readonly WorkbookRef[] {
  const refs: WorkbookRef[] = []
  const seenRefs = new Set<string>()
  const seenObjects = new WeakSet<object>()

  function pushRef(ref: WorkbookRef): void {
    const key = `${ref.kind}:${ref.id}`
    if (!seenRefs.has(key)) {
      seenRefs.add(key)
      refs.push(ref)
    }
  }

  function visit(current: unknown): void {
    if (current === null || current === undefined || typeof current !== 'object') {
      return
    }
    if (seenObjects.has(current)) {
      return
    }
    seenObjects.add(current)

    if (isWorkbookRef(current)) {
      pushRef(current)
      if (current.kind === 'column') {
        visit(current.rows)
        visit(current.table)
      }
      if (current.kind === 'rows') {
        visit(current.table)
      }
      return
    }

    if (Array.isArray(current)) {
      current.forEach(visit)
      return
    }

    Object.values(current).forEach(visit)
  }

  visit(value)
  return refs
}

export interface FindTableOptions {
  readonly name?: string
  readonly sheetName?: string
  readonly headers?: readonly string[]
}

export interface FindColumnOptions {
  readonly table: WorkbookTableRef
  readonly rows?: WorkbookRowsRef
  readonly name: string
}

export interface FindRowsOptions {
  readonly table: WorkbookTableRef
  readonly where: {
    readonly column: string
    readonly op: WorkbookRowOperator
    readonly value: LiteralInput
  }
}

export interface FindRangeCellInput extends CellRangeRef {}

export type FindRangeInput =
  | FindRangeCellInput
  | {
      readonly sheetName: string
      readonly address: string
    }
  | {
      readonly sheetName: string
      readonly startAddress: string
      readonly endAddress?: string
    }

export interface WorkbookFindApi {
  readonly findTable: (options: FindTableOptions) => WorkbookTableRef
  readonly findColumn: (options: FindColumnOptions) => WorkbookColumnRef
  readonly findRange: (input: FindRangeInput) => WorkbookRangeRef
  readonly findName: (name: string) => WorkbookNameRef
  readonly findRows: (options: FindRowsOptions) => WorkbookRowsRef
}

export interface WorkbookFindNamespace extends WorkbookFindApi {
  readonly table: (options: FindTableOptions) => WorkbookTableRef
  readonly column: (options: FindColumnOptions) => WorkbookColumnRef
  readonly range: (input: FindRangeInput) => WorkbookRangeRef
  readonly name: (name: string) => WorkbookNameRef
  readonly rows: (options: FindRowsOptions) => WorkbookRowsRef
}

function refIdPart(value: string): string {
  let encoded = ''
  for (const char of value) {
    encoded += /^[A-Za-z0-9]$/.test(char) ? char : `_x${char.codePointAt(0)!.toString(16)}_`
  }
  return encoded === '' ? 'ref' : encoded
}

function refId(...parts: readonly (string | undefined)[]): string {
  return parts
    .filter((part) => part !== undefined)
    .map((part) => refIdPart(part))
    .join(REF_ID_SEPARATOR)
}

function joinRefId(...parts: readonly string[]): string {
  return parts.join(REF_ID_SEPARATOR)
}

function requiredSelectorText(value: unknown, field: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Workbook selector ${field} must be a string`)
  }
  const trimmed = value.trim()
  if (trimmed === '') {
    throw new Error(`Workbook selector ${field} cannot be empty`)
  }
  return trimmed
}

function optionalSelectorText(value: unknown, field: string): string | undefined {
  if (value === undefined) {
    return undefined
  }
  return requiredSelectorText(value, field)
}

function normalizeHeaders(headers: unknown): readonly string[] | undefined {
  if (headers === undefined) {
    return undefined
  }
  if (!Array.isArray(headers)) {
    throw new Error('Workbook table headers must be an array')
  }
  if (headers.length === 0) {
    throw new Error('Workbook table headers cannot be empty')
  }
  return headers.map((header) => requiredSelectorText(header, 'table header'))
}

interface NormalizedCellAddress {
  readonly row: number
  readonly col: number
  readonly text: string
}

function normalizeCellAddress(value: unknown, field: string): NormalizedCellAddress {
  const raw = requiredSelectorText(value, field)
  try {
    const parsed = parseCellAddress(raw)
    if (parsed.sheetName !== undefined) {
      throw new Error('qualified')
    }
    return {
      row: parsed.row,
      col: parsed.col,
      text: formatAddress(parsed.row, parsed.col),
    }
  } catch {
    throw new Error(`Workbook range ${field} is invalid: ${raw}`)
  }
}

function assertRangeOrder(start: NormalizedCellAddress, end: NormalizedCellAddress): void {
  if (end.row < start.row || end.col < start.col) {
    throw new Error('Workbook range endAddress must not be before startAddress')
  }
}

function isWorkbookRowOperator(value: unknown): value is WorkbookRowOperator {
  return typeof value === 'string' && WORKBOOK_ROW_OPERATOR_SET.has(value)
}

function normalizeRowOperator(value: unknown): WorkbookRowOperator {
  if (!isWorkbookRowOperator(value)) {
    throw new Error(`Unsupported workbook row operator: ${String(value)}`)
  }
  return value
}

function normalizeRowsValue(value: unknown): LiteralInput {
  if (!isLiteralInput(value)) {
    throw new Error('Workbook rows selector value must be a finite JSON literal')
  }
  return value
}

function literalIdPart(value: LiteralInput): string {
  const source = JSON.stringify(value) ?? 'null'
  return refId(value === null ? 'null' : typeof value, source)
}

function literalLabel(value: LiteralInput): string {
  return JSON.stringify(value)
}

function hideHelper<T extends object, Key extends keyof T & string>(target: T, key: Key): void {
  Object.defineProperty(target, key, {
    enumerable: false,
  })
}

function rangeLabel(range: CellRangeRef): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

export function normalizeRangeRef(input: FindRangeInput): CellRangeRef {
  const selector = selectorObject(input, 'range selector')
  if ('startAddress' in selector) {
    const sheetName = requiredSelectorText(selector['sheetName'], 'sheet name')
    const start = normalizeCellAddress(selector['startAddress'], 'startAddress')
    const end = normalizeCellAddress(selector['endAddress'] ?? selector['startAddress'], 'endAddress')
    assertRangeOrder(start, end)
    return {
      sheetName,
      startAddress: start.text,
      endAddress: end.text,
    }
  }
  if ('address' in selector) {
    const sheetName = requiredSelectorText(selector['sheetName'], 'sheet name')
    const address = normalizeCellAddress(selector['address'], 'address')
    return {
      sheetName,
      startAddress: address.text,
      endAddress: address.text,
    }
  }
  throw new Error('Workbook range selector must include address or startAddress')
}

export function createWorkbookTableRef(options: FindTableOptions): WorkbookTableRef {
  const selector = selectorObject(options, 'table selector')
  const tableName = optionalSelectorText(selector['name'], 'table name')
  const sheetName = optionalSelectorText(selector['sheetName'], 'sheet name')
  const headers = normalizeHeaders(selector['headers'])
  if (tableName === undefined && sheetName === undefined && headers === undefined) {
    throw new Error('Workbook table selector needs a name, sheet name, or headers')
  }
  const id = refId('table', sheetName, tableName, ...(headers ?? []))
  const label = tableName ?? (headers ? `table with ${headers.join(', ')}` : `${sheetName} table`)
  const table: WorkbookTableRef = {
    kind: 'table',
    id,
    label,
    ...(tableName !== undefined ? { name: tableName } : {}),
    ...(sheetName !== undefined ? { sheetName } : {}),
    ...(headers !== undefined ? { headers: Object.freeze([...headers]) } : {}),
    column(name) {
      return createWorkbookColumnRef({ table, name })
    },
  }
  hideHelper(table, 'column')
  return Object.freeze(table)
}

export function createWorkbookColumnRef(options: FindColumnOptions): WorkbookColumnRef {
  const selector = selectorObject(options, 'column selector')
  const table = selector['table']
  if (!isWorkbookTableRef(table)) {
    throw new Error('Workbook column selector requires a table')
  }
  const rows = selector['rows']
  if (rows !== undefined && !isWorkbookRowsRef(rows)) {
    throw new Error('Workbook column selector rows must be a rows ref')
  }
  if (rows !== undefined && rows.table.id !== table.id) {
    throw new Error('Workbook column selector rows must belong to the table')
  }
  const name = requiredSelectorText(selector['name'], 'column name')
  const ownerId = rows?.id ?? table.id
  const ownerLabel = rows?.label ?? table.label
  const column: WorkbookColumnRef = {
    kind: 'column',
    id: joinRefId(ownerId, refIdPart(name)),
    label: `${ownerLabel}.${name}`,
    table,
    ...(rows !== undefined ? { rows } : {}),
    name,
  }
  return Object.freeze(column)
}

export function createWorkbookRangeRef(input: FindRangeInput): WorkbookRangeRef {
  const range = normalizeRangeRef(input)
  return Object.freeze({
    kind: 'range',
    id: refId('range', range.sheetName, range.startAddress, range.endAddress),
    label: rangeLabel(range),
    range: Object.freeze(range),
  })
}

export function createWorkbookNameRef(name: string): WorkbookNameRef {
  const normalizedName = requiredSelectorText(name, 'name')
  return Object.freeze({
    kind: 'name',
    id: refId('name', normalizedName),
    label: normalizedName,
    name: normalizedName,
  })
}

export function createWorkbookRowsRef(options: FindRowsOptions): WorkbookRowsRef {
  const selector = selectorObject(options, 'rows selector')
  const table = selector['table']
  if (!isWorkbookTableRef(table)) {
    throw new Error('Workbook rows selector requires a table')
  }
  const rawWhere = selector['where']
  if (!isRecord(rawWhere)) {
    throw new Error('Workbook rows selector requires a where object')
  }
  const where = {
    column: requiredSelectorText(rawWhere['column'], 'row column'),
    op: normalizeRowOperator(rawWhere['op']),
    value: normalizeRowsValue(rawWhere['value']),
  }
  const ownerId = table.id
  const ownerLabel = table.label
  const rows: WorkbookRowsRef = {
    kind: 'rows',
    id: joinRefId(ownerId, refIdPart(where.column), refIdPart(where.op), literalIdPart(where.value)),
    label: `${ownerLabel} rows where ${where.column} ${where.op} ${literalLabel(where.value)}`,
    table,
    where: Object.freeze(where),
    column(name) {
      return createWorkbookColumnRef({ table: rows.table, rows, name })
    },
  }
  hideHelper(rows, 'column')
  return Object.freeze(rows)
}

export function findTable(options: FindTableOptions): WorkbookTableRef {
  return createWorkbookTableRef(options)
}

export function findColumn(options: FindColumnOptions): WorkbookColumnRef {
  return createWorkbookColumnRef(options)
}

export function findRange(input: FindRangeInput): WorkbookRangeRef {
  return createWorkbookRangeRef(input)
}

export function findName(name: string): WorkbookNameRef {
  return createWorkbookNameRef(name)
}

export function findRows(options: FindRowsOptions): WorkbookRowsRef {
  return createWorkbookRowsRef(options)
}

export function createWorkbookFindApi(): WorkbookFindApi {
  return {
    findTable,
    findColumn,
    findRange,
    findName,
    findRows,
  }
}

export const find: WorkbookFindNamespace = Object.freeze({
  ...createWorkbookFindApi(),
  table: findTable,
  column: findColumn,
  range: findRange,
  name: findName,
  rows: findRows,
})
