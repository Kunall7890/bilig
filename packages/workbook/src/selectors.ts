import { formatAddress, parseCellAddress } from '@bilig/formula'
import { isLiteralInput, type CellRangeRef, type LiteralInput } from '@bilig/protocol'
import type { WorkbookRowsRef, WorkbookTableRef } from './find.js'

export type WorkbookRowOperator = 'eq' | 'neq' | 'contains' | 'startsWith' | 'gt' | 'gte' | 'lt' | 'lte'
export type WorkbookRowValueType = 'number' | 'string' | 'boolean' | 'null'

export const workbookRowOperators = Object.freeze([
  'eq',
  'neq',
  'contains',
  'startsWith',
  'gt',
  'gte',
  'lt',
  'lte',
] as const satisfies readonly WorkbookRowOperator[])
const WORKBOOK_ROW_OPERATOR_SET = new Set<string>(workbookRowOperators)

export const workbookRowOperatorValueTypes = Object.freeze({
  eq: Object.freeze(['number', 'string', 'boolean', 'null']),
  neq: Object.freeze(['number', 'string', 'boolean', 'null']),
  contains: Object.freeze(['string']),
  startsWith: Object.freeze(['string']),
  gt: Object.freeze(['number', 'string']),
  gte: Object.freeze(['number', 'string']),
  lt: Object.freeze(['number', 'string']),
  lte: Object.freeze(['number', 'string']),
} as const satisfies Record<WorkbookRowOperator, readonly WorkbookRowValueType[]>)

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
  readonly sheetName?: string
  readonly table?: WorkbookTableRef
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

interface OptionalDataValue {
  readonly status: 'missing' | 'present'
  readonly value?: unknown
}

export interface NormalizedFindTableOptions {
  readonly name?: string
  readonly sheetName?: string
  readonly headers?: readonly string[]
}

export interface NormalizedFindColumnOptions {
  readonly table: unknown
  readonly rows?: unknown
  readonly name: string
}

export interface NormalizedFindRowsOptions {
  readonly sheetName?: string
  readonly table?: unknown
  readonly where: {
    readonly column: string
    readonly op: WorkbookRowOperator
    readonly value: LiteralInput
  }
}

function selectorObject(value: unknown, label: string): object {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${label} must be an object`)
  }
  return value
}

function optionalDataValue(value: object, key: string, label: string): OptionalDataValue {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return { status: 'missing' }
  }
  if (!('value' in descriptor)) {
    throw new Error(`${label} must be a data property`)
  }
  return {
    status: 'present',
    value: descriptor.value,
  }
}

function requiredDataValue(value: object, key: string, label: string): unknown {
  const property = optionalDataValue(value, key, label)
  if (property.status === 'missing') {
    throw new Error(`${label} is required`)
  }
  return property.value
}

export function cleanIdPart(value: string): string {
  const cleaned = value
    .trim()
    .replaceAll(/[^A-Za-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '')
  return cleaned === '' ? 'ref' : cleaned
}

export function requiredSelectorText(value: unknown, field: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Workbook selector ${field} must be a string`)
  }
  const trimmed = value.trim()
  if (trimmed === '') {
    throw new Error(`Workbook selector ${field} cannot be empty`)
  }
  return trimmed
}

function optionalSelectorText(property: OptionalDataValue, field: string): string | undefined {
  if (property.status === 'missing') {
    return undefined
  }
  return requiredSelectorText(property.value, field)
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
  const normalized: string[] = []
  for (let index = 0; index < headers.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(headers, String(index))
    if (descriptor === undefined || !descriptor.enumerable) {
      throw new Error('Workbook table headers must contain only data strings')
    }
    if (!('value' in descriptor)) {
      throw new Error(`Workbook table headers[${index}] must be a data property`)
    }
    normalized.push(requiredSelectorText(descriptor.value, 'table header'))
  }
  normalized.sort((left, right) => left.localeCompare(right))
  for (let index = 1; index < normalized.length; index += 1) {
    if (normalized[index] === normalized[index - 1]) {
      throw new Error(`Workbook table headers cannot contain duplicates: ${normalized[index]}`)
    }
  }
  return Object.freeze(normalized)
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

export function isWorkbookRowOperator(value: unknown): value is WorkbookRowOperator {
  return typeof value === 'string' && WORKBOOK_ROW_OPERATOR_SET.has(value)
}

function normalizeRowOperator(value: unknown): WorkbookRowOperator {
  if (!isWorkbookRowOperator(value)) {
    throw new Error(`Unsupported workbook row operator: ${String(value)}`)
  }
  return value
}

function workbookRowValueType(value: LiteralInput): WorkbookRowValueType {
  if (value === null) {
    return 'null'
  }
  if (typeof value === 'number') {
    return 'number'
  }
  if (typeof value === 'string') {
    return 'string'
  }
  return 'boolean'
}

export function isWorkbookRowValueCompatible(op: WorkbookRowOperator, value: LiteralInput): boolean {
  const accepted = workbookRowOperatorValueTypes[op] as readonly WorkbookRowValueType[]
  return accepted.includes(workbookRowValueType(value))
}

function acceptedValueTypeLabel(types: readonly WorkbookRowValueType[]): string {
  if (types.length === 1) {
    return types[0]!
  }
  if (types.length === 2) {
    return `${types[0]} or ${types[1]}`
  }
  return `${types.slice(0, -1).join(', ')}, or ${types[types.length - 1]!}`
}

function normalizeRowsValue(value: unknown, op: WorkbookRowOperator): LiteralInput {
  if (!isLiteralInput(value)) {
    throw new Error('Workbook rows selector value must be a finite JSON literal')
  }
  if (!isWorkbookRowValueCompatible(op, value)) {
    throw new Error(`Workbook rows selector operator ${op} requires a ${acceptedValueTypeLabel(workbookRowOperatorValueTypes[op])} value`)
  }
  return value
}

export function literalIdPart(value: LiteralInput): string {
  const source = JSON.stringify(value) ?? 'null'
  let encoded = ''
  for (const char of source) {
    encoded += /[A-Za-z0-9]/.test(char) ? char : `_${char.charCodeAt(0).toString(16)}`
  }
  return cleanIdPart(`${typeof value}_${encoded}`)
}

export function literalLabel(value: LiteralInput): string {
  return JSON.stringify(value)
}

export function normalizeRangeRef(input: FindRangeInput): CellRangeRef {
  const source = selectorObject(input, 'Workbook range selector')
  const startAddress = optionalDataValue(source, 'startAddress', 'Workbook range selector startAddress')
  if (startAddress.status === 'present') {
    const sheetName = requiredSelectorText(requiredDataValue(source, 'sheetName', 'Workbook range selector sheetName'), 'sheet name')
    const start = normalizeCellAddress(startAddress.value, 'startAddress')
    const endAddress = optionalDataValue(source, 'endAddress', 'Workbook range selector endAddress')
    const end = normalizeCellAddress(endAddress.status === 'present' ? endAddress.value : startAddress.value, 'endAddress')
    assertRangeOrder(start, end)
    return Object.freeze({
      sheetName,
      startAddress: start.text,
      endAddress: end.text,
    })
  }

  const address = optionalDataValue(source, 'address', 'Workbook range selector address')
  if (address.status === 'present') {
    const sheetName = requiredSelectorText(requiredDataValue(source, 'sheetName', 'Workbook range selector sheetName'), 'sheet name')
    const normalizedAddress = normalizeCellAddress(address.value, 'address')
    return Object.freeze({
      sheetName,
      startAddress: normalizedAddress.text,
      endAddress: normalizedAddress.text,
    })
  }

  throw new Error('Workbook range selector must include address or startAddress')
}

export function normalizeTableSelector(options: FindTableOptions): NormalizedFindTableOptions {
  const source = selectorObject(options, 'Workbook table selector')
  const tableName = optionalSelectorText(optionalDataValue(source, 'name', 'Workbook table selector name'), 'table name')
  const sheetName = optionalSelectorText(optionalDataValue(source, 'sheetName', 'Workbook table selector sheetName'), 'sheet name')
  const headers = normalizeHeaders(optionalDataValue(source, 'headers', 'Workbook table selector headers').value)
  if (tableName === undefined && sheetName === undefined && headers === undefined) {
    throw new Error('Workbook table selector needs a name, sheet name, or headers')
  }
  return {
    ...(tableName !== undefined ? { name: tableName } : {}),
    ...(sheetName !== undefined ? { sheetName } : {}),
    ...(headers !== undefined ? { headers } : {}),
  }
}

export function normalizeColumnSelector(options: FindColumnOptions): NormalizedFindColumnOptions {
  const source = selectorObject(options, 'Workbook column selector')
  const table = requiredDataValue(source, 'table', 'Workbook column selector table')
  const rows = optionalDataValue(source, 'rows', 'Workbook column selector rows')
  const name = requiredSelectorText(requiredDataValue(source, 'name', 'Workbook column selector name'), 'column name')
  return {
    table,
    ...(rows.status === 'present' ? { rows: rows.value } : {}),
    name,
  }
}

export function normalizeRowsSelector(options: FindRowsOptions): NormalizedFindRowsOptions {
  const source = selectorObject(options, 'Workbook rows selector')
  const sheetName = optionalSelectorText(optionalDataValue(source, 'sheetName', 'Workbook rows selector sheetName'), 'sheet name')
  const table = optionalDataValue(source, 'table', 'Workbook rows selector table')
  if (table.status === 'missing' && sheetName === undefined) {
    throw new Error('Workbook rows selector requires a table or sheet name')
  }
  const where = selectorObject(requiredDataValue(source, 'where', 'Workbook rows selector where'), 'Workbook rows selector where')
  const column = requiredSelectorText(requiredDataValue(where, 'column', 'Workbook rows selector where.column'), 'row column')
  const op = normalizeRowOperator(requiredDataValue(where, 'op', 'Workbook rows selector where.op'))
  return {
    ...(sheetName !== undefined ? { sheetName } : {}),
    ...(table.status === 'present' ? { table: table.value } : {}),
    where: Object.freeze({
      column,
      op,
      value: normalizeRowsValue(requiredDataValue(where, 'value', 'Workbook rows selector where.value'), op),
    }),
  }
}
