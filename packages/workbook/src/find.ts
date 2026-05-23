import { isLiteralInput, type CellRangeRef, type LiteralInput } from '@bilig/protocol'
import {
  cleanIdPart,
  literalIdPart,
  literalLabel,
  normalizeColumnSelector,
  normalizeRangeRef,
  normalizeRowsSelector,
  normalizeTableSelector,
  requiredSelectorText,
  isWorkbookRowOperator,
  type FindColumnOptions,
  type FindRangeInput,
  type FindRowsOptions,
  type FindTableOptions,
  type WorkbookRowOperator,
} from './selectors.js'

export {
  isWorkbookRowOperator,
  isWorkbookRowValueCompatible,
  normalizeRangeRef,
  workbookRowOperators,
  workbookRowOperatorValueTypes,
} from './selectors.js'
export type {
  FindColumnOptions,
  FindRangeCellInput,
  FindRangeInput,
  FindRowsOptions,
  FindTableOptions,
  WorkbookRowOperator,
  WorkbookRowValueType,
} from './selectors.js'

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

function hasOwnString<Key extends string>(value: object, key: Key): value is Record<Key, string> {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && typeof descriptor.value === 'string'
}

function hasOwnFunction<Key extends string>(value: object, key: Key): value is Record<Key, (...args: never[]) => unknown> {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && typeof descriptor.value === 'function'
}

export interface WorkbookRowsRef extends WorkbookBaseRef {
  readonly kind: 'rows'
  readonly sheetName?: string
  readonly table?: WorkbookTableRef
  readonly where: {
    readonly column: string
    readonly op: WorkbookRowOperator
    readonly value: LiteralInput
  }
  readonly column: (name: string) => WorkbookColumnRef
}

export type WorkbookRef = WorkbookRangeRef | WorkbookNameRef | WorkbookTableRef | WorkbookColumnRef | WorkbookRowsRef

export interface WorkbookBaseRefData {
  readonly kind: WorkbookRefKind
  readonly id: string
  readonly label: string
}

export interface WorkbookRangeRefData extends WorkbookBaseRefData {
  readonly kind: 'range'
  readonly range: CellRangeRef
}

export interface WorkbookNameRefData extends WorkbookBaseRefData {
  readonly kind: 'name'
  readonly name: string
}

export interface WorkbookTableRefData extends WorkbookBaseRefData {
  readonly kind: 'table'
  readonly name?: string
  readonly sheetName?: string
  readonly headers?: readonly string[]
}

export interface WorkbookColumnRefData extends WorkbookBaseRefData {
  readonly kind: 'column'
  readonly table: WorkbookTableRefData
  readonly rows?: WorkbookRowsRefData
  readonly name: string
}

export interface WorkbookRowsRefData extends WorkbookBaseRefData {
  readonly kind: 'rows'
  readonly sheetName?: string
  readonly table?: WorkbookTableRefData
  readonly where: {
    readonly column: string
    readonly op: WorkbookRowOperator
    readonly value: LiteralInput
  }
}

export type WorkbookRefData =
  | WorkbookRangeRefData
  | WorkbookNameRefData
  | WorkbookTableRefData
  | WorkbookColumnRefData
  | WorkbookRowsRefData

export const workbookRefKinds = Object.freeze(['range', 'name', 'table', 'column', 'rows'] as const satisfies readonly WorkbookRefKind[])
const WORKBOOK_REF_KIND_SET = new Set<string>(workbookRefKinds)

export function isWorkbookRefKind(value: unknown): value is WorkbookRefKind {
  return typeof value === 'string' && WORKBOOK_REF_KIND_SET.has(value)
}

function hasValidBaseRef(value: object): value is WorkbookBaseRef {
  return hasOwnString(value, 'kind') && isWorkbookRefKind(value.kind) && hasOwnString(value, 'id') && hasOwnString(value, 'label')
}

function hasOptionalString(value: object, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor === undefined || typeof descriptor.value === 'string'
}

function isString(value: unknown): value is string {
  return typeof value === 'string'
}

function arrayEveryData<T>(value: unknown, predicate: (entry: unknown) => entry is T): value is readonly T[]
function arrayEveryData(value: unknown, predicate: (entry: unknown) => boolean): boolean
function arrayEveryData(value: unknown, predicate: (entry: unknown) => boolean): boolean {
  if (!Array.isArray(value)) {
    return false
  }

  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor) || !predicate(descriptor.value)) {
      return false
    }
  }

  return true
}

function ownArrayDataValues(value: readonly unknown[]): readonly unknown[] {
  const entries: unknown[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor !== undefined && descriptor.enumerable && 'value' in descriptor) {
      entries.push(descriptor.value)
    }
  }
  return entries
}

function copyStringArrayData(value: readonly string[]): string[] {
  const entries: string[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor) || typeof descriptor.value !== 'string') {
      throw new Error('Workbook string array data is invalid')
    }
    entries.push(descriptor.value)
  }
  return entries
}

function hasOptionalStringArray(value: object, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor === undefined || arrayEveryData(descriptor.value, isString)
}

function ownEnumerableDataValues(value: object): readonly unknown[] {
  const values: unknown[] = []
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if (descriptor.enumerable && 'value' in descriptor) {
      values.push(descriptor.value)
    }
  })
  return values
}

function ownEnumerableDataEntries(value: object): readonly (readonly [string, unknown])[] {
  const entries: [string, unknown][] = []
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (descriptor.enumerable && 'value' in descriptor) {
      entries.push([key, descriptor.value])
    }
  })
  return entries
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
  return isWorkbookTableRefData(value) && hasOwnFunction(value, 'column')
}

function isWorkbookTableRefData(value: unknown): value is WorkbookTableRefData {
  return (
    typeof value === 'object' &&
    value !== null &&
    hasValidBaseRef(value) &&
    value.kind === 'table' &&
    hasOptionalString(value, 'name') &&
    hasOptionalString(value, 'sheetName') &&
    hasOptionalStringArray(value, 'headers')
  )
}

function isWorkbookRowsRef(value: unknown): value is WorkbookRowsRef {
  return isWorkbookRowsRefData(value) && hasOwnFunction(value, 'column')
}

function isWorkbookRowsRefData(value: unknown): value is WorkbookRowsRefData {
  if (typeof value !== 'object' || value === null || !hasValidBaseRef(value) || value.kind !== 'rows') {
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
  return hasOptionalString(value, 'sheetName') && (table === undefined || isWorkbookTableRefData(table))
}

function isWorkbookColumnRef(value: unknown): value is WorkbookColumnRef {
  if (!isWorkbookColumnRefData(value)) {
    return false
  }
  const table = Object.getOwnPropertyDescriptor(value, 'table')?.value
  const rows = Object.getOwnPropertyDescriptor(value, 'rows')?.value
  return isWorkbookTableRef(table) && (rows === undefined || isWorkbookRowsRef(rows))
}

function isWorkbookColumnRefData(value: unknown): value is WorkbookColumnRefData {
  if (typeof value !== 'object' || value === null || !hasValidBaseRef(value) || value.kind !== 'column' || !hasOwnString(value, 'name')) {
    return false
  }
  const table = Object.getOwnPropertyDescriptor(value, 'table')?.value
  const rows = Object.getOwnPropertyDescriptor(value, 'rows')?.value
  return isWorkbookTableRefData(table) && (rows === undefined || isWorkbookRowsRefData(rows))
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

export function isWorkbookRefData(value: unknown): value is WorkbookRefData {
  if (typeof value !== 'object' || value === null || !hasValidBaseRef(value)) {
    return false
  }
  switch (value.kind) {
    case 'range':
      return isCellRangeRef(Object.getOwnPropertyDescriptor(value, 'range')?.value)
    case 'name':
      return hasOwnString(value, 'name')
    case 'table':
      return isWorkbookTableRefData(value)
    case 'column':
      return isWorkbookColumnRefData(value)
    case 'rows':
      return isWorkbookRowsRefData(value)
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
      ownArrayDataValues(current).forEach(visit)
      return
    }

    ownEnumerableDataValues(current).forEach(visit)
  }

  visit(value)
  return refs
}

export function toWorkbookRefData(ref: WorkbookRef | WorkbookRefData): WorkbookRefData {
  if (!isWorkbookRefData(ref)) {
    throw new Error('Workbook ref data is invalid')
  }
  switch (ref.kind) {
    case 'range':
      return {
        kind: 'range',
        id: ref.id,
        label: ref.label,
        range: {
          sheetName: ref.range.sheetName,
          startAddress: ref.range.startAddress,
          endAddress: ref.range.endAddress,
        },
      }
    case 'name':
      return {
        kind: 'name',
        id: ref.id,
        label: ref.label,
        name: ref.name,
      }
    case 'table':
      return {
        kind: 'table',
        id: ref.id,
        label: ref.label,
        ...(ref.name !== undefined ? { name: ref.name } : {}),
        ...(ref.sheetName !== undefined ? { sheetName: ref.sheetName } : {}),
        ...(ref.headers !== undefined ? { headers: copyStringArrayData(ref.headers) } : {}),
      }
    case 'column':
      return {
        kind: 'column',
        id: ref.id,
        label: ref.label,
        table: toWorkbookTableRefData(ref.table),
        ...(ref.rows !== undefined ? { rows: toWorkbookRowsRefData(ref.rows) } : {}),
        name: ref.name,
      }
    case 'rows':
      return {
        kind: 'rows',
        id: ref.id,
        label: ref.label,
        ...(ref.sheetName !== undefined ? { sheetName: ref.sheetName } : {}),
        ...(ref.table !== undefined ? { table: toWorkbookTableRefData(ref.table) } : {}),
        where: {
          column: ref.where.column,
          op: ref.where.op,
          value: ref.where.value,
        },
      }
  }
}

function toWorkbookTableRefData(ref: WorkbookTableRef | WorkbookTableRefData): WorkbookTableRefData {
  const data = toWorkbookRefData(ref)
  if (data.kind !== 'table') {
    throw new Error('Workbook table ref data is invalid')
  }
  return data
}

function toWorkbookRowsRefData(ref: WorkbookRowsRef | WorkbookRowsRefData): WorkbookRowsRefData {
  const data = toWorkbookRefData(ref)
  if (data.kind !== 'rows') {
    throw new Error('Workbook rows ref data is invalid')
  }
  return data
}

export function collectWorkbookRefData(value: unknown): readonly WorkbookRefData[] {
  const refs: WorkbookRefData[] = []
  const seenRefs = new Set<string>()
  const seenObjects = new WeakSet<object>()

  function pushRef(ref: WorkbookRefData): void {
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

    if (isWorkbookRefData(current)) {
      const ref = toWorkbookRefData(current)
      pushRef(ref)
      if (ref.kind === 'column') {
        visit(ref.rows)
        visit(ref.table)
      }
      if (ref.kind === 'rows') {
        visit(ref.table)
      }
      return
    }

    if (Array.isArray(current)) {
      ownArrayDataValues(current).forEach(visit)
      return
    }

    ownEnumerableDataValues(current).forEach(visit)
  }

  visit(value)
  return refs
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

export function createWorkbookTableRef(options: FindTableOptions): WorkbookTableRef {
  const { name: tableName, sheetName, headers } = normalizeTableSelector(options)
  const id = cleanIdPart(['table', sheetName, tableName, ...(headers ?? [])].filter(Boolean).join('_'))
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
  const { table, rows, name } = normalizeColumnSelector(options)
  if (!isWorkbookTableRef(table)) {
    throw new Error('Workbook column selector table must be a workbook table ref')
  }
  if (rows !== undefined && !isWorkbookRowsRef(rows)) {
    throw new Error('Workbook column selector rows must be a workbook rows ref')
  }
  const ownerId = rows?.id ?? table.id
  const ownerLabel = rows?.label ?? table.label
  const column: WorkbookColumnRef = {
    kind: 'column',
    id: cleanIdPart(`${ownerId}_${name}`),
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
    id: cleanIdPart(`range_${range.sheetName}_${range.startAddress}_${range.endAddress}`),
    label: rangeLabel(range),
    range: Object.freeze(range),
  })
}

export function createWorkbookNameRef(name: string): WorkbookNameRef {
  const normalizedName = requiredSelectorText(name, 'name')
  return Object.freeze({
    kind: 'name',
    id: cleanIdPart(`name_${normalizedName}`),
    label: normalizedName,
    name: normalizedName,
  })
}

export function createWorkbookRowsRef(options: FindRowsOptions): WorkbookRowsRef {
  const { sheetName, table, where } = normalizeRowsSelector(options)
  if (table !== undefined && !isWorkbookTableRef(table)) {
    throw new Error('Workbook rows selector table must be a workbook table ref')
  }
  const ownerId = table?.id ?? sheetName
  const ownerLabel = table?.label ?? sheetName
  const rows: WorkbookRowsRef = {
    kind: 'rows',
    id: cleanIdPart(`${ownerId}_${where.column}_${where.op}_${literalIdPart(where.value)}`),
    label: `${ownerLabel} rows where ${where.column} ${where.op} ${literalLabel(where.value)}`,
    ...(sheetName !== undefined ? { sheetName } : {}),
    ...(table !== undefined ? { table } : {}),
    where,
    column(name) {
      if (rows.table === undefined) {
        throw new Error('Rows column selection requires a table-backed row selector')
      }
      return createWorkbookColumnRef({ table: rows.table, rows, name })
    },
  }
  hideHelper(rows, 'column')
  return Object.freeze(rows)
}

function refDataKey(ref: WorkbookRefData): string {
  return `${ref.kind}:${ref.id}`
}

function sameRefData(left: WorkbookRefData | undefined, right: WorkbookRefData | undefined): boolean {
  return left !== undefined && right !== undefined && refDataKey(left) === refDataKey(right)
}

function hydrateTableRefData(data: WorkbookTableRefData): WorkbookTableRef {
  const table: WorkbookTableRef = {
    kind: 'table',
    id: data.id,
    label: data.label,
    ...(data.name !== undefined ? { name: data.name } : {}),
    ...(data.sheetName !== undefined ? { sheetName: data.sheetName } : {}),
    ...(data.headers !== undefined ? { headers: Object.freeze(copyStringArrayData(data.headers)) } : {}),
    column(name) {
      return createWorkbookColumnRef({ table, name })
    },
  }
  hideHelper(table, 'column')
  return Object.freeze(table)
}

function hydrateRowsRefData(data: WorkbookRowsRefData, sharedTable?: WorkbookTableRef): WorkbookRowsRef {
  const table = data.table === undefined ? undefined : sameRefData(data.table, sharedTable) ? sharedTable : hydrateTableRefData(data.table)
  const rows: WorkbookRowsRef = {
    kind: 'rows',
    id: data.id,
    label: data.label,
    ...(data.sheetName !== undefined ? { sheetName: data.sheetName } : {}),
    ...(table !== undefined ? { table } : {}),
    where: Object.freeze({
      column: data.where.column,
      op: data.where.op,
      value: data.where.value,
    }),
    column(name) {
      if (rows.table === undefined) {
        throw new Error('Rows column selection requires a table-backed row selector')
      }
      return createWorkbookColumnRef({ table: rows.table, rows, name })
    },
  }
  hideHelper(rows, 'column')
  return Object.freeze(rows)
}

export function hydrateWorkbookRef(data: WorkbookRefData): WorkbookRef {
  if (!isWorkbookRefData(data)) {
    throw new Error('Workbook ref data is invalid')
  }
  switch (data.kind) {
    case 'range':
      return Object.freeze({
        kind: 'range',
        id: data.id,
        label: data.label,
        range: Object.freeze({
          sheetName: data.range.sheetName,
          startAddress: data.range.startAddress,
          endAddress: data.range.endAddress,
        }),
      })
    case 'name':
      return Object.freeze({
        kind: 'name',
        id: data.id,
        label: data.label,
        name: data.name,
      })
    case 'table':
      return hydrateTableRefData(data)
    case 'column': {
      const table = hydrateTableRefData(data.table)
      const rows = data.rows === undefined ? undefined : hydrateRowsRefData(data.rows, table)
      return Object.freeze({
        kind: 'column',
        id: data.id,
        label: data.label,
        table,
        ...(rows !== undefined ? { rows } : {}),
        name: data.name,
      })
    }
    case 'rows':
      return hydrateRowsRefData(data)
  }
}

function isHydratableContainer(value: object): boolean {
  const prototype = Object.getPrototypeOf(value)
  return Array.isArray(value) || prototype === Object.prototype || prototype === null
}

function defineOwnDataProperty(target: object, key: string, value: unknown): void {
  Object.defineProperty(target, key, {
    value,
    enumerable: true,
    configurable: true,
    writable: true,
  })
}

function hydrateWorkbookRefsValue(value: unknown, seen: WeakMap<object, unknown>): unknown {
  if (value === null || value === undefined || typeof value !== 'object') {
    return value
  }
  if (isWorkbookRefData(value)) {
    return hydrateWorkbookRef(value)
  }
  if (!isHydratableContainer(value)) {
    return value
  }
  const cached = seen.get(value)
  if (cached !== undefined) {
    return cached
  }
  if (Array.isArray(value)) {
    const hydrated: unknown[] = []
    seen.set(value, hydrated)
    ownArrayDataValues(value).forEach((entry) => {
      hydrated.push(hydrateWorkbookRefsValue(entry, seen))
    })
    return Object.freeze(hydrated)
  }
  const hydrated: Record<string, unknown> = {}
  seen.set(value, hydrated)
  ownEnumerableDataEntries(value).forEach(([key, entry]) => {
    defineOwnDataProperty(hydrated, key, hydrateWorkbookRefsValue(entry, seen))
  })
  return Object.freeze(hydrated)
}

export function hydrateWorkbookRefs(value: unknown): unknown {
  return hydrateWorkbookRefsValue(value, new WeakMap())
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
