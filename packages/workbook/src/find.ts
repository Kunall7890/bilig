import type { CellRangeRef, LiteralInput } from '@bilig/protocol'

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
  readonly name: string
}

export type WorkbookRowOperator = 'eq' | 'neq' | 'contains' | 'startsWith' | 'gt' | 'gte' | 'lt' | 'lte'

export interface WorkbookRowsRef extends WorkbookBaseRef {
  readonly kind: 'rows'
  readonly sheetName?: string
  readonly table?: WorkbookTableRef
  readonly where: {
    readonly column: string
    readonly op: WorkbookRowOperator
    readonly value: LiteralInput
  }
}

export type WorkbookRef = WorkbookRangeRef | WorkbookNameRef | WorkbookTableRef | WorkbookColumnRef | WorkbookRowsRef

const WORKBOOK_REF_KINDS = new Set<string>(['range', 'name', 'table', 'column', 'rows'])

export function isWorkbookRef(value: unknown): value is WorkbookRef {
  return (
    typeof value === 'object' &&
    value !== null &&
    'kind' in value &&
    typeof value.kind === 'string' &&
    WORKBOOK_REF_KINDS.has(value.kind) &&
    'id' in value &&
    typeof value.id === 'string' &&
    'label' in value &&
    typeof value.label === 'string'
  )
}

export function collectWorkbookRefs(value: unknown): readonly WorkbookRef[] {
  const refs: WorkbookRef[] = []
  const seenRefs = new Set<string>()
  const seenObjects = new WeakSet<object>()

  function visit(current: unknown): void {
    if (current === null || current === undefined || typeof current !== 'object') {
      return
    }
    if (seenObjects.has(current)) {
      return
    }
    seenObjects.add(current)

    if (isWorkbookRef(current)) {
      const key = `${current.kind}:${current.id}`
      if (!seenRefs.has(key)) {
        seenRefs.add(key)
        refs.push(current)
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

export type FindRangeInput =
  | CellRangeRef
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

function cleanIdPart(value: string): string {
  const cleaned = value
    .trim()
    .replaceAll(/[^A-Za-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '')
  return cleaned === '' ? 'ref' : cleaned
}

function rangeLabel(range: CellRangeRef): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

export function normalizeRangeRef(input: FindRangeInput): CellRangeRef {
  if ('startAddress' in input) {
    return {
      sheetName: input.sheetName,
      startAddress: input.startAddress,
      endAddress: input.endAddress ?? input.startAddress,
    }
  }
  if ('address' in input) {
    return {
      sheetName: input.sheetName,
      startAddress: input.address,
      endAddress: input.address,
    }
  }
  return input
}

export function createWorkbookTableRef(options: FindTableOptions): WorkbookTableRef {
  const id = cleanIdPart(['table', options.sheetName, options.name, ...(options.headers ?? [])].filter(Boolean).join('_'))
  const label = options.name ?? (options.headers ? `table with ${options.headers.join(', ')}` : 'table')
  const table: WorkbookTableRef = {
    kind: 'table',
    id,
    label,
    ...(options.name !== undefined ? { name: options.name } : {}),
    ...(options.sheetName !== undefined ? { sheetName: options.sheetName } : {}),
    ...(options.headers !== undefined ? { headers: [...options.headers] } : {}),
    column(name) {
      return createWorkbookColumnRef({ table, name })
    },
  }
  return table
}

export function createWorkbookColumnRef(options: FindColumnOptions): WorkbookColumnRef {
  return {
    kind: 'column',
    id: cleanIdPart(`${options.table.id}_${options.name}`),
    label: `${options.table.label}.${options.name}`,
    table: options.table,
    name: options.name,
  }
}

export function createWorkbookRangeRef(input: FindRangeInput): WorkbookRangeRef {
  const range = normalizeRangeRef(input)
  return {
    kind: 'range',
    id: cleanIdPart(`range_${range.sheetName}_${range.startAddress}_${range.endAddress}`),
    label: rangeLabel(range),
    range,
  }
}

export function createWorkbookNameRef(name: string): WorkbookNameRef {
  return {
    kind: 'name',
    id: cleanIdPart(`name_${name}`),
    label: name,
    name,
  }
}

export function createWorkbookRowsRef(options: FindRowsOptions): WorkbookRowsRef {
  const owner = options.table?.id ?? options.sheetName ?? 'rows'
  return {
    kind: 'rows',
    id: cleanIdPart(`${owner}_${options.where.column}_${options.where.op}`),
    label: `${owner} rows where ${options.where.column} ${options.where.op}`,
    ...(options.sheetName !== undefined ? { sheetName: options.sheetName } : {}),
    ...(options.table !== undefined ? { table: options.table } : {}),
    where: options.where,
  }
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
