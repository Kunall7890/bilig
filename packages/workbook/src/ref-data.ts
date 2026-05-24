import { isLiteralInput } from '@bilig/protocol'
import { isObjectRecord } from './data-properties.js'
import type { WorkbookRefData, WorkbookRefKind } from './find.js'
import { isWorkbookRowOperator, isWorkbookRowValueCompatible } from './selectors.js'

export type WorkbookRefDataIssueCode = 'invalid_type' | 'missing_field' | 'invalid_field'

export interface WorkbookRefDataIssue {
  readonly code: WorkbookRefDataIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookRefDataCheckResult =
  | {
      readonly status: 'valid'
      readonly ref: WorkbookRefData
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookRefDataIssue[]
    }

const workbookRefDataKinds = Object.freeze(['range', 'name', 'table', 'column', 'rows'] as const satisfies readonly WorkbookRefKind[])
const workbookRefDataKindSet = new Set<string>(workbookRefDataKinds)

function refDataIssue(code: WorkbookRefDataIssueCode, path: string, message: string): WorkbookRefDataIssue {
  return Object.freeze({
    code,
    path,
    message,
  })
}

function invalidType(path: string, label: string): WorkbookRefDataIssue {
  return refDataIssue('invalid_type', path, `${label} must be an object`)
}

function invalidObjectRecord(path: string, label: string): WorkbookRefDataIssue {
  return refDataIssue('invalid_type', path, `${label} must be an object record`)
}

function isNonArrayObject(value: unknown): value is object {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isRefDataObject(value: unknown): value is Record<string, unknown> {
  return isObjectRecord(value)
}

function isWorkbookRefDataKind(value: unknown): value is WorkbookRefKind {
  return typeof value === 'string' && workbookRefDataKindSet.has(value)
}

function ownDataDescriptor(value: Record<string, unknown>, key: string): PropertyDescriptor | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined || !descriptor.enumerable) {
    return undefined
  }
  return descriptor
}

function ownDataValue(value: Record<string, unknown>, key: string): unknown {
  const descriptor = ownDataDescriptor(value, key)
  return descriptor !== undefined && 'value' in descriptor ? descriptor.value : undefined
}

function pushRequiredStringIssue(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string, key: string): void {
  const fieldPath = `${path}.${key}`
  const descriptor = ownDataDescriptor(value, key)
  if (descriptor === undefined) {
    issues.push(refDataIssue('missing_field', fieldPath, `Workbook ref data ${fieldPath} is required`))
    return
  }
  if (!('value' in descriptor)) {
    issues.push(refDataIssue('invalid_field', fieldPath, `Workbook ref data ${fieldPath} must be a data property`))
    return
  }
  if (typeof descriptor.value !== 'string') {
    issues.push(refDataIssue('invalid_field', fieldPath, `Workbook ref data ${fieldPath} must be a string`))
  }
}

function pushOptionalStringIssue(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string, key: string): void {
  const fieldPath = `${path}.${key}`
  const descriptor = ownDataDescriptor(value, key)
  if (descriptor === undefined) {
    return
  }
  if (!('value' in descriptor)) {
    issues.push(refDataIssue('invalid_field', fieldPath, `Workbook ref data ${fieldPath} must be a data property`))
    return
  }
  if (typeof descriptor.value !== 'string') {
    issues.push(refDataIssue('invalid_field', fieldPath, `Workbook ref data ${fieldPath} must be a string`))
  }
}

function pushOptionalStringArrayIssue(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string, key: string): void {
  const fieldPath = `${path}.${key}`
  const descriptor = ownDataDescriptor(value, key)
  if (descriptor === undefined) {
    return
  }
  if (!('value' in descriptor)) {
    issues.push(refDataIssue('invalid_field', fieldPath, `Workbook ref data ${fieldPath} must be a data property`))
    return
  }
  if (!Array.isArray(descriptor.value)) {
    issues.push(refDataIssue('invalid_field', fieldPath, `Workbook ref data ${fieldPath} must be an array`))
    return
  }
  const arrayDescriptors = Object.getOwnPropertyDescriptors(descriptor.value)
  for (let index = 0; index < descriptor.value.length; index += 1) {
    const entryPath = `${fieldPath}[${index}]`
    const entry = arrayDescriptors[String(index)]
    if (entry === undefined || !entry.enumerable || !('value' in entry)) {
      issues.push(refDataIssue('invalid_field', entryPath, `Workbook ref data ${entryPath} must be a data property`))
      continue
    }
    if (typeof entry.value !== 'string') {
      issues.push(refDataIssue('invalid_field', entryPath, `Workbook ref data ${entryPath} must be a string`))
    }
  }
}

function pushBaseRefDataIssues(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string): WorkbookRefKind | undefined {
  pushRequiredStringIssue(issues, value, path, 'kind')
  pushRequiredStringIssue(issues, value, path, 'id')
  pushRequiredStringIssue(issues, value, path, 'label')
  const kind = ownDataValue(value, 'kind')
  if (!isWorkbookRefDataKind(kind)) {
    if (typeof kind === 'string') {
      issues.push(refDataIssue('invalid_field', `${path}.kind`, `Workbook ref data ${path}.kind must be a workbook ref kind`))
    }
    return undefined
  }
  return kind
}

function pushRangeIssues(issues: WorkbookRefDataIssue[], value: unknown, path: string): void {
  if (!isNonArrayObject(value)) {
    issues.push(invalidType(path, `Workbook ref data ${path}`))
    return
  }
  if (!isRefDataObject(value)) {
    issues.push(invalidObjectRecord(path, `Workbook ref data ${path}`))
    return
  }
  pushRequiredStringIssue(issues, value, path, 'sheetName')
  pushRequiredStringIssue(issues, value, path, 'startAddress')
  pushRequiredStringIssue(issues, value, path, 'endAddress')
}

function pushNameRefDataIssues(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string): void {
  pushRequiredStringIssue(issues, value, path, 'name')
}

function pushTableRefDataIssues(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string): void {
  pushOptionalStringIssue(issues, value, path, 'name')
  pushOptionalStringIssue(issues, value, path, 'sheetName')
  pushOptionalStringArrayIssue(issues, value, path, 'headers')
}

function pushRowsWhereIssues(issues: WorkbookRefDataIssue[], value: unknown, path: string): void {
  if (!isNonArrayObject(value)) {
    issues.push(invalidType(path, `Workbook ref data ${path}`))
    return
  }
  if (!isRefDataObject(value)) {
    issues.push(invalidObjectRecord(path, `Workbook ref data ${path}`))
    return
  }
  pushRequiredStringIssue(issues, value, path, 'column')
  const op = ownDataValue(value, 'op')
  if (!isWorkbookRowOperator(op)) {
    const descriptor = ownDataDescriptor(value, 'op')
    issues.push(
      refDataIssue(
        descriptor === undefined ? 'missing_field' : 'invalid_field',
        `${path}.op`,
        `Workbook ref data ${path}.op must be a workbook row operator`,
      ),
    )
    return
  }
  const rowValue = ownDataValue(value, 'value')
  if (!isLiteralInput(rowValue)) {
    const descriptor = ownDataDescriptor(value, 'value')
    issues.push(
      refDataIssue(
        descriptor === undefined ? 'missing_field' : 'invalid_field',
        `${path}.value`,
        `Workbook ref data ${path}.value must be a finite JSON literal`,
      ),
    )
    return
  }
  if (!isWorkbookRowValueCompatible(op, rowValue)) {
    issues.push(refDataIssue('invalid_field', `${path}.value`, `Workbook ref data ${path}.value is not compatible with row operator ${op}`))
  }
}

function pushRowsRefDataIssues(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string): void {
  pushOptionalStringIssue(issues, value, path, 'sheetName')
  const table = ownDataDescriptor(value, 'table')
  if (table !== undefined) {
    if (!('value' in table)) {
      issues.push(refDataIssue('invalid_field', `${path}.table`, `Workbook ref data ${path}.table must be a data property`))
    } else {
      pushRefDataIssues(issues, table.value, `${path}.table`)
      if (isRefDataObject(table.value) && ownDataValue(table.value, 'kind') !== 'table') {
        issues.push(refDataIssue('invalid_field', `${path}.table.kind`, `Workbook ref data ${path}.table must be a table ref`))
      }
    }
  }
  pushRowsWhereIssues(issues, ownDataValue(value, 'where'), `${path}.where`)
}

function pushColumnRefDataIssues(issues: WorkbookRefDataIssue[], value: Record<string, unknown>, path: string): void {
  pushRequiredStringIssue(issues, value, path, 'name')
  const table = ownDataValue(value, 'table')
  pushRefDataIssues(issues, table, `${path}.table`)
  if (isRefDataObject(table) && ownDataValue(table, 'kind') !== 'table') {
    issues.push(refDataIssue('invalid_field', `${path}.table.kind`, `Workbook ref data ${path}.table must be a table ref`))
  }
  const rows = ownDataDescriptor(value, 'rows')
  if (rows !== undefined) {
    if (!('value' in rows)) {
      issues.push(refDataIssue('invalid_field', `${path}.rows`, `Workbook ref data ${path}.rows must be a data property`))
    } else {
      pushRefDataIssues(issues, rows.value, `${path}.rows`)
      if (isRefDataObject(rows.value) && ownDataValue(rows.value, 'kind') !== 'rows') {
        issues.push(refDataIssue('invalid_field', `${path}.rows.kind`, `Workbook ref data ${path}.rows must be a rows ref`))
      }
    }
  }
}

function pushRefDataIssues(issues: WorkbookRefDataIssue[], value: unknown, path: string): void {
  if (!isNonArrayObject(value)) {
    issues.push(invalidType(path, `Workbook ref data ${path}`))
    return
  }
  if (!isRefDataObject(value)) {
    issues.push(invalidObjectRecord(path, `Workbook ref data ${path}`))
    return
  }
  const kind = pushBaseRefDataIssues(issues, value, path)
  switch (kind) {
    case 'range':
      pushRangeIssues(issues, ownDataValue(value, 'range'), `${path}.range`)
      return
    case 'name':
      pushNameRefDataIssues(issues, value, path)
      return
    case 'table':
      pushTableRefDataIssues(issues, value, path)
      return
    case 'column':
      pushColumnRefDataIssues(issues, value, path)
      return
    case 'rows':
      pushRowsRefDataIssues(issues, value, path)
      return
    case undefined:
      return
  }
}

function freezeRefData<T>(value: T, seen = new WeakSet<object>()): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeRefData(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}

function requiredStringValue(value: Record<string, unknown>, key: string): string {
  const raw = ownDataValue(value, key)
  if (typeof raw !== 'string') {
    throw new Error(`Workbook ref data ${key} is invalid`)
  }
  return raw
}

function optionalStringValue(value: Record<string, unknown>, key: string): string | undefined {
  const raw = ownDataValue(value, key)
  if (raw === undefined) {
    return undefined
  }
  if (typeof raw !== 'string') {
    throw new Error(`Workbook ref data ${key} is invalid`)
  }
  return raw
}

function optionalStringArrayValue(value: Record<string, unknown>, key: string): readonly string[] | undefined {
  const raw = ownDataValue(value, key)
  if (raw === undefined) {
    return undefined
  }
  if (!Array.isArray(raw)) {
    throw new Error(`Workbook ref data ${key} is invalid`)
  }
  const entries: string[] = []
  const descriptors = Object.getOwnPropertyDescriptors(raw)
  for (let index = 0; index < raw.length; index += 1) {
    const entry = descriptors[String(index)]
    if (entry === undefined || !entry.enumerable || !('value' in entry) || typeof entry.value !== 'string') {
      throw new Error(`Workbook ref data ${key} is invalid`)
    }
    entries.push(entry.value)
  }
  return Object.freeze(entries)
}

function normalizeRangeData(value: unknown): Extract<WorkbookRefData, { kind: 'range' }>['range'] {
  if (!isRefDataObject(value)) {
    throw new Error('Workbook ref data range is invalid')
  }
  return Object.freeze({
    sheetName: requiredStringValue(value, 'sheetName'),
    startAddress: requiredStringValue(value, 'startAddress'),
    endAddress: requiredStringValue(value, 'endAddress'),
  })
}

function normalizeRowsWhere(value: unknown): Extract<WorkbookRefData, { kind: 'rows' }>['where'] {
  if (!isRefDataObject(value)) {
    throw new Error('Workbook ref data where is invalid')
  }
  const op = ownDataValue(value, 'op')
  const rowValue = ownDataValue(value, 'value')
  if (!isWorkbookRowOperator(op) || !isLiteralInput(rowValue) || !isWorkbookRowValueCompatible(op, rowValue)) {
    throw new Error('Workbook ref data where is invalid')
  }
  return Object.freeze({
    column: requiredStringValue(value, 'column'),
    op,
    value: rowValue,
  })
}

function normalizeRefData(value: unknown): WorkbookRefData {
  if (!isRefDataObject(value)) {
    throw new Error('Workbook ref data is invalid')
  }
  const kind = ownDataValue(value, 'kind')
  if (!isWorkbookRefDataKind(kind)) {
    throw new Error('Workbook ref data kind is invalid')
  }
  const id = requiredStringValue(value, 'id')
  const label = requiredStringValue(value, 'label')
  switch (kind) {
    case 'range':
      return Object.freeze({
        kind,
        id,
        label,
        range: normalizeRangeData(ownDataValue(value, 'range')),
      })
    case 'name':
      return Object.freeze({
        kind,
        id,
        label,
        name: requiredStringValue(value, 'name'),
      })
    case 'table': {
      const name = optionalStringValue(value, 'name')
      const sheetName = optionalStringValue(value, 'sheetName')
      const headers = optionalStringArrayValue(value, 'headers')
      return Object.freeze({
        kind,
        id,
        label,
        ...(name !== undefined ? { name } : {}),
        ...(sheetName !== undefined ? { sheetName } : {}),
        ...(headers !== undefined ? { headers } : {}),
      })
    }
    case 'column': {
      const table = normalizeRefData(ownDataValue(value, 'table'))
      if (table.kind !== 'table') {
        throw new Error('Workbook ref data table is invalid')
      }
      const rowsValue = ownDataValue(value, 'rows')
      const rows = rowsValue === undefined ? undefined : normalizeRefData(rowsValue)
      if (rows !== undefined && rows.kind !== 'rows') {
        throw new Error('Workbook ref data rows is invalid')
      }
      return Object.freeze({
        kind,
        id,
        label,
        table,
        ...(rows !== undefined ? { rows } : {}),
        name: requiredStringValue(value, 'name'),
      })
    }
    case 'rows': {
      const tableValue = ownDataValue(value, 'table')
      const table = tableValue === undefined ? undefined : normalizeRefData(tableValue)
      if (table !== undefined && table.kind !== 'table') {
        throw new Error('Workbook ref data table is invalid')
      }
      const sheetName = optionalStringValue(value, 'sheetName')
      return Object.freeze({
        kind,
        id,
        label,
        ...(sheetName !== undefined ? { sheetName } : {}),
        ...(table !== undefined ? { table } : {}),
        where: normalizeRowsWhere(ownDataValue(value, 'where')),
      })
    }
  }
}

export function checkWorkbookRefData(value: unknown): WorkbookRefDataCheckResult {
  const issues: WorkbookRefDataIssue[] = []
  pushRefDataIssues(issues, value, 'ref')
  if (issues.length > 0) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze(issues),
    })
  }
  return Object.freeze({
    status: 'valid',
    ref: freezeRefData(normalizeRefData(value)),
    issues: Object.freeze([] as const),
  })
}
