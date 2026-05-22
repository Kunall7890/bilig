import {
  ErrorCode,
  ValueTag,
  type CellValue,
  type LiteralInput,
  type PivotAggregation,
  type WorkbookPivotSnapshot,
  type WorkbookPivotValueSnapshot,
} from '@bilig/protocol'

export type PivotDefinitionInput = Pick<
  WorkbookPivotSnapshot,
  'groupBy' | 'values' | 'columnFields' | 'filters' | 'pageFields' | 'hiddenItems'
>

export type PivotMaterializationResult =
  | {
      kind: 'ok'
      rows: number
      cols: number
      values: CellValue[]
    }
  | {
      kind: 'error'
      code: ErrorCode.Value
      rows: 1
      cols: 1
      values: [CellValue]
    }

interface MaterializedPivotField extends WorkbookPivotValueSnapshot {
  columnIndex: number
  headerLabel: string
}

interface MaterializedSourceField {
  columnIndex: number
  headerLabel: string
}

interface MaterializedPivotFilter {
  field: MaterializedSourceField
  includedValues?: readonly LiteralInput[]
  hiddenValues?: readonly LiteralInput[]
}

export interface PivotAggregateState {
  sum: number
  count: number
  numericCount: number
  min: number
  max: number
  product: number
}

interface ColumnBucket {
  key: string
  keyValues: CellValue[]
}

interface GroupBucket {
  keyValues: CellValue[]
  aggregatesByColumnKey: Map<string, PivotAggregateState[]>
}

export function materializePivotTable(
  definition: PivotDefinitionInput,
  sourceRows: readonly (readonly CellValue[])[],
): PivotMaterializationResult {
  if (definition.values.length === 0) {
    return pivotConfigError()
  }

  const headerRow = sourceRows[0]
  if (!headerRow || headerRow.length === 0) {
    return pivotConfigError()
  }

  const headerLookup = new Map<string, { index: number; label: string }>()
  for (let columnIndex = 0; columnIndex < headerRow.length; columnIndex += 1) {
    const label = headerLabel(headerRow[columnIndex] ?? emptyValue())
    const normalized = normalizeHeader(label)
    if (normalized.length === 0 || headerLookup.has(normalized)) {
      continue
    }
    headerLookup.set(normalized, { index: columnIndex, label })
  }

  const groupFields = resolveSourceFields(definition.groupBy, headerLookup)
  if (groupFields.some((field) => field === undefined)) {
    return pivotConfigError()
  }
  const materializedGroupFields = groupFields.filter((field): field is MaterializedSourceField => field !== undefined)

  const columnFields = resolveSourceFields(definition.columnFields ?? [], headerLookup)
  if (columnFields.some((field) => field === undefined)) {
    return pivotConfigError()
  }
  const materializedColumnFields = columnFields.filter((field): field is MaterializedSourceField => field !== undefined)

  const valueFields = definition.values.map((field) => resolveValueField(field, headerLookup))
  if (valueFields.some((field) => field === undefined)) {
    return pivotConfigError()
  }
  const materializedValueFields = valueFields.filter((field): field is MaterializedPivotField => field !== undefined)
  const filters = resolvePivotFilters(definition, headerLookup)
  if (filters === null) {
    return pivotConfigError()
  }

  const buckets = new Map<string, GroupBucket>()
  const columnBuckets: ColumnBucket[] = materializedColumnFields.length === 0 ? [{ key: '', keyValues: [] }] : []
  const seenColumnKeys = new Set(columnBuckets.map((bucket) => bucket.key))
  for (let rowIndex = 1; rowIndex < sourceRows.length; rowIndex += 1) {
    const row = sourceRows[rowIndex] ?? []
    if (!rowPassesFilters(row, filters)) {
      continue
    }
    const keyValues = materializedGroupFields.map((field) => cloneOutputCell(row[field.columnIndex] ?? emptyValue()))
    const columnKeyValues = materializedColumnFields.map((field) => cloneOutputCell(row[field.columnIndex] ?? emptyValue()))
    const hasObservedValue = hasMeaningfulRowValue(keyValues, columnKeyValues, materializedValueFields, row)
    if (!hasObservedValue) {
      continue
    }
    const key = keyValues.map(cellValueKey).join('\u001f')
    const columnKey = columnKeyValues.map(cellValueKey).join('\u001f')
    if (!seenColumnKeys.has(columnKey)) {
      seenColumnKeys.add(columnKey)
      columnBuckets.push({ key: columnKey, keyValues: columnKeyValues })
    }
    let bucket = buckets.get(key)
    if (!bucket) {
      bucket = {
        keyValues,
        aggregatesByColumnKey: new Map(),
      }
      buckets.set(key, bucket)
    }
    let aggregates = bucket.aggregatesByColumnKey.get(columnKey)
    if (!aggregates) {
      aggregates = Array.from({ length: materializedValueFields.length }, () => emptyPivotAggregateState())
      bucket.aggregatesByColumnKey.set(columnKey, aggregates)
    }
    for (let valueIndex = 0; valueIndex < materializedValueFields.length; valueIndex += 1) {
      const field = materializedValueFields[valueIndex]!
      const cell = row[field.columnIndex] ?? emptyValue()
      accumulatePivotAggregateValue(aggregates[valueIndex]!, cell)
    }
  }

  const cols = definition.groupBy.length + columnBuckets.length * materializedValueFields.length
  const values: CellValue[] = []
  for (let groupIndex = 0; groupIndex < materializedGroupFields.length; groupIndex += 1) {
    values.push(stringValue(materializedGroupFields[groupIndex]!.headerLabel))
  }
  for (const columnBucket of columnBuckets) {
    for (let valueIndex = 0; valueIndex < materializedValueFields.length; valueIndex += 1) {
      const label = outputLabel(materializedValueFields[valueIndex]!)
      values.push(stringValue(materializedColumnFields.length === 0 ? label : `${columnBucketLabel(columnBucket)} ${label}`))
    }
  }
  buckets.forEach((bucket) => {
    bucket.keyValues.forEach((keyValue) => {
      values.push(cloneOutputCell(keyValue))
    })
    for (const columnBucket of columnBuckets) {
      const aggregates = bucket.aggregatesByColumnKey.get(columnBucket.key)
      for (let valueIndex = 0; valueIndex < materializedValueFields.length; valueIndex += 1) {
        values.push(numberValue(finalizePivotAggregate(materializedValueFields[valueIndex]!.summarizeBy, aggregates?.[valueIndex])))
      }
    }
  })

  return {
    kind: 'ok',
    rows: buckets.size + 1,
    cols,
    values,
  }
}

function resolveSourceFields(
  fieldNames: readonly string[],
  headerLookup: Map<string, { index: number; label: string }>,
): (MaterializedSourceField | undefined)[] {
  return fieldNames.map((fieldName) => {
    const resolved = headerLookup.get(normalizeHeader(fieldName))
    return resolved ? { columnIndex: resolved.index, headerLabel: resolved.label } : undefined
  })
}

function resolveValueField(
  field: WorkbookPivotValueSnapshot,
  headerLookup: Map<string, { index: number; label: string }>,
): MaterializedPivotField | undefined {
  const resolved = headerLookup.get(normalizeHeader(field.sourceColumn))
  if (!resolved) {
    return undefined
  }
  return {
    ...field,
    columnIndex: resolved.index,
    headerLabel: resolved.label,
  }
}

function resolvePivotFilters(
  definition: PivotDefinitionInput,
  headerLookup: Map<string, { index: number; label: string }>,
): MaterializedPivotFilter[] | null {
  const filters: MaterializedPivotFilter[] = []
  for (const filter of definition.filters ?? []) {
    const field = resolveSourceFields([filter.sourceColumn], headerLookup)[0]
    if (!field) {
      return null
    }
    filters.push({
      field,
      ...(filter.includedValues ? { includedValues: filter.includedValues } : {}),
      ...(filter.hiddenValues ? { hiddenValues: filter.hiddenValues } : {}),
    })
  }
  for (const pageField of definition.pageFields ?? []) {
    const field = resolveSourceFields([pageField.sourceColumn], headerLookup)[0]
    if (!field) {
      return null
    }
    filters.push({
      field,
      ...(pageField.selectedValue !== undefined ? { includedValues: [pageField.selectedValue] } : {}),
    })
  }
  for (const hiddenItems of definition.hiddenItems ?? []) {
    const field = resolveSourceFields([hiddenItems.sourceColumn], headerLookup)[0]
    if (!field) {
      return null
    }
    filters.push({ field, hiddenValues: hiddenItems.values })
  }
  return filters
}

function rowPassesFilters(row: readonly CellValue[], filters: readonly MaterializedPivotFilter[]): boolean {
  for (const filter of filters) {
    const value = cellValueToLiteral(row[filter.field.columnIndex] ?? emptyValue())
    if (filter.includedValues && !filter.includedValues.some((candidate) => literalValuesEqual(candidate, value))) {
      return false
    }
    if (filter.hiddenValues?.some((candidate) => literalValuesEqual(candidate, value))) {
      return false
    }
  }
  return true
}

export function emptyPivotAggregateState(): PivotAggregateState {
  return {
    sum: 0,
    count: 0,
    numericCount: 0,
    min: Number.POSITIVE_INFINITY,
    max: Number.NEGATIVE_INFINITY,
    product: 1,
  }
}

export function accumulatePivotAggregateValue(state: PivotAggregateState, value: CellValue): void {
  if (!isEmptyValue(value)) {
    state.count += 1
  }
  if (value.tag !== ValueTag.Number) {
    return
  }
  state.sum += value.value
  state.numericCount += 1
  state.min = Math.min(state.min, value.value)
  state.max = Math.max(state.max, value.value)
  state.product *= value.value
}

export function finalizePivotAggregate(mode: PivotAggregation, state: PivotAggregateState | undefined): number {
  if (!state) {
    return 0
  }
  switch (mode) {
    case 'sum':
      return state.sum
    case 'count':
      return state.count
    case 'countNums':
      return state.numericCount
    case 'average':
      return state.numericCount === 0 ? 0 : state.sum / state.numericCount
    case 'min':
      return state.numericCount === 0 ? 0 : state.min
    case 'max':
      return state.numericCount === 0 ? 0 : state.max
    case 'product':
      return state.numericCount === 0 ? 0 : state.product
  }
}

function outputLabel(field: MaterializedPivotField): string {
  const customLabel = field.outputLabel?.trim()
  if (customLabel && customLabel.length > 0) {
    return customLabel
  }
  return `${field.summarizeBy.toUpperCase()} of ${field.headerLabel}`
}

function hasMeaningfulRowValue(
  keyValues: readonly CellValue[],
  columnKeyValues: readonly CellValue[],
  valueFields: readonly MaterializedPivotField[],
  row: readonly CellValue[],
): boolean {
  if (keyValues.some((value) => !isEmptyValue(value)) || columnKeyValues.some((value) => !isEmptyValue(value))) {
    return true
  }
  return valueFields.some((field) => !isEmptyValue(row[field.columnIndex] ?? emptyValue()))
}

function columnBucketLabel(bucket: ColumnBucket): string {
  const label = bucket.keyValues
    .map(displayCellValue)
    .filter((part) => part.length > 0)
    .join(' / ')
  return label.length > 0 ? label : '(blank)'
}

function displayCellValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return String(value.code)
  }
}

function cellValueToLiteral(value: CellValue): LiteralInput {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return value.code
  }
}

function literalValuesEqual(left: LiteralInput, right: LiteralInput): boolean {
  if (typeof left === 'number' && typeof right === 'number') {
    return Object.is(left, right) || left === right
  }
  return left === right
}

function cellValueKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'E'
    case ValueTag.Number:
      return `N:${Object.is(value.value, -0) ? '-0' : String(value.value)}`
    case ValueTag.Boolean:
      return value.value ? 'B:1' : 'B:0'
    case ValueTag.String:
      return `S:${value.value}`
    case ValueTag.Error:
      return `R:${value.code}`
  }
}

function headerLabel(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value.trim()
    case ValueTag.Error:
      return ''
  }
}

function normalizeHeader(value: string): string {
  return value.trim().toUpperCase()
}

function isEmptyValue(value: CellValue): boolean {
  return value.tag === ValueTag.Empty
}

function cloneOutputCell(value: CellValue): CellValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return emptyValue()
    case ValueTag.Number:
      return numberValue(value.value)
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: value.value }
    case ValueTag.String:
      return stringValue(value.value)
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: value.code }
  }
}

function pivotConfigError(): PivotMaterializationResult {
  return {
    kind: 'error',
    code: ErrorCode.Value,
    rows: 1,
    cols: 1,
    values: [{ tag: ValueTag.Error, code: ErrorCode.Value }],
  }
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty }
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}
