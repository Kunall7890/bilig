import * as XLSX from 'xlsx'

import type { LiteralInput, WorkbookPivotSnapshot, WorkbookPivotValueSnapshot } from '@bilig/protocol'
import { escapeXml, spreadsheetNamespace } from './xlsx-pivot-artifacts.js'
import { defaultDataFieldVerb, subtotalValue } from './xlsx-pivot-aggregate-xml.js'

interface PivotCacheTable {
  readonly fields: readonly {
    readonly name: string
    readonly values: readonly LiteralInput[]
  }[]
  readonly rows: readonly (readonly LiteralInput[])[]
}

interface PivotLayoutBucket {
  readonly values: readonly LiteralInput[]
}

function pivotOutputRange(pivot: WorkbookPivotSnapshot): string {
  const start = XLSX.utils.decode_cell(pivot.address)
  const end = {
    r: start.r + Math.max(1, pivot.rows) - 1,
    c: start.c + Math.max(1, pivot.cols) - 1,
  }
  return XLSX.utils.encode_range({ s: start, e: end })
}

function uniqueValues(values: readonly LiteralInput[]): LiteralInput[] {
  const seen = new Set<string>()
  const output: LiteralInput[] = []
  for (const value of values) {
    const key = literalValueKey(value)
    if (!seen.has(key)) {
      seen.add(key)
      output.push(value)
    }
  }
  return output
}

function literalValueKey(value: LiteralInput): string {
  return `${typeof value}:${String(value)}`
}

function literalValuesEqual(left: LiteralInput, right: LiteralInput): boolean {
  return literalValueKey(left) === literalValueKey(right)
}

function pivotHiddenValuesForField(pivot: WorkbookPivotSnapshot, fieldName: string): readonly LiteralInput[] {
  return pivot.hiddenItems?.find((entry) => entry.sourceColumn === fieldName)?.values ?? []
}

function pivotFieldItemsXml(cacheTable: PivotCacheTable, index: number, hiddenValues: readonly LiteralInput[]): string {
  const values = uniqueValues(cacheTable.fields[index]?.values ?? [])
  const items = values.map((value, itemIndex) => {
    const hidden = hiddenValues.some((candidate) => literalValuesEqual(candidate, value)) ? ' h="1"' : ''
    return `<item x="${String(itemIndex)}"${hidden}/>`
  })
  items.push('<item t="default"/>')
  return `<items count="${String(items.length)}">${items.join('')}</items>`
}

function pivotFieldXml(index: number, fields: readonly string[], pivot: WorkbookPivotSnapshot, cacheTable: PivotCacheTable): string {
  const fieldName = fields[index] ?? ''
  if (pivot.groupBy.includes(fieldName)) {
    return `<pivotField axis="axisRow" showAll="0">${pivotFieldItemsXml(cacheTable, index, pivotHiddenValuesForField(pivot, fieldName))}</pivotField>`
  }
  if ((pivot.columnFields ?? []).includes(fieldName)) {
    return `<pivotField axis="axisCol" showAll="0">${pivotFieldItemsXml(cacheTable, index, pivotHiddenValuesForField(pivot, fieldName))}</pivotField>`
  }
  if ((pivot.pageFields ?? []).some((field) => field.sourceColumn === fieldName)) {
    return `<pivotField axis="axisPage" showAll="0">${pivotFieldItemsXml(cacheTable, index, pivotHiddenValuesForField(pivot, fieldName))}</pivotField>`
  }
  if (pivot.values.some((value) => value.sourceColumn === fieldName)) {
    return '<pivotField dataField="1" showAll="0"/>'
  }
  return '<pivotField showAll="0"/>'
}

function buildRowFieldsXml(fieldIndexes: readonly number[]): string {
  if (fieldIndexes.length === 0) {
    return ''
  }
  return [
    `<rowFields count="${String(fieldIndexes.length)}">`,
    ...fieldIndexes.map((index) => `<field x="${String(index)}"/>`),
    '</rowFields>',
  ].join('')
}

function buildColumnFieldsXml(fieldIndexes: readonly number[]): string {
  if (fieldIndexes.length === 0) {
    return ''
  }
  return [
    `<colFields count="${String(fieldIndexes.length)}">`,
    ...fieldIndexes.map((index) => `<field x="${String(index)}"/>`),
    '</colFields>',
  ].join('')
}

function buildPageFieldsXml(
  pageFields: readonly NonNullable<WorkbookPivotSnapshot['pageFields']>[number][],
  fieldIndexesByName: ReadonlyMap<string, number>,
  cacheTable: PivotCacheTable,
): string {
  const fields = pageFields.flatMap((pageField) => {
    const fieldIndex = fieldIndexesByName.get(pageField.sourceColumn)
    if (fieldIndex === undefined) {
      return []
    }
    const item =
      pageField.selectedValue === undefined ? '' : ` item="${String(pivotFieldItemIndex(cacheTable, fieldIndex, pageField.selectedValue))}"`
    return [`<pageField fld="${String(fieldIndex)}"${item}/>`]
  })
  if (fields.length === 0) {
    return ''
  }
  return [`<pageFields count="${String(fields.length)}">`, ...fields, '</pageFields>'].join('')
}

function pivotFieldItemIndex(cacheTable: PivotCacheTable, fieldIndex: number, value: LiteralInput): number {
  const values = uniqueValues(cacheTable.fields[fieldIndex]?.values ?? [])
  const index = values.findIndex((candidate) => literalValuesEqual(candidate, value))
  return index >= 0 ? index : 0
}

function pivotItemAxisXml(
  name: 'rowItems' | 'colItems',
  buckets: readonly PivotLayoutBucket[],
  fieldIndexes: readonly number[],
  cacheTable: PivotCacheTable,
): string {
  if (buckets.length === 0 || fieldIndexes.length === 0) {
    return ''
  }
  const items = buckets.map((bucket) => {
    const indexes = fieldIndexes.map(
      (fieldIndex, valueIndex) => `<x v="${String(pivotFieldItemIndex(cacheTable, fieldIndex, bucket.values[valueIndex] ?? null))}"/>`,
    )
    return `<i>${indexes.join('')}</i>`
  })
  return [`<${name} count="${String(items.length)}">`, ...items, `</${name}>`].join('')
}

function rowPassesPivotExportFilters(
  row: readonly LiteralInput[],
  fieldIndexesByName: ReadonlyMap<string, number>,
  pivot: WorkbookPivotSnapshot,
): boolean {
  for (const filter of pivot.filters ?? []) {
    const index = fieldIndexesByName.get(filter.sourceColumn)
    if (index === undefined) {
      return false
    }
    const value = row[index] ?? null
    if (filter.includedValues && !filter.includedValues.some((candidate) => literalValuesEqual(candidate, value))) {
      return false
    }
    if (filter.hiddenValues?.some((candidate) => literalValuesEqual(candidate, value))) {
      return false
    }
  }
  for (const pageField of pivot.pageFields ?? []) {
    const index = fieldIndexesByName.get(pageField.sourceColumn)
    if (index === undefined) {
      return false
    }
    if (pageField.selectedValue !== undefined && !literalValuesEqual(pageField.selectedValue, row[index] ?? null)) {
      return false
    }
  }
  for (const hiddenItems of pivot.hiddenItems ?? []) {
    const index = fieldIndexesByName.get(hiddenItems.sourceColumn)
    if (index === undefined) {
      return false
    }
    const value = row[index] ?? null
    if (hiddenItems.values.some((candidate) => literalValuesEqual(candidate, value))) {
      return false
    }
  }
  return true
}

function buildPivotLayoutBuckets(
  cacheTable: PivotCacheTable,
  pivot: WorkbookPivotSnapshot,
  fieldIndexes: readonly number[],
  fieldIndexesByName: ReadonlyMap<string, number>,
): PivotLayoutBucket[] {
  const seen = new Set<string>()
  const buckets: PivotLayoutBucket[] = []
  for (const row of cacheTable.rows) {
    if (!rowPassesPivotExportFilters(row, fieldIndexesByName, pivot)) {
      continue
    }
    const values = fieldIndexes.map((index) => row[index] ?? null)
    const key = values.map(literalValueKey).join('\u001f')
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    buckets.push({ values })
  }
  return buckets
}

function defaultDataFieldName(value: WorkbookPivotValueSnapshot): string {
  if (value.outputLabel && value.outputLabel.trim().length > 0) {
    return value.outputLabel.trim()
  }
  return `${defaultDataFieldVerb(value.summarizeBy)} of ${value.sourceColumn}`
}

function buildDataFieldsXml(values: readonly WorkbookPivotValueSnapshot[], fieldIndexesByName: ReadonlyMap<string, number>): string {
  if (values.length === 0) {
    return ''
  }
  const fields = values.flatMap((value) => {
    const fieldIndex = fieldIndexesByName.get(value.sourceColumn)
    if (fieldIndex === undefined) {
      return []
    }
    return [
      `<dataField name="${escapeXml(defaultDataFieldName(value))}" fld="${String(fieldIndex)}" subtotal="${subtotalValue(
        value.summarizeBy,
      )}"/>`,
    ]
  })
  return fields.length > 0 ? [`<dataFields count="${String(fields.length)}">`, ...fields, '</dataFields>'].join('') : ''
}

export function buildPivotTableDefinitionXml(pivot: WorkbookPivotSnapshot, cacheId: number, cacheTable: PivotCacheTable): string {
  const fieldNames = cacheTable.fields.map((field) => field.name)
  const fieldIndexesByName = new Map(fieldNames.map((name, index) => [name, index]))
  const rowFieldIndexes = pivot.groupBy.flatMap((fieldName) => {
    const index = fieldIndexesByName.get(fieldName)
    return index === undefined ? [] : [index]
  })
  const columnFieldIndexes = (pivot.columnFields ?? []).flatMap((fieldName) => {
    const index = fieldIndexesByName.get(fieldName)
    return index === undefined ? [] : [index]
  })
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotTableDefinition xmlns="${spreadsheetNamespace}" name="${escapeXml(pivot.name)}" cacheId="${String(
      cacheId,
    )}" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0" rowGrandTotals="0" colGrandTotals="0">`,
    `<location ref="${escapeXml(pivotOutputRange(pivot))}" firstHeaderRow="1" firstDataRow="2" firstDataCol="${String(
      Math.max(1, rowFieldIndexes.length),
    )}"/>`,
    `<pivotFields count="${String(fieldNames.length)}">`,
    ...fieldNames.map((_, index) => pivotFieldXml(index, fieldNames, pivot, cacheTable)),
    '</pivotFields>',
    buildRowFieldsXml(rowFieldIndexes),
    pivotItemAxisXml(
      'rowItems',
      buildPivotLayoutBuckets(cacheTable, pivot, rowFieldIndexes, fieldIndexesByName),
      rowFieldIndexes,
      cacheTable,
    ),
    buildColumnFieldsXml(columnFieldIndexes),
    pivotItemAxisXml(
      'colItems',
      buildPivotLayoutBuckets(cacheTable, pivot, columnFieldIndexes, fieldIndexesByName),
      columnFieldIndexes,
      cacheTable,
    ),
    buildPageFieldsXml(pivot.pageFields ?? [], fieldIndexesByName, cacheTable),
    buildDataFieldsXml(pivot.values, fieldIndexesByName),
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>',
    '</pivotTableDefinition>',
  ].join('')
}
