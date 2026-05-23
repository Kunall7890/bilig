import type {
  LiteralInput,
  WorkbookChartLegendPosition,
  WorkbookChartSeriesOrientation,
  WorkbookChartSnapshot,
  WorkbookChartType,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatRuleSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationRuleSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookHyperlinkSnapshot,
  WorkbookImageSnapshot,
  WorkbookNoteSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookShapeSnapshot,
  WorkbookShapeType,
  WorkbookSheetProtectionSnapshot,
  WorkbookValidationListSourceSnapshot,
} from '@bilig/protocol'
import type { WorkbookTableOp } from '@bilig/workbook'
import { BinaryProtocolError, type BinaryReader, type BinaryWriter } from './binary-io.js'
import {
  assertNever,
  decodeCellRangeRef,
  decodeCellStylePatch,
  decodeLiteral,
  decodePivotAggregation,
  encodeCellRangeRef,
  encodeCellStylePatch,
  encodeLiteral,
  encodePivotAggregation,
} from './binary-value-codec.js'

function decodeValidationComparisonOperator(
  value: string,
): 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual' {
  switch (value) {
    case 'between':
    case 'notBetween':
    case 'equal':
    case 'notEqual':
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual':
      return value
    default:
      throw new BinaryProtocolError('Unknown data validation comparison operator')
  }
}

function decodeValidationErrorStyle(value: string): 'stop' | 'warning' | 'information' {
  switch (value) {
    case 'stop':
    case 'warning':
    case 'information':
      return value
    default:
      throw new BinaryProtocolError('Unknown data validation error style')
  }
}

function encodeChartType(writer: BinaryWriter, chartType: WorkbookChartType): void {
  switch (chartType) {
    case 'column':
      writer.u8(1)
      return
    case 'bar':
      writer.u8(2)
      return
    case 'line':
      writer.u8(3)
      return
    case 'area':
      writer.u8(4)
      return
    case 'pie':
      writer.u8(5)
      return
    case 'scatter':
      writer.u8(6)
      return
  }
}

function decodeChartType(reader: BinaryReader): WorkbookChartType {
  switch (reader.u8()) {
    case 1:
      return 'column'
    case 2:
      return 'bar'
    case 3:
      return 'line'
    case 4:
      return 'area'
    case 5:
      return 'pie'
    case 6:
      return 'scatter'
    default:
      throw new BinaryProtocolError('Unknown chart type tag')
  }
}

function encodeChartSeriesOrientation(writer: BinaryWriter, orientation: WorkbookChartSeriesOrientation): void {
  writer.u8(orientation === 'rows' ? 1 : 2)
}

function decodeChartSeriesOrientation(reader: BinaryReader): WorkbookChartSeriesOrientation {
  switch (reader.u8()) {
    case 1:
      return 'rows'
    case 2:
      return 'columns'
    default:
      throw new BinaryProtocolError('Unknown chart orientation tag')
  }
}

function encodeChartLegendPosition(writer: BinaryWriter, position: WorkbookChartLegendPosition): void {
  switch (position) {
    case 'top':
      writer.u8(1)
      return
    case 'right':
      writer.u8(2)
      return
    case 'bottom':
      writer.u8(3)
      return
    case 'left':
      writer.u8(4)
      return
    case 'hidden':
      writer.u8(5)
      return
  }
}

function decodeChartLegendPosition(reader: BinaryReader): WorkbookChartLegendPosition {
  switch (reader.u8()) {
    case 1:
      return 'top'
    case 2:
      return 'right'
    case 3:
      return 'bottom'
    case 4:
      return 'left'
    case 5:
      return 'hidden'
    default:
      throw new BinaryProtocolError('Unknown chart legend position tag')
  }
}

export function encodeChart(writer: BinaryWriter, chart: WorkbookChartSnapshot): void {
  writer.string(chart.id)
  writer.string(chart.sheetName)
  writer.string(chart.address)
  encodeCellRangeRef(writer, chart.source)
  encodeChartType(writer, chart.chartType)
  writer.bool(chart.seriesOrientation !== undefined)
  if (chart.seriesOrientation !== undefined) {
    encodeChartSeriesOrientation(writer, chart.seriesOrientation)
  }
  writer.bool(chart.firstRowAsHeaders !== undefined)
  if (chart.firstRowAsHeaders !== undefined) {
    writer.bool(chart.firstRowAsHeaders)
  }
  writer.bool(chart.firstColumnAsLabels !== undefined)
  if (chart.firstColumnAsLabels !== undefined) {
    writer.bool(chart.firstColumnAsLabels)
  }
  writer.bool(chart.title !== undefined)
  if (chart.title !== undefined) {
    writer.string(chart.title)
  }
  writer.bool(chart.legendPosition !== undefined)
  if (chart.legendPosition !== undefined) {
    encodeChartLegendPosition(writer, chart.legendPosition)
  }
  writer.u32(chart.rows)
  writer.u32(chart.cols)
}

export function decodeChart(reader: BinaryReader): WorkbookChartSnapshot {
  const chart: WorkbookChartSnapshot = {
    id: reader.string(),
    sheetName: reader.string(),
    address: reader.string(),
    source: decodeCellRangeRef(reader),
    chartType: decodeChartType(reader),
    rows: 0,
    cols: 0,
  }
  const hasSeriesOrientation = reader.bool()
  if (hasSeriesOrientation) {
    chart.seriesOrientation = decodeChartSeriesOrientation(reader)
  }
  const hasFirstRowAsHeaders = reader.bool()
  if (hasFirstRowAsHeaders) {
    chart.firstRowAsHeaders = reader.bool()
  }
  const hasFirstColumnAsLabels = reader.bool()
  if (hasFirstColumnAsLabels) {
    chart.firstColumnAsLabels = reader.bool()
  }
  const hasTitle = reader.bool()
  if (hasTitle) {
    chart.title = reader.string()
  }
  const hasLegendPosition = reader.bool()
  if (hasLegendPosition) {
    chart.legendPosition = decodeChartLegendPosition(reader)
  }
  chart.rows = reader.u32()
  chart.cols = reader.u32()
  return chart
}

export function encodeImage(writer: BinaryWriter, image: WorkbookImageSnapshot): void {
  writer.string(image.id)
  writer.string(image.sheetName)
  writer.string(image.address)
  writer.string(image.sourceUrl)
  writer.u32(image.rows)
  writer.u32(image.cols)
  writer.bool(image.altText !== undefined)
  if (image.altText !== undefined) {
    writer.string(image.altText)
  }
}

export function decodeImage(reader: BinaryReader): WorkbookImageSnapshot {
  const image: WorkbookImageSnapshot = {
    id: reader.string(),
    sheetName: reader.string(),
    address: reader.string(),
    sourceUrl: reader.string(),
    rows: reader.u32(),
    cols: reader.u32(),
  }
  if (reader.bool()) {
    image.altText = reader.string()
  }
  return image
}

function encodeShapeType(writer: BinaryWriter, shapeType: WorkbookShapeType): void {
  switch (shapeType) {
    case 'rectangle':
      writer.u8(0)
      return
    case 'roundedRectangle':
      writer.u8(1)
      return
    case 'ellipse':
      writer.u8(2)
      return
    case 'line':
      writer.u8(3)
      return
    case 'arrow':
      writer.u8(4)
      return
    case 'textBox':
      writer.u8(5)
      return
    default:
      assertNever(shapeType)
  }
}

function decodeShapeType(reader: BinaryReader): WorkbookShapeType {
  switch (reader.u8()) {
    case 0:
      return 'rectangle'
    case 1:
      return 'roundedRectangle'
    case 2:
      return 'ellipse'
    case 3:
      return 'line'
    case 4:
      return 'arrow'
    case 5:
      return 'textBox'
    default:
      throw new BinaryProtocolError('Unknown shape type tag')
  }
}

export function encodeShape(writer: BinaryWriter, shape: WorkbookShapeSnapshot): void {
  writer.string(shape.id)
  writer.string(shape.sheetName)
  writer.string(shape.address)
  encodeShapeType(writer, shape.shapeType)
  writer.u32(shape.rows)
  writer.u32(shape.cols)
  writer.bool(shape.text !== undefined)
  if (shape.text !== undefined) {
    writer.string(shape.text)
  }
  writer.bool(shape.fillColor !== undefined)
  if (shape.fillColor !== undefined) {
    writer.string(shape.fillColor)
  }
  writer.bool(shape.strokeColor !== undefined)
  if (shape.strokeColor !== undefined) {
    writer.string(shape.strokeColor)
  }
}

export function decodeShape(reader: BinaryReader): WorkbookShapeSnapshot {
  const shape: WorkbookShapeSnapshot = {
    id: reader.string(),
    sheetName: reader.string(),
    address: reader.string(),
    shapeType: decodeShapeType(reader),
    rows: reader.u32(),
    cols: reader.u32(),
  }
  if (reader.bool()) {
    shape.text = reader.string()
  }
  if (reader.bool()) {
    shape.fillColor = reader.string()
  }
  if (reader.bool()) {
    shape.strokeColor = reader.string()
  }
  return shape
}

export function encodePivotValue(writer: BinaryWriter, value: WorkbookPivotValueSnapshot): void {
  writer.string(value.sourceColumn)
  encodePivotAggregation(writer, value.summarizeBy)
  writer.bool(value.outputLabel !== undefined)
  if (value.outputLabel !== undefined) {
    writer.string(value.outputLabel)
  }
}

export function decodePivotValue(reader: BinaryReader): WorkbookPivotValueSnapshot {
  const sourceColumn = reader.string()
  const summarizeBy = decodePivotAggregation(reader)
  const hasLabel = reader.bool()
  const result: WorkbookPivotValueSnapshot = {
    sourceColumn,
    summarizeBy,
  }
  if (hasLabel) {
    result.outputLabel = reader.string()
  }
  return result
}

export function encodeTable(writer: BinaryWriter, table: WorkbookTableOp): void {
  writer.string(table.name)
  writer.string(table.sheetName)
  writer.string(table.startAddress)
  writer.string(table.endAddress)
  writer.stringArray(table.columnNames)
  writer.bool(table.headerRow)
  writer.bool(table.totalsRow)
}

export function decodeTable(reader: BinaryReader): WorkbookTableOp {
  return {
    name: reader.string(),
    sheetName: reader.string(),
    startAddress: reader.string(),
    endAddress: reader.string(),
    columnNames: reader.stringArray(),
    headerRow: reader.bool(),
    totalsRow: reader.bool(),
  }
}

function encodeValidationListSource(writer: BinaryWriter, source: WorkbookValidationListSourceSnapshot): void {
  switch (source.kind) {
    case 'named-range':
      writer.u8(0)
      writer.string(source.name)
      return
    case 'cell-ref':
      writer.u8(1)
      writer.string(source.sheetName)
      writer.string(source.address)
      return
    case 'range-ref':
      writer.u8(2)
      writer.string(source.sheetName)
      writer.string(source.startAddress)
      writer.string(source.endAddress)
      return
    case 'structured-ref':
      writer.u8(3)
      writer.string(source.tableName)
      writer.string(source.columnName)
      return
    default:
      assertNever(source)
  }
}

function decodeValidationListSource(reader: BinaryReader): WorkbookValidationListSourceSnapshot {
  switch (reader.u8()) {
    case 0:
      return { kind: 'named-range', name: reader.string() }
    case 1:
      return { kind: 'cell-ref', sheetName: reader.string(), address: reader.string() }
    case 2:
      return {
        kind: 'range-ref',
        sheetName: reader.string(),
        startAddress: reader.string(),
        endAddress: reader.string(),
      }
    case 3:
      return {
        kind: 'structured-ref',
        tableName: reader.string(),
        columnName: reader.string(),
      }
    default:
      throw new BinaryProtocolError('Unknown validation source tag')
  }
}

function encodeDataValidationRule(writer: BinaryWriter, rule: WorkbookDataValidationRuleSnapshot): void {
  switch (rule.kind) {
    case 'list':
      writer.u8(0)
      writer.bool(rule.values !== undefined)
      if (rule.values) {
        writer.u32(rule.values.length)
        rule.values.forEach((value) => encodeLiteral(writer, value))
      }
      writer.bool(rule.source !== undefined)
      if (rule.source) {
        encodeValidationListSource(writer, rule.source)
      }
      return
    case 'checkbox':
      writer.u8(1)
      writer.bool(rule.checkedValue !== undefined)
      if (rule.checkedValue !== undefined) {
        encodeLiteral(writer, rule.checkedValue)
      }
      writer.bool(rule.uncheckedValue !== undefined)
      if (rule.uncheckedValue !== undefined) {
        encodeLiteral(writer, rule.uncheckedValue)
      }
      return
    case 'any':
      writer.u8(7)
      return
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength':
      writer.u8(
        {
          whole: 2,
          decimal: 3,
          date: 4,
          time: 5,
          textLength: 6,
        }[rule.kind],
      )
      writer.string(rule.operator)
      writer.u32(rule.values.length)
      rule.values.forEach((value) => encodeLiteral(writer, value))
      return
    default:
      assertNever(rule)
  }
}

function decodeDataValidationRule(reader: BinaryReader): WorkbookDataValidationRuleSnapshot {
  const tag = reader.u8()
  switch (tag) {
    case 0: {
      const hasValues = reader.bool()
      const values = hasValues
        ? (() => {
            const count = reader.u32()
            const items: LiteralInput[] = []
            for (let index = 0; index < count; index += 1) {
              items.push(decodeLiteral(reader))
            }
            return items
          })()
        : undefined
      const hasSource = reader.bool()
      const rule: Extract<WorkbookDataValidationRuleSnapshot, { kind: 'list' }> = {
        kind: 'list',
      }
      if (values) {
        rule.values = values
      }
      if (hasSource) {
        rule.source = decodeValidationListSource(reader)
      }
      return rule
    }
    case 1: {
      const hasCheckedValue = reader.bool()
      const checkedValue = hasCheckedValue ? decodeLiteral(reader) : undefined
      const hasUncheckedValue = reader.bool()
      const rule: Extract<WorkbookDataValidationRuleSnapshot, { kind: 'checkbox' }> = {
        kind: 'checkbox',
      }
      if (checkedValue !== undefined) {
        rule.checkedValue = checkedValue
      }
      if (hasUncheckedValue) {
        rule.uncheckedValue = decodeLiteral(reader)
      }
      return rule
    }
    case 2:
    case 3:
    case 4:
    case 5:
    case 6: {
      const operator = decodeValidationComparisonOperator(reader.string())
      const count = reader.u32()
      const values: LiteralInput[] = []
      for (let index = 0; index < count; index += 1) {
        values.push(decodeLiteral(reader))
      }
      const kind = {
        2: 'whole',
        3: 'decimal',
        4: 'date',
        5: 'time',
        6: 'textLength',
      } as const satisfies Record<2 | 3 | 4 | 5 | 6, 'whole' | 'decimal' | 'date' | 'time' | 'textLength'>
      const nextKind = kind[tag]
      if (!nextKind) {
        throw new BinaryProtocolError('Unknown scalar data validation rule tag')
      }
      return {
        kind: nextKind,
        operator,
        values,
      }
    }
    case 7:
      return { kind: 'any' }
    default:
      throw new BinaryProtocolError('Unknown data validation rule tag')
  }
}

export function encodeDataValidation(writer: BinaryWriter, validation: WorkbookDataValidationSnapshot): void {
  encodeCellRangeRef(writer, validation.range)
  encodeDataValidationRule(writer, validation.rule)
  writer.bool(validation.allowBlank !== undefined)
  if (validation.allowBlank !== undefined) {
    writer.bool(validation.allowBlank)
  }
  writer.bool(validation.showDropdown !== undefined)
  if (validation.showDropdown !== undefined) {
    writer.bool(validation.showDropdown)
  }
  writer.bool(validation.promptTitle !== undefined)
  if (validation.promptTitle !== undefined) {
    writer.string(validation.promptTitle)
  }
  writer.bool(validation.promptMessage !== undefined)
  if (validation.promptMessage !== undefined) {
    writer.string(validation.promptMessage)
  }
  writer.bool(validation.errorStyle !== undefined)
  if (validation.errorStyle !== undefined) {
    writer.string(validation.errorStyle)
  }
  writer.bool(validation.errorTitle !== undefined)
  if (validation.errorTitle !== undefined) {
    writer.string(validation.errorTitle)
  }
  writer.bool(validation.errorMessage !== undefined)
  if (validation.errorMessage !== undefined) {
    writer.string(validation.errorMessage)
  }
}

export function decodeDataValidation(reader: BinaryReader): WorkbookDataValidationSnapshot {
  const range = decodeCellRangeRef(reader)
  const rule = decodeDataValidationRule(reader)
  const validation: WorkbookDataValidationSnapshot = {
    range,
    rule,
  }
  if (reader.bool()) {
    validation.allowBlank = reader.bool()
  }
  if (reader.bool()) {
    validation.showDropdown = reader.bool()
  }
  if (reader.bool()) {
    validation.promptTitle = reader.string()
  }
  if (reader.bool()) {
    validation.promptMessage = reader.string()
  }
  if (reader.bool()) {
    validation.errorStyle = decodeValidationErrorStyle(reader.string())
  }
  if (reader.bool()) {
    validation.errorTitle = reader.string()
  }
  if (reader.bool()) {
    validation.errorMessage = reader.string()
  }
  return validation
}

function encodeConditionalFormatRule(writer: BinaryWriter, rule: WorkbookConditionalFormatRuleSnapshot): void {
  switch (rule.kind) {
    case 'cellIs':
      writer.u8(0)
      writer.string(rule.operator)
      writer.u32(rule.values.length)
      rule.values.forEach((value) => encodeLiteral(writer, value))
      return
    case 'textContains':
      writer.u8(1)
      writer.string(rule.text)
      writer.bool(rule.caseSensitive !== undefined)
      if (rule.caseSensitive !== undefined) {
        writer.bool(rule.caseSensitive)
      }
      return
    case 'formula':
      writer.u8(2)
      writer.string(rule.formula)
      return
    case 'blanks':
      writer.u8(3)
      return
    case 'notBlanks':
      writer.u8(4)
      return
    default:
      assertNever(rule)
  }
}

function decodeConditionalFormatRule(reader: BinaryReader): WorkbookConditionalFormatRuleSnapshot {
  switch (reader.u8()) {
    case 0: {
      const operator = decodeValidationComparisonOperator(reader.string())
      const count = reader.u32()
      const values: LiteralInput[] = []
      for (let index = 0; index < count; index += 1) {
        values.push(decodeLiteral(reader))
      }
      return {
        kind: 'cellIs',
        operator,
        values,
      }
    }
    case 1: {
      const rule: Extract<WorkbookConditionalFormatRuleSnapshot, { kind: 'textContains' }> = {
        kind: 'textContains',
        text: reader.string(),
      }
      if (reader.bool()) {
        rule.caseSensitive = reader.bool()
      }
      return rule
    }
    case 2:
      return { kind: 'formula', formula: reader.string() }
    case 3:
      return { kind: 'blanks' }
    case 4:
      return { kind: 'notBlanks' }
    default:
      throw new BinaryProtocolError('Unknown conditional format rule tag')
  }
}

export function encodeConditionalFormat(writer: BinaryWriter, format: WorkbookConditionalFormatSnapshot): void {
  writer.string(format.id)
  encodeCellRangeRef(writer, format.range)
  encodeConditionalFormatRule(writer, format.rule)
  encodeCellStylePatch(writer, format.style)
  writer.bool(format.stopIfTrue !== undefined)
  if (format.stopIfTrue !== undefined) {
    writer.bool(format.stopIfTrue)
  }
  writer.bool(format.priority !== undefined)
  if (format.priority !== undefined) {
    writer.f64(format.priority)
  }
}

export function decodeConditionalFormat(reader: BinaryReader): WorkbookConditionalFormatSnapshot {
  const format: WorkbookConditionalFormatSnapshot = {
    id: reader.string(),
    range: decodeCellRangeRef(reader),
    rule: decodeConditionalFormatRule(reader),
    style: decodeCellStylePatch(reader),
  }
  if (reader.bool()) {
    format.stopIfTrue = reader.bool()
  }
  if (reader.bool()) {
    format.priority = reader.f64()
  }
  return format
}

export function encodeSheetProtection(writer: BinaryWriter, protection: WorkbookSheetProtectionSnapshot): void {
  writer.string(protection.sheetName)
  writer.bool(protection.hideFormulas !== undefined)
  if (protection.hideFormulas !== undefined) {
    writer.bool(protection.hideFormulas)
  }
}

export function decodeSheetProtection(reader: BinaryReader): WorkbookSheetProtectionSnapshot {
  const protection: WorkbookSheetProtectionSnapshot = {
    sheetName: reader.string(),
  }
  if (reader.bool()) {
    protection.hideFormulas = reader.bool()
  }
  return protection
}

export function encodeRangeProtection(writer: BinaryWriter, protection: WorkbookRangeProtectionSnapshot): void {
  writer.string(protection.id)
  encodeCellRangeRef(writer, protection.range)
  writer.bool(protection.hideFormulas !== undefined)
  if (protection.hideFormulas !== undefined) {
    writer.bool(protection.hideFormulas)
  }
}

export function decodeRangeProtection(reader: BinaryReader): WorkbookRangeProtectionSnapshot {
  const protection: WorkbookRangeProtectionSnapshot = {
    id: reader.string(),
    range: decodeCellRangeRef(reader),
  }
  if (reader.bool()) {
    protection.hideFormulas = reader.bool()
  }
  return protection
}

function encodeCommentEntry(writer: BinaryWriter, entry: WorkbookCommentThreadSnapshot['comments'][number]): void {
  writer.string(entry.id)
  writer.string(entry.body)
  writer.bool(entry.authorUserId !== undefined)
  if (entry.authorUserId !== undefined) {
    writer.string(entry.authorUserId)
  }
  writer.bool(entry.authorDisplayName !== undefined)
  if (entry.authorDisplayName !== undefined) {
    writer.string(entry.authorDisplayName)
  }
  writer.bool(entry.createdAtUnixMs !== undefined)
  if (entry.createdAtUnixMs !== undefined) {
    writer.u32(entry.createdAtUnixMs)
  }
}

function decodeCommentEntry(reader: BinaryReader): WorkbookCommentThreadSnapshot['comments'][number] {
  const entry: WorkbookCommentThreadSnapshot['comments'][number] = {
    id: reader.string(),
    body: reader.string(),
  }
  if (reader.bool()) {
    entry.authorUserId = reader.string()
  }
  if (reader.bool()) {
    entry.authorDisplayName = reader.string()
  }
  if (reader.bool()) {
    entry.createdAtUnixMs = reader.u32()
  }
  return entry
}

export function encodeCommentThread(writer: BinaryWriter, thread: WorkbookCommentThreadSnapshot): void {
  writer.string(thread.threadId)
  writer.string(thread.sheetName)
  writer.string(thread.address)
  writer.u32(thread.comments.length)
  thread.comments.forEach((entry) => encodeCommentEntry(writer, entry))
  writer.bool(thread.resolved !== undefined)
  if (thread.resolved !== undefined) {
    writer.bool(thread.resolved)
  }
  writer.bool(thread.resolvedByUserId !== undefined)
  if (thread.resolvedByUserId !== undefined) {
    writer.string(thread.resolvedByUserId)
  }
  writer.bool(thread.resolvedAtUnixMs !== undefined)
  if (thread.resolvedAtUnixMs !== undefined) {
    writer.u32(thread.resolvedAtUnixMs)
  }
}

export function decodeCommentThread(reader: BinaryReader): WorkbookCommentThreadSnapshot {
  const thread: WorkbookCommentThreadSnapshot = {
    threadId: reader.string(),
    sheetName: reader.string(),
    address: reader.string(),
    comments: [],
  }
  const count = reader.u32()
  for (let index = 0; index < count; index += 1) {
    thread.comments.push(decodeCommentEntry(reader))
  }
  if (reader.bool()) {
    thread.resolved = reader.bool()
  }
  if (reader.bool()) {
    thread.resolvedByUserId = reader.string()
  }
  if (reader.bool()) {
    thread.resolvedAtUnixMs = reader.u32()
  }
  return thread
}

export function encodeNote(writer: BinaryWriter, note: WorkbookNoteSnapshot): void {
  writer.string(note.sheetName)
  writer.string(note.address)
  writer.string(note.text)
}

export function decodeNote(reader: BinaryReader): WorkbookNoteSnapshot {
  return {
    sheetName: reader.string(),
    address: reader.string(),
    text: reader.string(),
  }
}

export function encodeHyperlink(writer: BinaryWriter, hyperlink: WorkbookHyperlinkSnapshot): void {
  writer.string(hyperlink.sheetName)
  writer.string(hyperlink.address)
  writer.string(hyperlink.target)
  writer.bool(hyperlink.tooltip !== undefined)
  if (hyperlink.tooltip !== undefined) {
    writer.string(hyperlink.tooltip)
  }
  writer.bool(hyperlink.display !== undefined)
  if (hyperlink.display !== undefined) {
    writer.string(hyperlink.display)
  }
}

export function decodeHyperlink(reader: BinaryReader): WorkbookHyperlinkSnapshot {
  const hyperlink: WorkbookHyperlinkSnapshot = {
    sheetName: reader.string(),
    address: reader.string(),
    target: reader.string(),
  }
  if (reader.bool()) {
    hyperlink.tooltip = reader.string()
  }
  if (reader.bool()) {
    hyperlink.display = reader.string()
  }
  return hyperlink
}
