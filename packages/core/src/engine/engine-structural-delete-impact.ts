import type { CellRangeRef, CellSnapshot, WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import { formatAddress, parseCellAddress, rewriteFormulaForStructuralTransform, rewriteRangeForStructuralTransform } from '@bilig/formula'
import type { WorkbookStore, WorkbookTableRecord } from '../workbook-store.js'

type StructuralDeleteAxis = 'row' | 'column'

function formulaWouldRewriteForDelete(
  formula: string,
  sheetName: string,
  axis: StructuralDeleteAxis,
  start: number,
  count: number,
): boolean {
  return (
    rewriteFormulaForStructuralTransform(formula, sheetName, sheetName, {
      kind: 'delete',
      axis,
      start,
      count,
    }) !== formula
  )
}

function cellHasSemanticDeleteImpact(args: {
  readonly workbook: WorkbookStore
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly cellIndex: number
}): boolean {
  const snapshot = args.getCellByIndex(args.cellIndex)
  if (snapshot.formula !== undefined) {
    return true
  }
  if (args.workbook.getCellFormat(args.cellIndex) !== undefined) {
    return true
  }
  return snapshot.value.tag === ValueTag.Number || snapshot.value.tag === ValueTag.Boolean || snapshot.value.tag === ValueTag.String
}

function hasFormulaReferenceAtOrAfter(args: {
  readonly workbook: WorkbookStore
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly sheetName: string
  readonly axis: StructuralDeleteAxis
  readonly start: number
  readonly count: number
}): boolean {
  let found = false
  for (const sheet of args.workbook.sheetsByName.values()) {
    if (found) {
      break
    }
    sheet.grid.forEachCell((cellIndex) => {
      if (found) {
        return
      }
      const formulaId = args.workbook.cellStore.formulaIds[cellIndex] ?? 0
      if (formulaId === 0) {
        return
      }
      const snapshot = args.getCellByIndex(cellIndex)
      if (snapshot.formula && formulaWouldRewriteForDelete(snapshot.formula, args.sheetName, args.axis, args.start, args.count)) {
        found = true
      }
    })
  }
  return found
}

function definedNameTouchesAxisDelete(
  value: WorkbookDefinedNameValueSnapshot,
  sheetName: string,
  axis: StructuralDeleteAxis,
  start: number,
  count: number,
): boolean {
  if (typeof value === 'string') {
    return formulaWouldRewriteForDelete(value, sheetName, axis, start, count)
  }
  if (value === null || typeof value !== 'object') {
    return false
  }
  switch (value.kind) {
    case 'scalar':
    case 'structured-ref':
      return false
    case 'formula':
      return formulaWouldRewriteForDelete(value.formula, sheetName, axis, start, count)
    case 'cell-ref': {
      if (value.sheetName !== sheetName) {
        return false
      }
      const parsed = parseCellAddress(value.address, value.sheetName)
      return axis === 'row' ? parsed.row >= start : parsed.col >= start
    }
    case 'range-ref': {
      if (value.sheetName !== sheetName) {
        return false
      }
      const end = parseCellAddress(value.endAddress, value.sheetName)
      return axis === 'row' ? end.row >= start : end.col >= start
    }
  }
}

function rangeTouchesAxisDelete(range: CellRangeRef, axis: StructuralDeleteAxis, start: number): boolean {
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return axis === 'row' ? end.row >= start : end.col >= start
}

function tableRangeAfterStructuralDelete(
  table: WorkbookTableRecord,
  axis: StructuralDeleteAxis,
  start: number,
  count: number,
): CellRangeRef | undefined {
  const rewritten = rewriteRangeForStructuralTransform(table.startAddress, table.endAddress, {
    kind: 'delete',
    axis,
    start,
    count,
  })
  if (!rewritten) {
    return undefined
  }
  const range: CellRangeRef = {
    sheetName: table.sheetName,
    startAddress: rewritten.startAddress,
    endAddress: rewritten.endAddress,
  }
  const startCell = parseCellAddress(range.startAddress, range.sheetName)
  const endCell = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(startCell.row, endCell.row)
  const endRow = Math.max(startCell.row, endCell.row)
  const endCol = Math.max(startCell.col, endCell.col)
  const minimumRowCount = (table.headerRow ? 1 : 0) + (table.totalsRow ? 1 : 0) + 1
  if (endRow - startRow + 1 >= minimumRowCount) {
    return range
  }
  return {
    ...range,
    endAddress: formatAddress(startRow + minimumRowCount - 1, endCol),
  }
}

function tableWouldRewriteForDelete(table: WorkbookTableRecord, axis: StructuralDeleteAxis, start: number, count: number): boolean {
  const rewritten = tableRangeAfterStructuralDelete(table, axis, start, count)
  return rewritten === undefined || rewritten.startAddress !== table.startAddress || rewritten.endAddress !== table.endAddress
}

function addressTouchesAxisDelete(sheetName: string, address: string, axis: StructuralDeleteAxis, start: number): boolean {
  const parsed = parseCellAddress(address, sheetName)
  return axis === 'row' ? parsed.row >= start : parsed.col >= start
}

export function hasEngineStructuralDeleteImpact(args: {
  readonly workbook: WorkbookStore
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly sheetName: string
  readonly axis: StructuralDeleteAxis
  readonly start: number
  readonly count: number
}): boolean {
  const sheet = args.workbook.getSheet(args.sheetName)
  if (!sheet) {
    return false
  }
  if (
    sheet.grid.someCellInAxisScope(args.axis, { start: args.start }, (cellIndex) =>
      cellHasSemanticDeleteImpact({
        workbook: args.workbook,
        getCellByIndex: args.getCellByIndex,
        cellIndex,
      }),
    )
  ) {
    return true
  }
  const axisEntries =
    args.axis === 'row' ? args.workbook.listRowAxisEntries(args.sheetName) : args.workbook.listColumnAxisEntries(args.sheetName)
  if (axisEntries.some((entry) => entry.index >= args.start)) {
    return true
  }
  const axisMetadata =
    args.axis === 'row' ? args.workbook.listRowMetadata(args.sheetName) : args.workbook.listColumnMetadata(args.sheetName)
  if (axisMetadata.some((record) => record.start + record.count - 1 >= args.start)) {
    return true
  }
  if (args.workbook.listStyleRanges(args.sheetName).some((record) => rangeTouchesAxisDelete(record.range, args.axis, args.start))) {
    return true
  }
  if (args.workbook.listFormatRanges(args.sheetName).some((record) => rangeTouchesAxisDelete(record.range, args.axis, args.start))) {
    return true
  }
  const freezePane = args.workbook.getFreezePane(args.sheetName)
  if (freezePane && (args.axis === 'row' ? freezePane.rows > args.start : freezePane.cols > args.start)) {
    return true
  }
  if (args.workbook.listFilters(args.sheetName).some((record) => rangeTouchesAxisDelete(record.range, args.axis, args.start))) {
    return true
  }
  if (args.workbook.listSorts(args.sheetName).some((record) => rangeTouchesAxisDelete(record.range, args.axis, args.start))) {
    return true
  }
  if (args.workbook.listDataValidations(args.sheetName).some((record) => rangeTouchesAxisDelete(record.range, args.axis, args.start))) {
    return true
  }
  if (args.workbook.listConditionalFormats(args.sheetName).some((record) => rangeTouchesAxisDelete(record.range, args.axis, args.start))) {
    return true
  }
  if (args.workbook.listRangeProtections(args.sheetName).some((record) => rangeTouchesAxisDelete(record.range, args.axis, args.start))) {
    return true
  }
  if (
    args.workbook
      .listCommentThreads(args.sheetName)
      .some((record) => addressTouchesAxisDelete(args.sheetName, record.address, args.axis, args.start))
  ) {
    return true
  }
  if (
    args.workbook
      .listNotes(args.sheetName)
      .some((record) => addressTouchesAxisDelete(args.sheetName, record.address, args.axis, args.start))
  ) {
    return true
  }
  if (
    args.workbook
      .listTables()
      .some((table) => table.sheetName === args.sheetName && tableWouldRewriteForDelete(table, args.axis, args.start, args.count))
  ) {
    return true
  }
  if (
    args.workbook
      .listSpills()
      .some((spill) => spill.sheetName === args.sheetName && addressTouchesAxisDelete(args.sheetName, spill.address, args.axis, args.start))
  ) {
    return true
  }
  if (
    args.workbook
      .listPivots()
      .some(
        (pivot) =>
          (pivot.sheetName === args.sheetName && addressTouchesAxisDelete(args.sheetName, pivot.address, args.axis, args.start)) ||
          (pivot.source?.sheetName === args.sheetName && rangeTouchesAxisDelete(pivot.source, args.axis, args.start)),
      )
  ) {
    return true
  }
  if (
    args.workbook
      .listCharts()
      .some(
        (chart) =>
          (chart.sheetName === args.sheetName && addressTouchesAxisDelete(args.sheetName, chart.address, args.axis, args.start)) ||
          (chart.source.sheetName === args.sheetName && rangeTouchesAxisDelete(chart.source, args.axis, args.start)),
      )
  ) {
    return true
  }
  if (
    args.workbook
      .listImages()
      .some((image) => image.sheetName === args.sheetName && addressTouchesAxisDelete(args.sheetName, image.address, args.axis, args.start))
  ) {
    return true
  }
  if (
    args.workbook
      .listShapes()
      .some((shape) => shape.sheetName === args.sheetName && addressTouchesAxisDelete(args.sheetName, shape.address, args.axis, args.start))
  ) {
    return true
  }
  if (
    args.workbook
      .listDefinedNames()
      .some((definedName) => definedNameTouchesAxisDelete(definedName.value, args.sheetName, args.axis, args.start, args.count))
  ) {
    return true
  }
  return hasFormulaReferenceAtOrAfter(args)
}
