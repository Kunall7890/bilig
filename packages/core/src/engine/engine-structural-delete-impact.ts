import type { CellRangeRef, CellSnapshot, WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import { parseCellAddress, rewriteFormulaForStructuralTransform } from '@bilig/formula'
import type { WorkbookStore } from '../workbook-store.js'
import { drawingArtifactsTouchStructuralDelete } from './services/structure-drawing-artifact-rewrite.js'
import {
  drawingChartPackageArtifactsTouchStructuralDelete,
  preservedChartPackageArtifactsTouchStructuralDelete,
} from './services/structure-chart-artifact-rewrite.js'
import { preservedSheetMetadataTouchesStructuralDelete } from './services/structure-preserved-sheet-metadata-rewrite.js'

type StructuralDeleteAxis = 'row' | 'column'

function formulaWouldRewriteForDelete(
  formula: string,
  sheetName: string,
  axis: StructuralDeleteAxis,
  start: number,
  count: number,
): boolean {
  const hasLeadingEquals = formula.startsWith('=')
  const source = hasLeadingEquals ? formula.slice(1) : formula
  const rewritten = rewriteFormulaForStructuralTransform(source, sheetName, sheetName, {
    kind: 'delete',
    axis,
    start,
    count,
  })
  return (hasLeadingEquals ? `=${rewritten}` : rewritten) !== formula
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
      .some(
        (table) =>
          table.sheetName === args.sheetName &&
          rangeTouchesAxisDelete(
            { sheetName: args.sheetName, startAddress: table.startAddress, endAddress: table.endAddress },
            args.axis,
            args.start,
          ),
      )
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
    preservedSheetMetadataTouchesStructuralDelete(args.workbook.metadata.preservedSheetMetadata.get(args.sheetName), args.axis, args.start)
  ) {
    return true
  }
  if (drawingArtifactsTouchStructuralDelete(args.workbook, args.sheetName)) {
    return true
  }
  if (drawingChartPackageArtifactsTouchStructuralDelete(args.workbook, args.sheetName, args.axis, args.start, args.count)) {
    return true
  }
  if (
    preservedChartPackageArtifactsTouchStructuralDelete(
      args.workbook.metadata.preservedWorkbookMetadata,
      args.sheetName,
      args.axis,
      args.start,
      args.count,
    )
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
