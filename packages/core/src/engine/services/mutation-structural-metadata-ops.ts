import { rewriteAddressForStructuralTransform, rewriteRangeForStructuralTransform, type StructuralAxisTransform } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { WorkbookStore } from '../../workbook-store.js'

type MutationStructuralMetadataWorkbook = Pick<
  WorkbookStore,
  | 'listDefinedNames'
  | 'listTables'
  | 'listSpills'
  | 'listPivots'
  | 'listCharts'
  | 'listImages'
  | 'listShapes'
  | 'listStyleRanges'
  | 'listFormatRanges'
  | 'listFilters'
  | 'listSorts'
  | 'listDataValidations'
  | 'listCommentThreads'
  | 'listNotes'
>

export function captureStructuralWorkbookMetadataOps(workbook: MutationStructuralMetadataWorkbook): EngineOp[] {
  const restoredOps: EngineOp[] = []
  workbook.listDefinedNames().forEach(({ name, value }) => {
    restoredOps.push({ kind: 'upsertDefinedName', name, value: structuredClone(value) })
  })
  workbook.listTables().forEach((table) => {
    restoredOps.push({
      kind: 'upsertTable',
      table: structuredClone(table),
    })
  })
  workbook.listSpills().forEach((spill) => {
    restoredOps.push({ kind: 'upsertSpillRange', sheetName: spill.sheetName, address: spill.address, rows: spill.rows, cols: spill.cols })
  })
  workbook.listPivots().forEach((pivot) => {
    if (!pivot.source) {
      return
    }
    restoredOps.push({
      kind: 'upsertPivotTable',
      name: pivot.name,
      sheetName: pivot.sheetName,
      address: pivot.address,
      source: { ...pivot.source },
      groupBy: [...pivot.groupBy],
      values: pivot.values.map((value) => Object.assign({}, value)),
      rows: pivot.rows,
      cols: pivot.cols,
    })
  })
  workbook.listCharts().forEach((chart) => {
    restoredOps.push({ kind: 'upsertChart', chart: structuredClone(chart) })
  })
  workbook.listImages().forEach((image) => {
    restoredOps.push({ kind: 'upsertImage', image: structuredClone(image) })
  })
  workbook.listShapes().forEach((shape) => {
    restoredOps.push({ kind: 'upsertShape', shape: structuredClone(shape) })
  })
  return restoredOps
}

export function clearStructuralSheetMetadataOps(
  workbook: MutationStructuralMetadataWorkbook,
  sheetName: string,
  transform: StructuralAxisTransform,
): EngineOp[] {
  const clearedOps: EngineOp[] = []
  workbook.listStyleRanges(sheetName).forEach((record) => {
    const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
    if (range) {
      clearedOps.push({
        kind: 'setStyleRange',
        range: { ...record.range, startAddress: range.startAddress, endAddress: range.endAddress },
        styleId: WorkbookStore.defaultStyleId,
      })
    }
  })
  workbook.listFormatRanges(sheetName).forEach((record) => {
    const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
    if (range) {
      clearedOps.push({
        kind: 'setFormatRange',
        range: { ...record.range, startAddress: range.startAddress, endAddress: range.endAddress },
        formatId: WorkbookStore.defaultFormatId,
      })
    }
  })
  workbook.listFilters(sheetName).forEach((filter) => {
    const range = rewriteRangeForStructuralTransform(filter.range.startAddress, filter.range.endAddress, transform)
    if (range) {
      clearedOps.push({
        kind: 'clearFilter',
        sheetName,
        range: { ...filter.range, startAddress: range.startAddress, endAddress: range.endAddress },
      })
    }
  })
  workbook.listSorts(sheetName).forEach((sort) => {
    const range = rewriteRangeForStructuralTransform(sort.range.startAddress, sort.range.endAddress, transform)
    if (range) {
      clearedOps.push({
        kind: 'clearSort',
        sheetName,
        range: { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
      })
    }
  })
  workbook.listDataValidations(sheetName).forEach((validation) => {
    const range = rewriteRangeForStructuralTransform(validation.range.startAddress, validation.range.endAddress, transform)
    if (range) {
      clearedOps.push({
        kind: 'clearDataValidation',
        sheetName,
        range: { ...validation.range, startAddress: range.startAddress, endAddress: range.endAddress },
      })
    }
  })
  workbook.listCommentThreads(sheetName).forEach((thread) => {
    const address = rewriteAddressForStructuralTransform(thread.address, transform)
    if (address) {
      clearedOps.push({ kind: 'deleteCommentThread', sheetName, address })
    }
  })
  workbook.listNotes(sheetName).forEach((note) => {
    const address = rewriteAddressForStructuralTransform(note.address, transform)
    if (address) {
      clearedOps.push({ kind: 'deleteNote', sheetName, address })
    }
  })
  return clearedOps
}
