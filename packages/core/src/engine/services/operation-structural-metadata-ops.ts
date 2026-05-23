import { MAX_COLS, MAX_ROWS, type CellRangeRef, type CellStyleRecord } from '@bilig/protocol'
import { formatAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import type { OpOrder } from '../../replica-state.js'
import { WORKBOOK_DEFAULT_STYLE_ID } from '../../workbook-default-style-format.js'
import { cellStyleKey, normalizeCellStyleRecord } from '../../workbook-store-records.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { assertNever } from './operation-change-helpers.js'
import type { MutationSource } from './operation-service-types.js'

export type OperationStructuralMetadataOp = Extract<
  EngineOp,
  {
    kind:
      | 'updateRowMetadata'
      | 'updateColumnMetadata'
      | 'setFreezePane'
      | 'clearFreezePane'
      | 'mergeCells'
      | 'unmergeCells'
      | 'setSheetProtection'
      | 'clearSheetProtection'
      | 'setFilter'
      | 'clearFilter'
      | 'setSort'
      | 'clearSort'
      | 'setDataValidation'
      | 'clearDataValidation'
      | 'upsertConditionalFormat'
      | 'deleteConditionalFormat'
      | 'upsertRangeProtection'
      | 'deleteRangeProtection'
      | 'upsertCommentThread'
      | 'deleteCommentThread'
      | 'upsertNote'
      | 'deleteNote'
      | 'upsertCellStyle'
      | 'upsertCellNumberFormat'
      | 'setStyleRange'
      | 'setFormatRange'
      | 'upsertChart'
      | 'deleteChart'
      | 'upsertImage'
      | 'deleteImage'
      | 'upsertShape'
      | 'deleteShape'
  }
>

type OperationStructuralMetadataWorkbook = Pick<
  WorkbookStore,
  | 'setRowMetadata'
  | 'setColumnMetadata'
  | 'setFreezePane'
  | 'clearFreezePane'
  | 'setMergeRange'
  | 'clearMergeRanges'
  | 'setSheetProtection'
  | 'clearSheetProtection'
  | 'setFilter'
  | 'deleteFilter'
  | 'setSort'
  | 'deleteSort'
  | 'setDataValidation'
  | 'deleteDataValidation'
  | 'setConditionalFormat'
  | 'deleteConditionalFormat'
  | 'setRangeProtection'
  | 'deleteRangeProtection'
  | 'setCommentThread'
  | 'deleteCommentThread'
  | 'setNote'
  | 'deleteNote'
  | 'getCellStyle'
  | 'listStyleRanges'
  | 'upsertCellStyle'
  | 'upsertCellNumberFormat'
  | 'setStyleRange'
  | 'coalesceStyleRanges'
  | 'setFormatRange'
  | 'setChart'
  | 'deleteChart'
  | 'setImage'
  | 'deleteImage'
  | 'setShape'
  | 'deleteShape'
  | 'sheetsByName'
>

interface OperationStructuralMetadataInvalidationSpan {
  readonly sheetName: string
  readonly startIndex: number
  readonly endIndex: number
}

export interface OperationStructuralMetadataChange {
  structuralInvalidation: boolean
  invalidatedRanges: CellRangeRef[]
  invalidatedRows: OperationStructuralMetadataInvalidationSpan[]
  invalidatedColumns: OperationStructuralMetadataInvalidationSpan[]
}

export function applyOperationStructuralMetadataOp(args: {
  readonly workbook: OperationStructuralMetadataWorkbook
  readonly op: OperationStructuralMetadataOp
  readonly order: OpOrder
  readonly source: MutationSource
  readonly setEntityVersionForOp: (op: EngineOp, order: OpOrder) => void
}): OperationStructuralMetadataChange {
  const change: OperationStructuralMetadataChange = {
    structuralInvalidation: false,
    invalidatedRanges: [],
    invalidatedRows: [],
    invalidatedColumns: [],
  }

  switch (args.op.kind) {
    case 'updateRowMetadata':
      args.workbook.setRowMetadata(
        args.op.sheetName,
        args.op.start,
        args.op.count,
        args.op.size,
        args.op.hidden,
        args.op.geometry,
        args.op.filterHidden,
      )
      change.invalidatedRows.push({
        sheetName: args.op.sheetName,
        startIndex: args.op.start,
        endIndex: args.op.start + args.op.count - 1,
      })
      break
    case 'updateColumnMetadata':
      args.workbook.setColumnMetadata(args.op.sheetName, args.op.start, args.op.count, args.op.size, args.op.hidden)
      change.invalidatedColumns.push({
        sheetName: args.op.sheetName,
        startIndex: args.op.start,
        endIndex: args.op.start + args.op.count - 1,
      })
      break
    case 'setFreezePane':
      args.workbook.setFreezePane(args.op.sheetName, args.op.rows, args.op.cols)
      change.structuralInvalidation = true
      break
    case 'clearFreezePane':
      args.workbook.clearFreezePane(args.op.sheetName)
      change.structuralInvalidation = true
      break
    case 'mergeCells':
      args.workbook.setMergeRange(args.op.range)
      change.invalidatedRanges.push({ ...args.op.range })
      change.structuralInvalidation = true
      break
    case 'unmergeCells':
      args.workbook.clearMergeRanges(args.op.range)
      change.invalidatedRanges.push({ ...args.op.range })
      change.structuralInvalidation = true
      break
    case 'setSheetProtection':
      args.workbook.setSheetProtection(args.op.protection)
      change.structuralInvalidation = true
      break
    case 'clearSheetProtection':
      args.workbook.clearSheetProtection(args.op.sheetName)
      change.structuralInvalidation = true
      break
    case 'setFilter':
      args.workbook.setFilter(args.op.sheetName, args.op.range)
      change.structuralInvalidation = true
      break
    case 'clearFilter':
      args.workbook.deleteFilter(args.op.sheetName, args.op.range)
      change.structuralInvalidation = true
      break
    case 'setSort':
      args.workbook.setSort(args.op.sheetName, args.op.range, args.op.keys)
      change.structuralInvalidation = true
      break
    case 'clearSort':
      args.workbook.deleteSort(args.op.sheetName, args.op.range)
      change.structuralInvalidation = true
      break
    case 'setDataValidation':
      args.workbook.setDataValidation(args.op.validation)
      change.structuralInvalidation = true
      break
    case 'clearDataValidation':
      args.workbook.deleteDataValidation(args.op.sheetName, args.op.range)
      change.structuralInvalidation = true
      break
    case 'upsertConditionalFormat':
      args.workbook.setConditionalFormat(args.op.format)
      change.structuralInvalidation = true
      break
    case 'deleteConditionalFormat':
      args.workbook.deleteConditionalFormat(args.op.id)
      change.structuralInvalidation = true
      break
    case 'upsertRangeProtection':
      args.workbook.setRangeProtection(args.op.protection)
      change.structuralInvalidation = true
      break
    case 'deleteRangeProtection':
      args.workbook.deleteRangeProtection(args.op.id)
      change.structuralInvalidation = true
      break
    case 'upsertCommentThread':
      args.workbook.setCommentThread(args.op.thread)
      change.structuralInvalidation = true
      break
    case 'deleteCommentThread':
      args.workbook.deleteCommentThread(args.op.sheetName, args.op.address)
      change.structuralInvalidation = true
      break
    case 'upsertNote':
      args.workbook.setNote(args.op.note)
      change.structuralInvalidation = true
      break
    case 'deleteNote':
      args.workbook.deleteNote(args.op.sheetName, args.op.address)
      change.structuralInvalidation = true
      break
    case 'upsertCellStyle': {
      const previousStyle = args.workbook.getCellStyle(args.op.style.id)
      const shouldInvalidateStyle = hasCellStylePresentationChanged(previousStyle, args.op.style)
      args.workbook.upsertCellStyle(args.op.style)
      if (shouldInvalidateStyle) {
        change.invalidatedRanges.push(...collectCellStyleInvalidatedRanges(args.workbook, args.op.style.id))
      }
      break
    }
    case 'upsertCellNumberFormat':
      args.workbook.upsertCellNumberFormat(args.op.format)
      break
    case 'setStyleRange':
      args.workbook.setStyleRange(args.op.range, args.op.styleId)
      if (args.source !== 'restore') {
        args.workbook.coalesceStyleRanges(args.op.range.sheetName)
      }
      change.invalidatedRanges.push(args.op.range)
      break
    case 'setFormatRange':
      args.workbook.setFormatRange(args.op.range, args.op.formatId)
      change.invalidatedRanges.push(args.op.range)
      break
    case 'upsertChart':
      args.workbook.setChart(args.op.chart)
      change.structuralInvalidation = true
      break
    case 'deleteChart':
      args.workbook.deleteChart(args.op.id)
      change.structuralInvalidation = true
      break
    case 'upsertImage':
      args.workbook.setImage(args.op.image)
      change.structuralInvalidation = true
      break
    case 'deleteImage':
      args.workbook.deleteImage(args.op.id)
      change.structuralInvalidation = true
      break
    case 'upsertShape':
      args.workbook.setShape(args.op.shape)
      change.structuralInvalidation = true
      break
    case 'deleteShape':
      args.workbook.deleteShape(args.op.id)
      change.structuralInvalidation = true
      break
    default:
      return assertNever(args.op)
  }

  args.setEntityVersionForOp(args.op, args.order)
  return change
}

function hasCellStylePresentationChanged(previousStyle: CellStyleRecord | undefined, nextStyle: CellStyleRecord): boolean {
  if (previousStyle === undefined) {
    return true
  }
  return cellStyleKey(normalizeCellStyleRecord(previousStyle)) !== cellStyleKey(normalizeCellStyleRecord(nextStyle))
}

function collectCellStyleInvalidatedRanges(workbook: OperationStructuralMetadataWorkbook, styleId: string): CellRangeRef[] {
  if (styleId === WORKBOOK_DEFAULT_STYLE_ID) {
    return [...workbook.sheetsByName.values()]
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => fullSheetRange(sheet.name))
  }

  const invalidated: CellRangeRef[] = []
  const seen = new Set<string>()
  for (const sheet of workbook.sheetsByName.values()) {
    for (const record of workbook.listStyleRanges(sheet.name)) {
      if (record.styleId !== styleId) {
        continue
      }
      const key = `${record.range.sheetName}\0${record.range.startAddress}\0${record.range.endAddress}`
      if (seen.has(key)) {
        continue
      }
      seen.add(key)
      invalidated.push({ ...record.range })
    }
  }
  return invalidated
}

function fullSheetRange(sheetName: string): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(0, 0),
    endAddress: formatAddress(MAX_ROWS - 1, MAX_COLS - 1),
  }
}
