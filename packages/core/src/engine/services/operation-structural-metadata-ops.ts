import type { CellRangeRef } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook'
import type { OpOrder } from '../../replica-state.js'
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
      args.workbook.setRowMetadata(args.op.sheetName, args.op.start, args.op.count, args.op.size, args.op.hidden)
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
    case 'upsertCellStyle':
      args.workbook.upsertCellStyle(args.op.style)
      break
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
