import type { CellRangeRef } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook'
import { assertNever, cellRange, rangesIntersect, throwProtectionBlocked } from './operation-change-helpers.js'

interface RangeProtectedRecord {
  readonly range: CellRangeRef
}

interface TableProtectedRecord {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
}

interface SourceProtectedRecord {
  readonly sheetName: string
  readonly address: string
  readonly source: CellRangeRef
}

interface PivotProtectedRecord {
  readonly sheetName: string
  readonly address: string
  readonly source?: CellRangeRef
}

interface AnchorProtectedRecord {
  readonly sheetName: string
  readonly address: string
}

export interface OperationProtectionAccess {
  readonly hasProtectionMetadataForSheet?: (sheetName: string) => boolean
  readonly getSheetProtection: (sheetName: string) => unknown
  readonly listRangeProtections: (sheetName: string) => readonly RangeProtectedRecord[]
  readonly getConditionalFormat: (id: string) => RangeProtectedRecord | undefined
  readonly getTable: (name: string) => TableProtectedRecord | undefined
  readonly getPivot: (sheetName: string, address: string) => PivotProtectedRecord | undefined
  readonly getChart: (id: string) => SourceProtectedRecord | undefined
  readonly getImage: (id: string) => AnchorProtectedRecord | undefined
  readonly getShape: (id: string) => AnchorProtectedRecord | undefined
}

export function sheetHasProtection(access: OperationProtectionAccess, sheetName: string): boolean {
  if (access.hasProtectionMetadataForSheet?.(sheetName) === false) {
    return false
  }
  return access.getSheetProtection(sheetName) !== undefined || access.listRangeProtections(sheetName).length > 0
}

export function rangeIsProtected(access: OperationProtectionAccess, range: CellRangeRef): boolean {
  if (access.hasProtectionMetadataForSheet?.(range.sheetName) === false) {
    return false
  }
  if (access.getSheetProtection(range.sheetName)) {
    return true
  }
  return access.listRangeProtections(range.sheetName).some((protection) => rangesIntersect(protection.range, range))
}

function rangeLabel(range: CellRangeRef): string {
  return `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

export function assertProtectionAllowsOp(access: OperationProtectionAccess, op: EngineOp): void {
  switch (op.kind) {
    case 'setSheetProtection':
    case 'clearSheetProtection':
    case 'upsertRangeProtection':
    case 'deleteRangeProtection':
    case 'upsertWorkbook':
    case 'setWorkbookMetadata':
    case 'setCalculationSettings':
    case 'setVolatileContext':
    case 'upsertDefinedName':
    case 'deleteDefinedName':
    case 'upsertCellStyle':
    case 'upsertCellNumberFormat':
      return
    case 'upsertSheet':
      return
    case 'renameSheet':
    case 'deleteSheet': {
      const sheetName = op.kind === 'renameSheet' ? op.oldName : op.name
      if (sheetHasProtection(access, sheetName)) {
        throwProtectionBlocked(`sheet ${sheetName} is protected`)
      }
      return
    }
    case 'insertRows':
    case 'deleteRows':
    case 'moveRows':
    case 'insertColumns':
    case 'deleteColumns':
    case 'moveColumns':
    case 'updateRowMetadata':
    case 'updateColumnMetadata':
    case 'setFreezePane':
    case 'clearFreezePane':
    case 'setConditionalFormatArtifacts':
    case 'clearConditionalFormatArtifacts':
      if (sheetHasProtection(access, op.sheetName)) {
        throwProtectionBlocked(`sheet ${op.sheetName} is protected`)
      }
      return
    case 'setFilter':
    case 'clearFilter':
    case 'setSort':
    case 'clearSort':
    case 'setStyleRange':
    case 'setFormatRange':
    case 'mergeCells':
    case 'unmergeCells':
      if (rangeIsProtected(access, op.range)) {
        throwProtectionBlocked(`range ${rangeLabel(op.range)} is protected`)
      }
      return
    case 'setDataValidation':
      if (rangeIsProtected(access, op.validation.range)) {
        throwProtectionBlocked(`range ${rangeLabel(op.validation.range)} is protected`)
      }
      return
    case 'clearDataValidation':
      if (rangeIsProtected(access, op.range)) {
        throwProtectionBlocked(`range ${rangeLabel(op.range)} is protected`)
      }
      return
    case 'upsertConditionalFormat':
      if (rangeIsProtected(access, op.format.range)) {
        throwProtectionBlocked(`range ${rangeLabel(op.format.range)} is protected`)
      }
      return
    case 'deleteConditionalFormat': {
      const existing = access.getConditionalFormat(op.id)
      if (existing && rangeIsProtected(access, existing.range)) {
        throwProtectionBlocked(`conditional format ${op.id} targets a protected range`)
      }
      return
    }
    case 'upsertCommentThread':
      if (rangeIsProtected(access, cellRange(op.thread.sheetName, op.thread.address))) {
        throwProtectionBlocked(`cell ${op.thread.sheetName}!${op.thread.address} is protected`)
      }
      return
    case 'deleteCommentThread':
      if (rangeIsProtected(access, cellRange(op.sheetName, op.address))) {
        throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
      }
      return
    case 'upsertNote':
      if (rangeIsProtected(access, cellRange(op.note.sheetName, op.note.address))) {
        throwProtectionBlocked(`cell ${op.note.sheetName}!${op.note.address} is protected`)
      }
      return
    case 'deleteNote':
      if (rangeIsProtected(access, cellRange(op.sheetName, op.address))) {
        throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
      }
      return
    case 'setCellValue':
    case 'setCellFormula':
    case 'setCellFormat':
    case 'clearCell':
      if (rangeIsProtected(access, cellRange(op.sheetName, op.address))) {
        throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
      }
      return
    case 'upsertTable':
      if (
        rangeIsProtected(access, {
          sheetName: op.table.sheetName,
          startAddress: op.table.startAddress,
          endAddress: op.table.endAddress,
        })
      ) {
        throwProtectionBlocked(`table ${op.table.name} overlaps a protected range`)
      }
      return
    case 'deleteTable': {
      const existing = access.getTable(op.name)
      if (
        existing &&
        rangeIsProtected(access, {
          sheetName: existing.sheetName,
          startAddress: existing.startAddress,
          endAddress: existing.endAddress,
        })
      ) {
        throwProtectionBlocked(`table ${op.name} overlaps a protected range`)
      }
      return
    }
    case 'upsertSpillRange':
    case 'deleteSpillRange':
      if (rangeIsProtected(access, cellRange(op.sheetName, op.address))) {
        throwProtectionBlocked(`cell ${op.sheetName}!${op.address} is protected`)
      }
      return
    case 'upsertPivotTable':
      if (
        sheetHasProtection(access, op.sheetName) ||
        rangeIsProtected(access, op.source) ||
        rangeIsProtected(access, cellRange(op.sheetName, op.address))
      ) {
        throwProtectionBlocked(`pivot ${op.name} touches protected workbook state`)
      }
      return
    case 'deletePivotTable': {
      const existing = access.getPivot(op.sheetName, op.address)
      if (
        existing &&
        (sheetHasProtection(access, existing.sheetName) ||
          (existing.source !== undefined && rangeIsProtected(access, existing.source)) ||
          rangeIsProtected(access, cellRange(existing.sheetName, existing.address)))
      ) {
        throwProtectionBlocked(`pivot at ${op.sheetName}!${op.address} touches protected workbook state`)
      }
      return
    }
    case 'upsertChart':
      if (
        sheetHasProtection(access, op.chart.sheetName) ||
        rangeIsProtected(access, op.chart.source) ||
        rangeIsProtected(access, cellRange(op.chart.sheetName, op.chart.address))
      ) {
        throwProtectionBlocked(`chart ${op.chart.id} touches protected workbook state`)
      }
      return
    case 'deleteChart': {
      const existing = access.getChart(op.id)
      if (
        existing &&
        (sheetHasProtection(access, existing.sheetName) ||
          rangeIsProtected(access, existing.source) ||
          rangeIsProtected(access, cellRange(existing.sheetName, existing.address)))
      ) {
        throwProtectionBlocked(`chart ${op.id} touches protected workbook state`)
      }
      return
    }
    case 'upsertImage':
      if (sheetHasProtection(access, op.image.sheetName) || rangeIsProtected(access, cellRange(op.image.sheetName, op.image.address))) {
        throwProtectionBlocked(`image ${op.image.id} touches protected workbook state`)
      }
      return
    case 'deleteImage': {
      const existing = access.getImage(op.id)
      if (
        existing &&
        (sheetHasProtection(access, existing.sheetName) || rangeIsProtected(access, cellRange(existing.sheetName, existing.address)))
      ) {
        throwProtectionBlocked(`image ${op.id} touches protected workbook state`)
      }
      return
    }
    case 'upsertShape':
      if (sheetHasProtection(access, op.shape.sheetName) || rangeIsProtected(access, cellRange(op.shape.sheetName, op.shape.address))) {
        throwProtectionBlocked(`shape ${op.shape.id} touches protected workbook state`)
      }
      return
    case 'deleteShape': {
      const existing = access.getShape(op.id)
      if (
        existing &&
        (sheetHasProtection(access, existing.sheetName) || rangeIsProtected(access, cellRange(existing.sheetName, existing.address)))
      ) {
        throwProtectionBlocked(`shape ${op.id} touches protected workbook state`)
      }
      return
    }
    default:
      assertNever(op)
      return
  }
}
