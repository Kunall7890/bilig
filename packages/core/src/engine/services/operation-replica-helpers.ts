import type { EngineOp } from '@bilig/workbook'
import { compareOpOrder, type OpOrder } from '../../replica-state.js'
import { normalizeDefinedName, pivotKey } from '../../workbook-store.js'
import { assertNever } from './operation-change-helpers.js'

export interface VersionStore {
  get(key: string): OpOrder | undefined
  set(key: string, value: OpOrder): void
}

export const noopVersionStore: VersionStore = {
  get() {
    return undefined
  },
  set() {
    return
  },
}

export interface OperationReplicaVersionWriter {
  readonly stores: {
    readonly entityVersions: VersionStore
    readonly sheetDeleteVersions: VersionStore
  }
  readonly setEntityVersionForOp: (op: EngineOp, order: OpOrder) => void
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly setSheetDeleteVersion: (sheetName: string, order: OpOrder) => void
}

export function createOperationReplicaVersionWriter(input: {
  readonly trackReplicaVersions: boolean
  readonly entityVersions: VersionStore
  readonly sheetDeleteVersions: VersionStore
}): OperationReplicaVersionWriter {
  const entityVersions = input.trackReplicaVersions ? input.entityVersions : noopVersionStore
  const sheetDeleteVersions = input.trackReplicaVersions ? input.sheetDeleteVersions : noopVersionStore
  return {
    stores: { entityVersions, sheetDeleteVersions },
    setEntityVersionForOp(op, order) {
      if (!input.trackReplicaVersions) {
        return
      }
      entityVersions.set(entityKeyForOp(op), order)
    },
    setCellEntityVersion(sheetName, address, order) {
      if (!input.trackReplicaVersions) {
        return
      }
      entityVersions.set(`cell:${sheetName}!${address}`, order)
    },
    setSheetDeleteVersion(sheetName, order) {
      if (!input.trackReplicaVersions) {
        return
      }
      sheetDeleteVersions.set(sheetName, order)
    },
  }
}

export function entityKeyForOp(op: EngineOp): string {
  switch (op.kind) {
    case 'upsertWorkbook':
      return 'workbook'
    case 'setWorkbookMetadata':
      return `workbook-meta:${op.key}`
    case 'setCalculationSettings':
      return 'workbook-calc'
    case 'setVolatileContext':
      return 'workbook-volatile'
    case 'upsertSheet':
    case 'deleteSheet':
      return `sheet:${op.name}`
    case 'renameSheet':
      return `sheet:${op.oldName}`
    case 'insertRows':
    case 'deleteRows':
    case 'moveRows':
      return `row-structure:${op.sheetName}`
    case 'insertColumns':
    case 'deleteColumns':
    case 'moveColumns':
      return `column-structure:${op.sheetName}`
    case 'updateRowMetadata':
      return `row-meta:${op.sheetName}:${op.start}:${op.count}`
    case 'updateColumnMetadata':
      return `column-meta:${op.sheetName}:${op.start}:${op.count}`
    case 'setFreezePane':
    case 'clearFreezePane':
      return `freeze:${op.sheetName}`
    case 'mergeCells':
    case 'unmergeCells':
      return `merge:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
    case 'setSheetProtection':
    case 'clearSheetProtection':
      return `sheet-protection:${op.kind === 'setSheetProtection' ? op.protection.sheetName : op.sheetName}`
    case 'setFilter':
    case 'clearFilter':
      return `filter:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
    case 'setSort':
    case 'clearSort':
      return `sort:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
    case 'setDataValidation':
      return `validation:${op.validation.range.sheetName}:${op.validation.range.startAddress}:${op.validation.range.endAddress}`
    case 'clearDataValidation':
      return `validation:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
    case 'upsertConditionalFormat':
      return `conditional-format:${op.format.id}`
    case 'deleteConditionalFormat':
      return `conditional-format:${op.id}`
    case 'upsertRangeProtection':
      return `range-protection:${op.protection.id}`
    case 'deleteRangeProtection':
      return `range-protection:${op.id}`
    case 'upsertCommentThread':
      return `comment:${op.thread.sheetName}!${op.thread.address}`
    case 'deleteCommentThread':
      return `comment:${op.sheetName}!${op.address}`
    case 'upsertNote':
      return `note:${op.note.sheetName}!${op.note.address}`
    case 'deleteNote':
      return `note:${op.sheetName}!${op.address}`
    case 'setCellFormat':
      return `format:${op.sheetName}!${op.address}`
    case 'upsertCellStyle':
      return `style:${op.style.id}`
    case 'upsertCellNumberFormat':
      return `number-format:${op.format.id}`
    case 'setStyleRange':
      return `style-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
    case 'setFormatRange':
      return `format-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`
    case 'setCellValue':
    case 'setCellFormula':
    case 'clearCell':
      return `cell:${op.sheetName}!${op.address}`
    case 'upsertDefinedName':
    case 'deleteDefinedName':
      return `defined-name:${normalizeDefinedName(op.name)}`
    case 'upsertTable':
      return `table:${normalizeDefinedName(op.table.name)}`
    case 'deleteTable':
      return `table:${normalizeDefinedName(op.name)}`
    case 'upsertSpillRange':
    case 'deleteSpillRange':
      return `spill:${op.sheetName}!${op.address}`
    case 'upsertPivotTable':
    case 'deletePivotTable':
      return `pivot:${pivotKey(op.sheetName, op.address)}`
    case 'upsertChart':
      return `chart:${op.chart.id.trim().toUpperCase()}`
    case 'deleteChart':
      return `chart:${op.id.trim().toUpperCase()}`
    case 'upsertImage':
      return `image:${op.image.id.trim().toUpperCase()}`
    case 'deleteImage':
      return `image:${op.id.trim().toUpperCase()}`
    case 'upsertShape':
      return `shape:${op.shape.id.trim().toUpperCase()}`
    case 'deleteShape':
      return `shape:${op.id.trim().toUpperCase()}`
    default:
      return assertNever(op)
  }
}

export function sheetDeleteBarrierForOp(op: EngineOp, sheetDeleteVersions: VersionStore): OpOrder | undefined {
  switch (op.kind) {
    case 'upsertWorkbook':
    case 'setWorkbookMetadata':
    case 'setCalculationSettings':
    case 'setVolatileContext':
    case 'deleteSheet':
    case 'upsertDefinedName':
    case 'deleteDefinedName':
    case 'upsertTable':
    case 'deleteTable':
      return undefined
    case 'updateRowMetadata':
    case 'updateColumnMetadata':
    case 'insertRows':
    case 'deleteRows':
    case 'moveRows':
    case 'insertColumns':
    case 'deleteColumns':
    case 'moveColumns':
    case 'setFreezePane':
    case 'clearFreezePane':
    case 'clearSheetProtection':
    case 'setFilter':
    case 'clearFilter':
    case 'setSort':
    case 'clearSort':
    case 'clearDataValidation':
    case 'deleteConditionalFormat':
    case 'deleteRangeProtection':
    case 'deleteCommentThread':
    case 'deleteNote':
    case 'setCellFormat':
    case 'setCellValue':
    case 'setCellFormula':
    case 'clearCell':
    case 'upsertSpillRange':
    case 'deleteSpillRange':
    case 'deletePivotTable':
      return sheetDeleteVersions.get(op.sheetName)
    case 'mergeCells':
    case 'unmergeCells':
    case 'setStyleRange':
    case 'setFormatRange':
      return sheetDeleteVersions.get(op.range.sheetName)
    case 'upsertCellNumberFormat':
    case 'upsertCellStyle':
      return undefined
    case 'upsertSheet':
      return sheetDeleteVersions.get(op.name)
    case 'renameSheet':
      return sheetDeleteVersions.get(op.oldName)
    case 'setDataValidation':
      return sheetDeleteVersions.get(op.validation.range.sheetName)
    case 'setSheetProtection':
      return sheetDeleteVersions.get(op.protection.sheetName)
    case 'upsertConditionalFormat':
      return sheetDeleteVersions.get(op.format.range.sheetName)
    case 'upsertRangeProtection':
      return sheetDeleteVersions.get(op.protection.range.sheetName)
    case 'upsertCommentThread':
      return sheetDeleteVersions.get(op.thread.sheetName)
    case 'upsertNote':
      return sheetDeleteVersions.get(op.note.sheetName)
    case 'upsertPivotTable':
      return sheetDeleteVersions.get(op.sheetName) ?? sheetDeleteVersions.get(op.source.sheetName)
    case 'upsertChart':
      return sheetDeleteVersions.get(op.chart.sheetName) ?? sheetDeleteVersions.get(op.chart.source.sheetName)
    case 'deleteChart':
      return undefined
    case 'upsertImage':
      return sheetDeleteVersions.get(op.image.sheetName)
    case 'deleteImage':
      return undefined
    case 'upsertShape':
      return sheetDeleteVersions.get(op.shape.sheetName)
    case 'deleteShape':
      return undefined
    default:
      return assertNever(op)
  }
}

export function shouldApplyOp(
  op: EngineOp,
  order: OpOrder,
  stores: {
    readonly entityVersions: VersionStore
    readonly sheetDeleteVersions: VersionStore
  },
): boolean {
  const sheetDeleteOrder = sheetDeleteBarrierForOp(op, stores.sheetDeleteVersions)
  if (sheetDeleteOrder && compareOpOrder(order, sheetDeleteOrder) <= 0) {
    return false
  }
  const existingOrder = stores.entityVersions.get(entityKeyForOp(op))
  if (existingOrder && compareOpOrder(order, existingOrder) <= 0) {
    return false
  }
  return true
}
