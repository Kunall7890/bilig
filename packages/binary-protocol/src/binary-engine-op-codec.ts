import type { WorkbookPivotValueSnapshot } from '@bilig/protocol'
import type { EngineOp, EngineOpBatch, WorkbookSortKey } from '@bilig/workbook-domain'
import { BinaryProtocolError, type BinaryReader, type BinaryWriter } from './binary-io.js'
import {
  assertNever,
  decodeAxisEntries,
  decodeCalculationMode,
  decodeCellNumberFormatRecord,
  decodeCellRangeRef,
  decodeCellStyleRecord,
  decodeCompatibilityMode,
  decodeDefinedNameValue,
  decodeLiteral,
  decodeNullableBoolean,
  decodeNullableNumber,
  decodeSortKey,
  encodeAxisEntries,
  encodeCalculationMode,
  encodeCellNumberFormatRecord,
  encodeCellRangeRef,
  encodeCellStyleRecord,
  encodeCompatibilityMode,
  encodeDefinedNameValue,
  encodeLiteral,
  encodeNullableBoolean,
  encodeNullableNumber,
  encodeSortKey,
} from './binary-value-codec.js'
import {
  decodeChart,
  decodeCommentThread,
  decodeConditionalFormat,
  decodeDataValidation,
  decodeImage,
  decodeNote,
  decodePivotValue,
  decodeRangeProtection,
  decodeShape,
  decodeSheetProtection,
  decodeTable,
  encodeChart,
  encodeCommentThread,
  encodeConditionalFormat,
  encodeDataValidation,
  encodeImage,
  encodeNote,
  encodePivotValue,
  encodeRangeProtection,
  encodeShape,
  encodeSheetProtection,
  encodeTable,
} from './binary-workbook-object-codec.js'

const OP_TAGS: Record<EngineOp['kind'], number> = {
  upsertWorkbook: 1,
  setWorkbookMetadata: 2,
  upsertSheet: 3,
  deleteSheet: 4,
  renameSheet: 37,
  updateRowMetadata: 5,
  updateColumnMetadata: 6,
  setFreezePane: 7,
  clearFreezePane: 8,
  setFilter: 9,
  clearFilter: 10,
  setSort: 11,
  clearSort: 12,
  setCellValue: 13,
  setCellFormula: 14,
  setCellFormat: 15,
  upsertCellStyle: 16,
  setStyleRange: 17,
  upsertCellNumberFormat: 35,
  setFormatRange: 36,
  clearCell: 18,
  upsertDefinedName: 19,
  deleteDefinedName: 20,
  upsertTable: 21,
  deleteTable: 22,
  upsertSpillRange: 23,
  deleteSpillRange: 24,
  upsertPivotTable: 25,
  deletePivotTable: 26,
  setCalculationSettings: 27,
  setVolatileContext: 28,
  insertRows: 29,
  deleteRows: 30,
  moveRows: 31,
  insertColumns: 32,
  deleteColumns: 33,
  moveColumns: 34,
  setSheetProtection: 46,
  clearSheetProtection: 47,
  setDataValidation: 38,
  clearDataValidation: 39,
  upsertCommentThread: 40,
  deleteCommentThread: 41,
  upsertNote: 42,
  deleteNote: 43,
  upsertConditionalFormat: 44,
  deleteConditionalFormat: 45,
  upsertRangeProtection: 48,
  deleteRangeProtection: 49,
  upsertChart: 50,
  deleteChart: 51,
  upsertImage: 52,
  deleteImage: 53,
  upsertShape: 54,
  deleteShape: 55,
  mergeCells: 56,
  unmergeCells: 57,
}

function encodeEngineOp(writer: BinaryWriter, op: EngineOp): void {
  writer.u8(OP_TAGS[op.kind])
  switch (op.kind) {
    case 'upsertWorkbook':
      writer.string(op.name)
      return
    case 'setWorkbookMetadata':
      writer.string(op.key)
      encodeLiteral(writer, op.value)
      return
    case 'setCalculationSettings':
      encodeCalculationMode(writer, op.settings.mode)
      encodeCompatibilityMode(writer, op.settings.compatibilityMode ?? 'excel-modern')
      return
    case 'setVolatileContext':
      writer.u32(op.context.recalcEpoch)
      return
    case 'upsertSheet':
      writer.string(op.name)
      writer.u32(op.order)
      return
    case 'deleteSheet':
      writer.string(op.name)
      return
    case 'renameSheet':
      writer.string(op.oldName)
      writer.string(op.newName)
      return
    case 'insertRows':
    case 'insertColumns':
      writer.string(op.sheetName)
      writer.u32(op.start)
      writer.u32(op.count)
      encodeAxisEntries(writer, op.entries)
      return
    case 'deleteRows':
    case 'deleteColumns':
      writer.string(op.sheetName)
      writer.u32(op.start)
      writer.u32(op.count)
      return
    case 'moveRows':
    case 'moveColumns':
      writer.string(op.sheetName)
      writer.u32(op.start)
      writer.u32(op.count)
      writer.u32(op.target)
      return
    case 'updateRowMetadata':
    case 'updateColumnMetadata':
      writer.string(op.sheetName)
      writer.u32(op.start)
      writer.u32(op.count)
      encodeNullableNumber(writer, op.size)
      encodeNullableBoolean(writer, op.hidden)
      return
    case 'setFreezePane':
      writer.string(op.sheetName)
      writer.u32(op.rows)
      writer.u32(op.cols)
      return
    case 'clearFreezePane':
      writer.string(op.sheetName)
      return
    case 'mergeCells':
    case 'unmergeCells':
      encodeCellRangeRef(writer, op.range)
      return
    case 'setSheetProtection':
      encodeSheetProtection(writer, op.protection)
      return
    case 'clearSheetProtection':
      writer.string(op.sheetName)
      return
    case 'setFilter':
    case 'clearFilter':
      writer.string(op.sheetName)
      encodeCellRangeRef(writer, op.range)
      return
    case 'setSort':
      writer.string(op.sheetName)
      encodeCellRangeRef(writer, op.range)
      writer.u32(op.keys.length)
      op.keys.forEach((key) => encodeSortKey(writer, key))
      return
    case 'clearSort':
      writer.string(op.sheetName)
      encodeCellRangeRef(writer, op.range)
      return
    case 'setDataValidation':
      encodeDataValidation(writer, op.validation)
      return
    case 'clearDataValidation':
      writer.string(op.sheetName)
      encodeCellRangeRef(writer, op.range)
      return
    case 'upsertConditionalFormat':
      encodeConditionalFormat(writer, op.format)
      return
    case 'deleteConditionalFormat':
      writer.string(op.id)
      writer.string(op.sheetName)
      return
    case 'upsertRangeProtection':
      encodeRangeProtection(writer, op.protection)
      return
    case 'deleteRangeProtection':
      writer.string(op.id)
      writer.string(op.sheetName)
      return
    case 'upsertCommentThread':
      encodeCommentThread(writer, op.thread)
      return
    case 'deleteCommentThread':
      writer.string(op.sheetName)
      writer.string(op.address)
      return
    case 'upsertNote':
      encodeNote(writer, op.note)
      return
    case 'deleteNote':
      writer.string(op.sheetName)
      writer.string(op.address)
      return
    case 'setCellValue':
      writer.string(op.sheetName)
      writer.string(op.address)
      encodeLiteral(writer, op.value)
      return
    case 'setCellFormula':
      writer.string(op.sheetName)
      writer.string(op.address)
      writer.string(op.formula)
      return
    case 'setCellFormat':
      writer.string(op.sheetName)
      writer.string(op.address)
      writer.bool(op.format !== null)
      if (op.format !== null) {
        writer.string(op.format)
      }
      return
    case 'upsertCellStyle':
      encodeCellStyleRecord(writer, op.style)
      return
    case 'setStyleRange':
      encodeCellRangeRef(writer, op.range)
      writer.string(op.styleId)
      return
    case 'upsertCellNumberFormat':
      encodeCellNumberFormatRecord(writer, op.format)
      return
    case 'setFormatRange':
      encodeCellRangeRef(writer, op.range)
      writer.string(op.formatId)
      return
    case 'clearCell':
      writer.string(op.sheetName)
      writer.string(op.address)
      return
    case 'upsertDefinedName':
      writer.string(op.name)
      encodeDefinedNameValue(writer, op.value)
      return
    case 'deleteDefinedName':
      writer.string(op.name)
      return
    case 'upsertTable':
      encodeTable(writer, op.table)
      return
    case 'deleteTable':
      writer.string(op.name)
      return
    case 'upsertSpillRange':
      writer.string(op.sheetName)
      writer.string(op.address)
      writer.u32(op.rows)
      writer.u32(op.cols)
      return
    case 'deleteSpillRange':
      writer.string(op.sheetName)
      writer.string(op.address)
      return
    case 'upsertPivotTable':
      writer.string(op.name)
      writer.string(op.sheetName)
      writer.string(op.address)
      encodeCellRangeRef(writer, op.source)
      writer.stringArray(op.groupBy)
      writer.u32(op.values.length)
      op.values.forEach((v) => encodePivotValue(writer, v))
      writer.u32(op.rows)
      writer.u32(op.cols)
      return
    case 'deletePivotTable':
      writer.string(op.sheetName)
      writer.string(op.address)
      return
    case 'upsertChart':
      encodeChart(writer, op.chart)
      return
    case 'deleteChart':
      writer.string(op.id)
      return
    case 'upsertImage':
      encodeImage(writer, op.image)
      return
    case 'deleteImage':
      writer.string(op.id)
      return
    case 'upsertShape':
      encodeShape(writer, op.shape)
      return
    case 'deleteShape':
      writer.string(op.id)
      return
    default:
      assertNever(op)
  }
}

function decodeEngineOp(reader: BinaryReader): EngineOp {
  switch (reader.u8()) {
    case 1:
      return { kind: 'upsertWorkbook', name: reader.string() }
    case 2:
      return { kind: 'setWorkbookMetadata', key: reader.string(), value: decodeLiteral(reader) }
    case 27:
      return {
        kind: 'setCalculationSettings',
        settings: {
          mode: decodeCalculationMode(reader),
          compatibilityMode: decodeCompatibilityMode(reader),
        },
      }
    case 28:
      return { kind: 'setVolatileContext', context: { recalcEpoch: reader.u32() } }
    case 3:
      return { kind: 'upsertSheet', name: reader.string(), order: reader.u32() }
    case 4:
      return { kind: 'deleteSheet', name: reader.string() }
    case 37:
      return { kind: 'renameSheet', oldName: reader.string(), newName: reader.string() }
    case 29:
      return {
        kind: 'insertRows',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        entries: decodeAxisEntries(reader),
      }
    case 30:
      return {
        kind: 'deleteRows',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
      }
    case 31:
      return {
        kind: 'moveRows',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        target: reader.u32(),
      }
    case 32:
      return {
        kind: 'insertColumns',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        entries: decodeAxisEntries(reader),
      }
    case 33:
      return {
        kind: 'deleteColumns',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
      }
    case 34:
      return {
        kind: 'moveColumns',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        target: reader.u32(),
      }
    case 5:
      return {
        kind: 'updateRowMetadata',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        size: decodeNullableNumber(reader),
        hidden: decodeNullableBoolean(reader),
      }
    case 6:
      return {
        kind: 'updateColumnMetadata',
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        size: decodeNullableNumber(reader),
        hidden: decodeNullableBoolean(reader),
      }
    case 7:
      return {
        kind: 'setFreezePane',
        sheetName: reader.string(),
        rows: reader.u32(),
        cols: reader.u32(),
      }
    case 8:
      return { kind: 'clearFreezePane', sheetName: reader.string() }
    case 56:
      return { kind: 'mergeCells', range: decodeCellRangeRef(reader) }
    case 57:
      return { kind: 'unmergeCells', range: decodeCellRangeRef(reader) }
    case 46:
      return {
        kind: 'setSheetProtection',
        protection: decodeSheetProtection(reader),
      }
    case 47:
      return { kind: 'clearSheetProtection', sheetName: reader.string() }
    case 9:
      return {
        kind: 'setFilter',
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      }
    case 10:
      return {
        kind: 'clearFilter',
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      }
    case 11:
      return {
        kind: 'setSort',
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
        keys: (() => {
          const count = reader.u32()
          const keys: WorkbookSortKey[] = []
          for (let index = 0; index < count; index += 1) {
            keys.push(decodeSortKey(reader))
          }
          return keys
        })(),
      }
    case 12:
      return {
        kind: 'clearSort',
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      }
    case 38:
      return {
        kind: 'setDataValidation',
        validation: decodeDataValidation(reader),
      }
    case 39:
      return {
        kind: 'clearDataValidation',
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      }
    case 44:
      return {
        kind: 'upsertConditionalFormat',
        format: decodeConditionalFormat(reader),
      }
    case 45:
      return {
        kind: 'deleteConditionalFormat',
        id: reader.string(),
        sheetName: reader.string(),
      }
    case 48:
      return {
        kind: 'upsertRangeProtection',
        protection: decodeRangeProtection(reader),
      }
    case 49:
      return {
        kind: 'deleteRangeProtection',
        id: reader.string(),
        sheetName: reader.string(),
      }
    case 40:
      return {
        kind: 'upsertCommentThread',
        thread: decodeCommentThread(reader),
      }
    case 41:
      return {
        kind: 'deleteCommentThread',
        sheetName: reader.string(),
        address: reader.string(),
      }
    case 42:
      return {
        kind: 'upsertNote',
        note: decodeNote(reader),
      }
    case 43:
      return {
        kind: 'deleteNote',
        sheetName: reader.string(),
        address: reader.string(),
      }
    case 13:
      return {
        kind: 'setCellValue',
        sheetName: reader.string(),
        address: reader.string(),
        value: decodeLiteral(reader),
      }
    case 14:
      return {
        kind: 'setCellFormula',
        sheetName: reader.string(),
        address: reader.string(),
        formula: reader.string(),
      }
    case 15: {
      const sheetName = reader.string()
      const address = reader.string()
      const hasFormat = reader.bool()
      return {
        kind: 'setCellFormat',
        sheetName,
        address,
        format: hasFormat ? reader.string() : null,
      }
    }
    case 16:
      return { kind: 'upsertCellStyle', style: decodeCellStyleRecord(reader) }
    case 17:
      return {
        kind: 'setStyleRange',
        range: decodeCellRangeRef(reader),
        styleId: reader.string(),
      }
    case 35:
      return { kind: 'upsertCellNumberFormat', format: decodeCellNumberFormatRecord(reader) }
    case 36:
      return {
        kind: 'setFormatRange',
        range: decodeCellRangeRef(reader),
        formatId: reader.string(),
      }
    case 18:
      return { kind: 'clearCell', sheetName: reader.string(), address: reader.string() }
    case 19:
      return {
        kind: 'upsertDefinedName',
        name: reader.string(),
        value: decodeDefinedNameValue(reader),
      }
    case 20:
      return { kind: 'deleteDefinedName', name: reader.string() }
    case 21:
      return { kind: 'upsertTable', table: decodeTable(reader) }
    case 22:
      return { kind: 'deleteTable', name: reader.string() }
    case 23:
      return {
        kind: 'upsertSpillRange',
        sheetName: reader.string(),
        address: reader.string(),
        rows: reader.u32(),
        cols: reader.u32(),
      }
    case 24:
      return {
        kind: 'deleteSpillRange',
        sheetName: reader.string(),
        address: reader.string(),
      }
    case 25:
      return {
        kind: 'upsertPivotTable',
        name: reader.string(),
        sheetName: reader.string(),
        address: reader.string(),
        source: decodeCellRangeRef(reader),
        groupBy: reader.stringArray(),
        values: (() => {
          const count = reader.u32()
          const values: WorkbookPivotValueSnapshot[] = []
          for (let i = 0; i < count; i++) values.push(decodePivotValue(reader))
          return values
        })(),
        rows: reader.u32(),
        cols: reader.u32(),
      }
    case 26:
      return {
        kind: 'deletePivotTable',
        sheetName: reader.string(),
        address: reader.string(),
      }
    case 50:
      return {
        kind: 'upsertChart',
        chart: decodeChart(reader),
      }
    case 51:
      return {
        kind: 'deleteChart',
        id: reader.string(),
      }
    case 52:
      return {
        kind: 'upsertImage',
        image: decodeImage(reader),
      }
    case 53:
      return {
        kind: 'deleteImage',
        id: reader.string(),
      }
    case 54:
      return {
        kind: 'upsertShape',
        shape: decodeShape(reader),
      }
    case 55:
      return {
        kind: 'deleteShape',
        id: reader.string(),
      }
    default:
      throw new BinaryProtocolError('Unknown engine op tag')
  }
}

export function encodeBatch(writer: BinaryWriter, batch: EngineOpBatch): void {
  writer.string(batch.id)
  writer.string(batch.replicaId)
  writer.u32(batch.clock.counter)
  writer.u32(batch.ops.length)
  batch.ops.forEach((op) => encodeEngineOp(writer, op))
}

export function decodeBatch(reader: BinaryReader): EngineOpBatch {
  const id = reader.string()
  const replicaId = reader.string()
  const counter = reader.u32()
  const opCount = reader.u32()
  const ops: EngineOp[] = []
  for (let index = 0; index < opCount; index += 1) {
    ops.push(decodeEngineOp(reader))
  }
  return {
    id,
    replicaId,
    clock: { counter },
    ops,
  }
}
