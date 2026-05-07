import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../../cell-store.js'
import { structuralTransformForOp } from '../../engine-structural-utils.js'
import { sheetMetadataToOps } from '../../engine-snapshot-utils.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeFormula, RuntimeStructuralFormulaSourceTransform, TransactionRecord } from '../runtime-state.js'
import { getRuntimeFormulaSource } from '../runtime-formula-source.js'
import {
  dependencyTouchesStructuralDeleteSpan,
  structuralDeletedCellUndoRecordToOps,
  structuralFormulaUndoRecordToOp,
  type StructuralDeletedCellUndoRecord,
  type StructuralFormulaUndoRecord,
} from './mutation-structural-undo-records.js'
import { captureStructuralWorkbookMetadataOps, clearStructuralSheetMetadataOps } from './mutation-structural-metadata-ops.js'

interface MutationStructuralDeleteInverseRuntime {
  readonly state: Pick<EngineRuntimeState, 'formulas' | 'counters'> & {
    readonly workbook: WorkbookStore
  }
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly toCellStateOps: (sheetName: string, address: string, snapshot: CellSnapshot) => EngineOp[]
  readonly getFormulaFamilyStructuralSourceTransform?: (cellIndex: number) => RuntimeStructuralFormulaSourceTransform | undefined
}

export interface MutationStructuralDeleteInverseHelpers {
  readonly captureFormulaCellStateForStructuralUndo: (sheetName: string, axis: 'row' | 'column', start: number, count: number) => EngineOp[]
  readonly buildStructuralDeleteInverseRecord: (op: Extract<EngineOp, { kind: 'deleteRows' | 'deleteColumns' }>) => TransactionRecord
}

function captureRuntimeFormulaSource(args: MutationStructuralDeleteInverseRuntime, cellIndex: number, formula: RuntimeFormula): string {
  return getRuntimeFormulaSource(formula, args.getFormulaFamilyStructuralSourceTransform?.(cellIndex))
}

export function createMutationStructuralDeleteInverseHelpers(
  args: MutationStructuralDeleteInverseRuntime,
): MutationStructuralDeleteInverseHelpers {
  const captureFormulaCellRecordsForStructuralUndo = (
    sheetName: string,
    axis: 'row' | 'column',
    start: number,
    count: number,
  ): StructuralFormulaUndoRecord[] => {
    const captured: StructuralFormulaUndoRecord[] = []
    args.state.formulas.forEach((formula, cellIndex) => {
      const ownerSheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      if (ownerSheetId === undefined) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(ownerSheetId)
      if (!ownerSheetName) {
        return
      }
      const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
      const axisIndex = axis === 'row' ? ownerPosition?.row : ownerPosition?.col
      if (ownerSheetName === sheetName && axisIndex !== undefined && axisIndex >= start && axisIndex < start + count) {
        return
      }
      const ownerPositionAffected = ownerSheetName === sheetName && axisIndex !== undefined && axisIndex >= start + count
      const dependencyPositionAffected =
        !ownerPositionAffected &&
        formula.compiled.deps.some((dependency) =>
          dependencyTouchesStructuralDeleteSpan(dependency, ownerSheetName, sheetName, axis, start),
        )
      const metadataSensitive =
        formula.compiled.symbolicNames.length > 0 ||
        formula.compiled.symbolicTables.length > 0 ||
        formula.compiled.symbolicSpills.length > 0
      if (!ownerPositionAffected && !dependencyPositionAffected && !metadataSensitive) {
        return
      }
      captured.push({
        sheetName: ownerSheetName,
        row: ownerPosition?.row ?? 0,
        col: ownerPosition?.col ?? 0,
        formula: captureRuntimeFormulaSource(args, cellIndex, formula),
      })
    })
    return captured
  }

  const captureFormulaCellStateForStructuralUndo = (sheetName: string, axis: 'row' | 'column', start: number, count: number): EngineOp[] =>
    captureFormulaCellRecordsForStructuralUndo(sheetName, axis, start, count).map(structuralFormulaUndoRecordToOp)

  const captureDeletedCellUndoRecord = (
    cellIndex: number,
    sheetName: string,
    row: number,
    col: number,
  ): StructuralDeletedCellUndoRecord | undefined => {
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    if ((flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
      return undefined
    }
    const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
    const formula = args.state.formulas.get(cellIndex)
    if (formula) {
      return {
        kind: 'formula',
        sheetName,
        row,
        col,
        formula: captureRuntimeFormulaSource(args, cellIndex, formula),
        ...(explicitFormat === undefined ? {} : { explicitFormat }),
      }
    }
    const tag: ValueTag = (args.state.workbook.cellStore.tags[cellIndex] ?? ValueTag.Empty) as ValueTag
    if (explicitFormat === undefined && (flags & CellFlags.AuthoredBlank) === 0 && (tag === ValueTag.Empty || tag === ValueTag.Error)) {
      return undefined
    }
    switch (tag) {
      case ValueTag.Number:
        return {
          kind: 'value',
          sheetName,
          row,
          col,
          value: args.state.workbook.cellStore.numbers[cellIndex] ?? 0,
          ...(explicitFormat === undefined ? {} : { explicitFormat }),
        }
      case ValueTag.Boolean:
        return {
          kind: 'value',
          sheetName,
          row,
          col,
          value: (args.state.workbook.cellStore.numbers[cellIndex] ?? 0) !== 0,
          ...(explicitFormat === undefined ? {} : { explicitFormat }),
        }
      case ValueTag.String: {
        const snapshot = args.getCellByIndex(cellIndex)
        if (explicitFormat === undefined) {
          delete snapshot.format
          delete snapshot.numberFormatId
        } else {
          snapshot.format = explicitFormat
        }
        return { kind: 'snapshot', sheetName, row, col, snapshot }
      }
      case ValueTag.Empty:
      case ValueTag.Error:
        return {
          kind: 'blank',
          sheetName,
          row,
          col,
          restoreExplicitBlank: (args.state.workbook.cellStore.versions[cellIndex] ?? 0) !== 0 || (flags & CellFlags.AuthoredBlank) !== 0,
          ...(explicitFormat === undefined ? {} : { explicitFormat }),
        }
    }
    return undefined
  }

  const captureDeletedCellUndoRecordsForStructuralUndo = (
    sheetName: string,
    axis: 'row' | 'column',
    start: number,
    count: number,
  ): StructuralDeletedCellUndoRecord[] => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return []
    }
    const axisIds = sheet.logicalAxisMap.snapshot(axis, start, count).map((entry) => entry.id)
    const records: StructuralDeletedCellUndoRecord[] = []
    sheet.logical.listResidentCellIndicesUnordered(axis, axisIds).forEach((cellIndex) => {
      const position = args.state.workbook.getCellPosition(cellIndex)
      if (!position) {
        return
      }
      const axisIndex = axis === 'row' ? position.row : position.col
      if (axisIndex < start || axisIndex >= start + count) {
        return
      }
      const record = captureDeletedCellUndoRecord(cellIndex, sheetName, position.row, position.col)
      if (record) {
        records.push(record)
      }
    })
    return records.toSorted((left, right) => left.row - right.row || left.col - right.col)
  }

  const createLazyStructuralDeleteInverseRecord = (
    prefixOpsBeforeDeletedCells: readonly EngineOp[],
    deletedCellRecords: readonly StructuralDeletedCellUndoRecord[],
    prefixOpsAfterDeletedCells: readonly EngineOp[],
    formulaRecords: readonly StructuralFormulaUndoRecord[],
    potentialNewCells: number,
  ): TransactionRecord => {
    const record: { kind: 'ops'; ops: EngineOp[]; potentialNewCells?: number } = {
      kind: 'ops',
      get ops() {
        if (cachedOps === undefined) {
          if (deletedCellRecords.length > 0) {
            addEngineCounter(args.state.counters, 'structuralUndoCapturedCells', deletedCellRecords.length)
          }
          cachedOps = [
            ...prefixOpsBeforeDeletedCells,
            ...deletedCellRecords.flatMap((deletedRecord) => structuralDeletedCellUndoRecordToOps(deletedRecord, args.toCellStateOps)),
            ...prefixOpsAfterDeletedCells,
            ...formulaRecords.map(structuralFormulaUndoRecordToOp),
          ]
        }
        return cachedOps
      },
      potentialNewCells,
    }
    let cachedOps: EngineOp[] | undefined
    return record
  }

  const buildStructuralDeleteInverseRecord = (op: Extract<EngineOp, { kind: 'deleteRows' | 'deleteColumns' }>): TransactionRecord => {
    const axis = op.kind === 'deleteRows' ? 'row' : 'column'
    const entries =
      axis === 'row'
        ? args.state.workbook.snapshotRowAxisEntries(op.sheetName, op.start, op.count)
        : args.state.workbook.snapshotColumnAxisEntries(op.sheetName, op.start, op.count)
    const transform = structuralTransformForOp(op)
    const prefixOpsBeforeDeletedCells: EngineOp[] = [
      ...clearStructuralSheetMetadataOps(args.state.workbook, op.sheetName, transform),
      axis === 'row'
        ? {
            kind: 'insertRows',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          }
        : {
            kind: 'insertColumns',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
      ...sheetMetadataToOps(args.state.workbook, op.sheetName, { includeAxisEntries: false }),
    ]
    const deletedCellRecords = captureDeletedCellUndoRecordsForStructuralUndo(op.sheetName, axis, op.start, op.count)
    const prefixOpsAfterDeletedCells = captureStructuralWorkbookMetadataOps(args.state.workbook)
    return createLazyStructuralDeleteInverseRecord(
      prefixOpsBeforeDeletedCells,
      deletedCellRecords,
      prefixOpsAfterDeletedCells,
      captureFormulaCellRecordsForStructuralUndo(op.sheetName, axis, op.start, op.count),
      1,
    )
  }

  return {
    captureFormulaCellStateForStructuralUndo,
    buildStructuralDeleteInverseRecord,
  }
}
