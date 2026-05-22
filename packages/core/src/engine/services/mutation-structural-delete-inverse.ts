import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { renameFormulaSheetReferences, rewriteFormulaForStructuralTransform } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import { CellFlags } from '../../cell-store.js'
import { mapStructuralAxisIndex, mapStructuralAxisInterval, structuralTransformForOp } from '../../engine-structural-utils.js'
import { sheetMetadataToOps } from '../../engine-snapshot-utils.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type {
  EngineRuntimeState,
  RuntimeFormula,
  RuntimeSheetRenameFormulaSourceTransform,
  RuntimeStructuralFormulaSourceTransform,
  TransactionRecord,
} from '../runtime-state.js'
import { getRuntimeFormulaSource } from '../runtime-formula-source.js'
import {
  dependencyTouchesStructuralDeleteSpan,
  structuralDeletedCellUndoRecordToOps,
  structuralFormulaUndoRecordToOp,
  type StructuralDeletedCellUndoRecord,
} from './mutation-structural-undo-records.js'
import { captureStructuralWorkbookMetadataOps, clearStructuralSheetMetadataOps } from './mutation-structural-metadata-ops.js'

interface MutationStructuralDeleteInverseRuntime {
  readonly state: Pick<EngineRuntimeState, 'formulas' | 'counters'> & {
    readonly workbook: WorkbookStore
  }
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly toCellStateOps: (sheetName: string, address: string, snapshot: CellSnapshot) => EngineOp[]
  readonly getFormulaFamilyStructuralSourceTransform?: (cellIndex: number) => RuntimeStructuralFormulaSourceTransform | undefined
  readonly hasFormulaFamilyStructuralSourceTransforms?: () => boolean
}

export interface MutationStructuralDeleteInverseHelpers {
  readonly captureFormulaCellStateForStructuralUndo: (sheetName: string, axis: 'row' | 'column', start: number, count: number) => EngineOp[]
  readonly buildStructuralDeleteInverseRecord: (
    op: Extract<EngineOp, { kind: 'deleteRows' | 'deleteColumns' }>,
    options?: { readonly includeStandaloneFormulaUndoOps?: boolean },
  ) => TransactionRecord
}

function captureRuntimeFormulaSource(
  args: MutationStructuralDeleteInverseRuntime,
  cellIndex: number,
  formula: RuntimeFormula,
  hasFamilyTransforms = args.hasFormulaFamilyStructuralSourceTransforms?.() === true,
): string {
  const familyTransform = hasFamilyTransforms ? args.getFormulaFamilyStructuralSourceTransform?.(cellIndex) : undefined
  return getRuntimeFormulaSource(formula, familyTransform)
}

interface CapturedRuntimeFormulaSource {
  readonly source: string
  readonly structuralSourceTransform?: RuntimeStructuralFormulaSourceTransform
  readonly inheritedStructuralSourceTransform?: RuntimeStructuralFormulaSourceTransform
  readonly sourceRenameTransforms?: readonly RuntimeSheetRenameFormulaSourceTransform[]
}

interface StructuralFormulaUndoCapture {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly source: CapturedRuntimeFormulaSource
}

function cloneStructuralSourceTransform(
  transform: RuntimeStructuralFormulaSourceTransform | undefined,
): RuntimeStructuralFormulaSourceTransform | undefined {
  if (!transform) {
    return undefined
  }
  return {
    ownerSheetName: transform.ownerSheetName,
    targetSheetName: transform.targetSheetName,
    transform: { ...transform.transform },
    preservesValue: transform.preservesValue,
  }
}

function captureRuntimeFormulaSourceHandle(
  args: MutationStructuralDeleteInverseRuntime,
  cellIndex: number,
  formula: RuntimeFormula,
  hasFamilyTransforms = args.hasFormulaFamilyStructuralSourceTransforms?.() === true,
): CapturedRuntimeFormulaSource {
  const ownStructuralSourceTransform = cloneStructuralSourceTransform(formula.structuralSourceTransform)
  const captured: CapturedRuntimeFormulaSource = {
    source: formula.source,
  }
  const withOwnTransform =
    ownStructuralSourceTransform === undefined ? captured : { ...captured, structuralSourceTransform: ownStructuralSourceTransform }
  const withRenameTransforms =
    formula.sourceRenameTransforms === undefined
      ? withOwnTransform
      : { ...withOwnTransform, sourceRenameTransforms: formula.sourceRenameTransforms.map((transform) => ({ ...transform })) }
  if (ownStructuralSourceTransform !== undefined) {
    return withRenameTransforms
  }
  const inheritedStructuralSourceTransform = hasFamilyTransforms
    ? cloneStructuralSourceTransform(args.getFormulaFamilyStructuralSourceTransform?.(cellIndex))
    : undefined
  return inheritedStructuralSourceTransform === undefined
    ? withRenameTransforms
    : { ...withRenameTransforms, inheritedStructuralSourceTransform }
}

function materializeRuntimeFormulaSource(captured: CapturedRuntimeFormulaSource): string {
  let source = captured.source
  captured.sourceRenameTransforms?.forEach((transform) => {
    source = renameFormulaSheetReferences(source, transform.oldSheetName, transform.newSheetName)
  })
  const deferred = captured.structuralSourceTransform ?? captured.inheritedStructuralSourceTransform
  if (!deferred) {
    return source
  }
  return rewriteFormulaForStructuralTransform(source, deferred.ownerSheetName, deferred.targetSheetName, deferred.transform)
}

function structuralFormulaUndoCaptureToOp(capture: StructuralFormulaUndoCapture): EngineOp {
  return structuralFormulaUndoRecordToOp({
    sheetName: capture.sheetName,
    row: capture.row,
    col: capture.col,
    formula: materializeRuntimeFormulaSource(capture.source),
  })
}

function canRestoreDirectAggregateThroughStructuralInverse(args: {
  readonly formula: RuntimeFormula
  readonly ownerSheetName: string
  readonly ownerAxisIndex: number | undefined
  readonly sheetName: string
  readonly axis: 'row' | 'column'
  readonly start: number
  readonly count: number
}): boolean {
  if (args.axis !== 'row' || args.ownerSheetName !== args.sheetName || args.ownerAxisIndex === undefined) {
    return false
  }
  const directAggregate = args.formula.directAggregate
  if (
    !directAggregate ||
    directAggregate.sheetName !== args.sheetName ||
    directAggregate.aggregateKind !== 'sum' ||
    args.formula.compiled.symbolicNames.length > 0 ||
    args.formula.compiled.symbolicTables.length > 0 ||
    args.formula.compiled.symbolicSpills.length > 0
  ) {
    return false
  }
  const transform = { kind: 'delete' as const, axis: 'row' as const, start: args.start, count: args.count }
  const deleteEnd = args.start + args.count - 1
  const deleteOverlapsAggregate = directAggregate.rowStart <= deleteEnd && directAggregate.rowEnd >= args.start
  const deleteTouchesAggregateBoundary = directAggregate.rowStart >= args.start || directAggregate.rowEnd <= deleteEnd
  if (deleteOverlapsAggregate && deleteTouchesAggregateBoundary) {
    return false
  }
  return (
    mapStructuralAxisIndex(args.ownerAxisIndex, transform) !== undefined &&
    mapStructuralAxisInterval(directAggregate.rowStart, directAggregate.rowEnd, transform) !== undefined
  )
}

function directFormulaRangeTouchesStructuralDeleteSpan(
  formula: RuntimeFormula,
  sheetName: string,
  axis: 'row' | 'column',
  start: number,
  count: number,
): boolean {
  const deleteEnd = start + count - 1
  const aggregate = formula.directAggregate
  if (aggregate?.sheetName === sheetName) {
    return axis === 'row'
      ? aggregate.rowStart <= deleteEnd && aggregate.rowEnd >= start
      : aggregate.col <= deleteEnd && aggregate.colEnd >= start
  }
  const criteria = formula.directCriteria
  if (!criteria) {
    return false
  }
  const directCriteriaRanges = [
    ...(criteria.aggregateRange ? [criteria.aggregateRange] : []),
    ...criteria.criteriaPairs.map((pair) => pair.range),
  ]
  return directCriteriaRanges.some((range) => {
    if (range.sheetName !== sheetName) {
      return false
    }
    return axis === 'row' ? range.rowStart <= deleteEnd && range.rowEnd >= start : range.col >= start && range.col <= deleteEnd
  })
}

export function createMutationStructuralDeleteInverseHelpers(
  args: MutationStructuralDeleteInverseRuntime,
): MutationStructuralDeleteInverseHelpers {
  const createSheetNameReader = (): ((sheetId: number) => string | undefined) => {
    const sheetNameById = new Map<number, string | undefined>()
    return (sheetId) => {
      if (sheetNameById.has(sheetId)) {
        return sheetNameById.get(sheetId)
      }
      const sheetName = args.state.workbook.getSheetNameById(sheetId)
      sheetNameById.set(sheetId, sheetName)
      return sheetName
    }
  }
  const createCellPositionReader = (): ((cellIndex: number) => { sheetId: number; row: number; col: number } | undefined) => {
    const sheetStructureVersionById = new Map<number, number | undefined>()
    const getSheetStructureVersionById = (sheetId: number): number | undefined => {
      if (sheetStructureVersionById.has(sheetId)) {
        return sheetStructureVersionById.get(sheetId)
      }
      const structureVersion = args.state.workbook.getSheetById(sheetId)?.structureVersion
      sheetStructureVersionById.set(sheetId, structureVersion)
      return structureVersion
    }
    return (cellIndex) => {
      const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      if (sheetId === undefined || sheetId === 0) {
        return undefined
      }
      if (getSheetStructureVersionById(sheetId) === 1) {
        const row = args.state.workbook.cellStore.rows[cellIndex]
        const col = args.state.workbook.cellStore.cols[cellIndex]
        return row === undefined || col === undefined ? undefined : { sheetId, row, col }
      }
      return args.state.workbook.getCellPosition(cellIndex)
    }
  }
  const captureFormulaCellRecordsForStructuralUndo = (
    sheetName: string,
    axis: 'row' | 'column',
    start: number,
    count: number,
    options: { readonly skipRedundantDirectAggregateCaptures?: boolean } = {},
  ): StructuralFormulaUndoCapture[] => {
    const captured: StructuralFormulaUndoCapture[] = []
    const hasFamilyTransforms = args.hasFormulaFamilyStructuralSourceTransforms?.() === true
    const getSheetNameById = createSheetNameReader()
    const readCellPosition = createCellPositionReader()
    args.state.formulas.forEach((formula, cellIndex) => {
      const ownerSheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      if (ownerSheetId === undefined) {
        return
      }
      const ownerSheetName = getSheetNameById(ownerSheetId)
      if (!ownerSheetName) {
        return
      }
      const ownerPosition = readCellPosition(cellIndex)
      const axisIndex = axis === 'row' ? ownerPosition?.row : ownerPosition?.col
      if (ownerSheetName === sheetName && axisIndex !== undefined && axisIndex >= start && axisIndex < start + count) {
        return
      }
      const ownerPositionAffected = ownerSheetName === sheetName && axisIndex !== undefined && axisIndex >= start + count
      if (
        options.skipRedundantDirectAggregateCaptures === true &&
        canRestoreDirectAggregateThroughStructuralInverse({
          formula,
          ownerSheetName,
          ownerAxisIndex: axisIndex,
          sheetName,
          axis,
          start,
          count,
        })
      ) {
        return
      }
      let dependencyPositionAffected = false
      if (!ownerPositionAffected) {
        if (formula.compiled.deps.length > 0) {
          addEngineCounter(args.state.counters, 'structuralUndoFormulaDependencyScans')
        }
        dependencyPositionAffected = formula.compiled.deps.some((dependency) =>
          dependencyTouchesStructuralDeleteSpan(dependency, ownerSheetName, sheetName, axis, start),
        )
      }
      const directFormulaRangeAffected = directFormulaRangeTouchesStructuralDeleteSpan(formula, sheetName, axis, start, count)
      const metadataSensitive =
        formula.compiled.symbolicNames.length > 0 ||
        formula.compiled.symbolicTables.length > 0 ||
        formula.compiled.symbolicSpills.length > 0
      if (!ownerPositionAffected && !dependencyPositionAffected && !directFormulaRangeAffected && !metadataSensitive) {
        return
      }
      captured.push({
        sheetName: ownerSheetName,
        row: ownerPosition?.row ?? 0,
        col: ownerPosition?.col ?? 0,
        source: captureRuntimeFormulaSourceHandle(args, cellIndex, formula, hasFamilyTransforms),
      })
    })
    return captured
  }

  const captureFormulaCellStateForStructuralUndo = (sheetName: string, axis: 'row' | 'column', start: number, count: number): EngineOp[] =>
    captureFormulaCellRecordsForStructuralUndo(sheetName, axis, start, count).map(structuralFormulaUndoCaptureToOp)

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
      const hasFamilyTransforms = args.hasFormulaFamilyStructuralSourceTransforms?.() === true
      return {
        kind: 'formula',
        sheetName,
        row,
        col,
        formula: captureRuntimeFormulaSource(args, cellIndex, formula, hasFamilyTransforms),
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
    const readCellPosition = createCellPositionReader()
    const axisIds = sheet.logicalAxisMap.snapshot(axis, start, count).map((entry) => entry.id)
    const records: StructuralDeletedCellUndoRecord[] = []
    sheet.logical.listResidentCellIndicesUnordered(axis, axisIds).forEach((cellIndex) => {
      const position = readCellPosition(cellIndex)
      if (!position || position.sheetId !== sheet.id) {
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
    formulaRecords: readonly StructuralFormulaUndoCapture[],
    potentialNewCells: number,
  ): TransactionRecord => {
    const record: { kind: 'ops'; ops: EngineOp[]; potentialNewCells?: number } = {
      kind: 'ops',
      get ops() {
        if (cachedOps === undefined) {
          if (deletedCellRecords.length > 0) {
            addEngineCounter(args.state.counters, 'structuralUndoCapturedCells', deletedCellRecords.length)
          }
          if (formulaRecords.length > 0) {
            addEngineCounter(args.state.counters, 'structuralUndoCapturedFormulas', formulaRecords.length)
          }
          cachedOps = [
            ...prefixOpsBeforeDeletedCells,
            ...deletedCellRecords.flatMap((deletedRecord) => structuralDeletedCellUndoRecordToOps(deletedRecord, args.toCellStateOps)),
            ...prefixOpsAfterDeletedCells,
            ...formulaRecords.map(structuralFormulaUndoCaptureToOp),
          ]
        }
        return cachedOps
      },
      potentialNewCells,
    }
    let cachedOps: EngineOp[] | undefined
    return record
  }

  const buildStructuralDeleteInverseRecord = (
    op: Extract<EngineOp, { kind: 'deleteRows' | 'deleteColumns' }>,
    options: { readonly includeStandaloneFormulaUndoOps?: boolean } = {},
  ): TransactionRecord => {
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
      captureFormulaCellRecordsForStructuralUndo(op.sheetName, axis, op.start, op.count, {
        skipRedundantDirectAggregateCaptures: options.includeStandaloneFormulaUndoOps !== true,
      }),
      1,
    )
  }

  return {
    captureFormulaCellStateForStructuralUndo,
    buildStructuralDeleteInverseRecord,
  }
}
