import { Effect } from 'effect'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import type { CellRangeRef, CellSnapshot } from '@bilig/protocol'
import { formatAddress, parseCellAddress, rewriteFormulaForStructuralTransform } from '@bilig/formula'
import { createBatch } from '../../replica-state.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { mapStructuralAxisIndex, structuralTransformForOp } from '../../engine-structural-utils.js'
import {
  cellMutationRefToEngineOp,
  cloneCellMutationRef,
  countPotentialNewCellsForMutationRefs,
  type EngineCellMutationRef,
  type EngineExistingLiteralCellMutationRef,
  type EngineExistingNumericCellMutationRef,
  type EngineExistingNumericCellMutationResult,
} from '../../cell-mutations-at.js'
import type {
  EngineRuntimeState,
  PreparedCellAddress,
  RuntimeStructuralFormulaSourceTransform,
  TransactionRecord,
} from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import { tryBuildFastMutationHistory, type FastMutationHistoryResult } from './mutation-history-fast-path.js'
import {
  cloneTransactionRecordOps,
  createLazyCellMutationTransactionRecord,
  createLazySingleOpTransactionRecord,
  createOpsTransactionRecord,
  createSingleExistingLiteralCellMutationTransactionRecord,
  createSingleExistingNumericCellMutationTransactionRecord,
  singleExistingLiteralCellMutationRecordToRef,
  singleExistingNumericCellMutationRecordToRef,
  transactionRecordOps,
} from './mutation-transaction-records.js'
import { normalizeRenderCommitOps } from './mutation-render-commit-normalizer.js'
import { isMutationStructuralInsertOp } from './mutation-cell-content-helpers.js'
import { createMutationCellRestoreHistoryHelpers, tryMutationCellRefsFromOps } from './mutation-cell-restore-history.js'
import { createMutationStructuralDeleteInverseHelpers } from './mutation-structural-delete-inverse.js'
import type { CellStateRestoreOptions } from './cell-state-service.js'
import type { EngineMutationService } from './mutation-service-types.js'
import { tryExecuteMutationRenderCommitFastPath } from './mutation-render-commit-fast-path.js'
import { createMutationRangeOperations } from './mutation-range-operations.js'
import { createMutationCoreInverseOps } from './mutation-core-inverse-ops.js'

export type { EngineMutationService } from './mutation-service-types.js'

export function createEngineMutationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | 'replicaState'
    | 'batchListeners'
    | 'formulas'
    | 'undoStack'
    | 'redoStack'
    | 'counters'
    | 'trackReplicaVersions'
    | 'getSyncClientConnection'
    | 'getTransactionReplayDepth'
    | 'setTransactionReplayDepth'
  > & {
    readonly workbook: WorkbookStore
  }
  readonly captureSheetCellState: (sheetName: string) => EngineOp[]
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly captureColumnRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly captureStoredCellOps: (cellIndex: number, sheetName: string, address: string) => EngineOp[]
  readonly restoreCellOps: (sheetName: string, address: string) => EngineOp[]
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly getFormulaFamilyStructuralSourceTransform?: (cellIndex: number) => RuntimeStructuralFormulaSourceTransform | undefined
  readonly hasFormulaFamilyStructuralSourceTransforms?: () => boolean
  readonly readRangeCells: (range: CellRangeRef) => CellSnapshot[][]
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
    options?: CellStateRestoreOptions,
  ) => EngineOp[]
  readonly applyBatchNow: (
    batch: EngineOpBatch,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
    options?: { readonly emitTracked?: boolean },
  ) => void
  readonly applyLocalSingleStructuralAxisOpWithoutBatchNow?: (
    op: Extract<EngineOp, { kind: 'insertRows' | 'insertColumns' }>,
    options?: { readonly emitTracked?: boolean },
  ) => boolean
  readonly applyCellMutationsAtBatchNow: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => void
  readonly applyExistingNumericCellMutationAtNow?: (
    request: EngineExistingNumericCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly applyExistingLiteralCellMutationAtNow?: (
    request: EngineExistingLiteralCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly settleImportedFormulas?: () => void
  readonly hasExternallyVisibleLocalMutationObservers?: () => boolean
}): EngineMutationService {
  const emptyBatchOps: EngineOp[] = []
  const shouldCreateLocalBatch = (): boolean =>
    args.state.trackReplicaVersions ||
    (args.state.batchListeners?.size ?? 0) > 0 ||
    (args.state.getSyncClientConnection?.() ?? null) !== null
  const hasExternallyVisibleBatchRequirement = (): boolean =>
    args.hasExternallyVisibleLocalMutationObservers?.() ??
    ((args.state.batchListeners?.size ?? 0) > 0 || (args.state.getSyncClientConnection?.() ?? null) !== null)

  const {
    restoreCellOpFromRef,
    tryRestoreSimpleCellOpFromStore,
    createLazyInverseCellMutationRecord,
    tryCreateSingleExistingNumericInverseCellMutationRecord,
    tryCreateExistingNumericInverseCellMutationRecord,
    buildFastMutationHistoryFromRefs,
  } = createMutationCellRestoreHistoryHelpers({
    workbook: args.state.workbook,
    formulas: args.state.formulas,
    getCellByIndex: args.getCellByIndex,
    ...(args.getFormulaFamilyStructuralSourceTransform
      ? { getFormulaFamilyStructuralSourceTransform: args.getFormulaFamilyStructuralSourceTransform }
      : {}),
    ...(args.hasFormulaFamilyStructuralSourceTransforms
      ? { hasFormulaFamilyStructuralSourceTransforms: args.hasFormulaFamilyStructuralSourceTransforms }
      : {}),
  })
  const { captureFormulaCellStateForStructuralMoveUndo, captureFormulaCellStateForStructuralUndo, buildStructuralDeleteInverseRecord } =
    createMutationStructuralDeleteInverseHelpers({
      state: args.state,
      getCellByIndex: args.getCellByIndex,
      toCellStateOps: (sheetName, address, snapshot, options) =>
        args.toCellStateOps(sheetName, address, snapshot, undefined, undefined, options),
      ...(args.getFormulaFamilyStructuralSourceTransform
        ? { getFormulaFamilyStructuralSourceTransform: args.getFormulaFamilyStructuralSourceTransform }
        : {}),
      ...(args.hasFormulaFamilyStructuralSourceTransforms
        ? { hasFormulaFamilyStructuralSourceTransforms: args.hasFormulaFamilyStructuralSourceTransforms }
        : {}),
    })
  const { buildInverseOps } = createMutationCoreInverseOps({
    workbook: args.state.workbook,
    captureSheetCellState: args.captureSheetCellState,
    captureRowRangeCellState: args.captureRowRangeCellState,
    captureColumnRangeCellState: args.captureColumnRangeCellState,
    restoreCellOps: args.restoreCellOps,
    captureFormulaCellStateForStructuralMoveUndo,
    captureFormulaCellStateForStructuralUndo,
  })

  const tryBuildSingleCellOpHistoryWithoutSnapshot = (
    ops: readonly EngineOp[],
    potentialNewCells: number | undefined,
    includeUndoOps: boolean,
    cloneForwardOp: boolean,
  ): FastMutationHistoryResult | null => {
    if (ops.length !== 1) {
      return null
    }
    const op = ops[0]!
    if (op.kind !== 'setCellValue' && op.kind !== 'setCellFormula' && op.kind !== 'clearCell') {
      return null
    }
    const inverseOp = tryRestoreSimpleCellOpFromStore(op.sheetName, op.address)
    if (inverseOp === null) {
      return null
    }
    const forwardOp = cloneForwardOp ? structuredClone(op) : op
    return {
      forward: createLazySingleOpTransactionRecord(forwardOp, potentialNewCells),
      inverse: createLazySingleOpTransactionRecord(inverseOp, 1),
      undoOps: includeUndoOps ? [structuredClone(inverseOp)] : null,
    }
  }

  const tryCellMutationRefsFromOps = (ops: readonly EngineOp[]): EngineCellMutationRef[] | null =>
    tryMutationCellRefsFromOps(args.state.workbook, ops)

  const buildInverseRecord = (ops: readonly EngineOp[], includeStandaloneFormulaUndoOps: boolean): TransactionRecord => {
    if (ops.length === 1 && (ops[0]?.kind === 'deleteRows' || ops[0]?.kind === 'deleteColumns')) {
      return buildStructuralDeleteInverseRecord(ops[0], { includeStandaloneFormulaUndoOps })
    }
    return { kind: 'ops', ops: buildInverseOps(ops), potentialNewCells: ops.length }
  }

  const canonicalizeForwardOps = (ops: readonly EngineOp[]): EngineOp[] =>
    ops.flatMap((op) => {
      if (op.kind === 'insertRows') {
        const entries = op.entries?.map((entry) => ({ ...entry })) ?? args.state.workbook.createRowAxisEntries(op.start, op.count)
        return [{ ...op, entries }]
      }

      if (op.kind === 'insertColumns') {
        const entries = op.entries?.map((entry) => ({ ...entry })) ?? args.state.workbook.createColumnAxisEntries(op.start, op.count)
        return [{ ...op, entries }]
      }

      if (op.kind === 'moveRows') {
        return [structuredClone(op), ...captureFormulaCellStateForStructuralMoveUndo(op.sheetName, 'row', op.start, op.count, op.target)]
      }

      if (op.kind === 'moveColumns') {
        return [structuredClone(op), ...captureFormulaCellStateForStructuralMoveUndo(op.sheetName, 'column', op.start, op.count, op.target)]
      }

      return [structuredClone(op)]
    })

  const structuralFormulaForwardCorrections = (
    op: Extract<EngineOp, { kind: 'insertRows' | 'insertColumns' }>,
    beforeFormulaOps: readonly EngineOp[],
    afterFormulaOps: readonly EngineOp[],
  ): EngineOp[] => {
    if (beforeFormulaOps.length === 0 || afterFormulaOps.length === 0) {
      return emptyBatchOps
    }
    const transform = structuralTransformForOp(op)
    const expectedFormulaByCell = new Map<string, string>()
    beforeFormulaOps.forEach((beforeOp) => {
      if (beforeOp.kind !== 'setCellFormula') {
        return
      }
      const owner = parseCellAddress(beforeOp.address, beforeOp.sheetName)
      const ownerRow =
        beforeOp.sheetName === op.sheetName && transform.axis === 'row' ? mapStructuralAxisIndex(owner.row, transform) : owner.row
      const ownerCol =
        beforeOp.sheetName === op.sheetName && transform.axis === 'column' ? mapStructuralAxisIndex(owner.col, transform) : owner.col
      if (ownerRow === undefined || ownerCol === undefined) {
        return
      }
      const address = formatAddress(ownerRow, ownerCol)
      const formula = rewriteFormulaForStructuralTransform(beforeOp.formula, beforeOp.sheetName, op.sheetName, transform)
      expectedFormulaByCell.set(`${beforeOp.sheetName}\0${address}`, formula)
    })
    return afterFormulaOps.flatMap((afterOp) => {
      if (afterOp.kind !== 'setCellFormula') {
        return emptyBatchOps
      }
      const expected = expectedFormulaByCell.get(`${afterOp.sheetName}\0${afterOp.address}`)
      return expected === afterOp.formula ? emptyBatchOps : [afterOp]
    })
  }

  const executeTransactionNow = (
    record: TransactionRecord,
    source: 'local' | 'restore' | 'undo' | 'redo',
    options: { readonly emitTracked?: boolean } = {},
  ): void => {
    if ((record.kind === 'ops' && record.ops.length === 0) || (record.kind === 'cell-mutations' && record.refs.length === 0)) {
      return
    }
    if (record.kind === 'single-existing-numeric-cell-mutation') {
      const ref = singleExistingNumericCellMutationRecordToRef(record)
      const refs = [ref]
      const batch = shouldCreateLocalBatch()
        ? createBatch(args.state.replicaState, [cellMutationRefToEngineOp(args.state.workbook, ref)])
        : null
      args.applyCellMutationsAtBatchNow(refs, batch, source, record.potentialNewCells)
      return
    }
    if (record.kind === 'single-existing-literal-cell-mutation') {
      const ref = singleExistingLiteralCellMutationRecordToRef(record)
      const refs = [ref]
      const batch = shouldCreateLocalBatch()
        ? createBatch(args.state.replicaState, [cellMutationRefToEngineOp(args.state.workbook, ref)])
        : null
      args.applyCellMutationsAtBatchNow(refs, batch, source, record.potentialNewCells)
      return
    }
    if (record.kind === 'cell-mutations') {
      const batch = shouldCreateLocalBatch()
        ? createBatch(
            args.state.replicaState,
            record.refs.map((ref) => cellMutationRefToEngineOp(args.state.workbook, ref)),
          )
        : null
      args.applyCellMutationsAtBatchNow(record.refs, batch, source, record.potentialNewCells)
      return
    }
    if (
      record.kind === 'single-op' &&
      source === 'local' &&
      options.emitTracked === false &&
      !shouldCreateLocalBatch() &&
      isMutationStructuralInsertOp(record.op) &&
      args.applyLocalSingleStructuralAxisOpWithoutBatchNow?.(record.op, options)
    ) {
      return
    }
    const batch = createBatch(args.state.replicaState, record.kind === 'single-op' ? [record.op] : record.ops)
    args.applyBatchNow(
      batch,
      source,
      record.potentialNewCells,
      record.kind === 'ops' ? record.preparedCellAddressesByOpIndex : undefined,
      options,
    )
  }

  const executeLocalNowWithCustomApply = (
    ops: EngineOp[],
    potentialNewCells: number | undefined,
    applyForward: (forward: TransactionRecord) => void,
    options: {
      returnUndoOps?: boolean
      reuseForwardOps?: boolean
      preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[]
      emitTracked?: boolean
    } = {},
  ): readonly EngineOp[] | null => {
    if (ops.length === 0) {
      return null
    }
    if (
      options.returnUndoOps === false &&
      ops.length === 1 &&
      options.preparedCellAddressesByOpIndex === undefined &&
      isMutationStructuralInsertOp(ops[0]!)
    ) {
      const op = ops[0]
      const appliedForwardOp = canonicalizeForwardOps([op])[0]!
      const inverseOps = buildInverseOps([op])
      const inverse =
        inverseOps.length === 1
          ? createLazySingleOpTransactionRecord(inverseOps[0]!, ops.length)
          : createOpsTransactionRecord(inverseOps, ops.length)
      const beforeFormulaForwardOps = captureFormulaCellStateForStructuralMoveUndo(
        op.sheetName,
        op.kind === 'insertRows' ? 'row' : 'column',
        op.start,
        op.count,
        op.start,
      )
      applyForward(
        potentialNewCells === undefined
          ? { kind: 'single-op', op: appliedForwardOp }
          : { kind: 'single-op', op: appliedForwardOp, potentialNewCells },
      )
      if (args.state.getTransactionReplayDepth() === 0) {
        const afterFormulaForwardOps = captureFormulaCellStateForStructuralMoveUndo(
          op.sheetName,
          op.kind === 'insertRows' ? 'row' : 'column',
          op.start,
          op.count,
          op.start,
        )
        const structuralFormulaForwardOps = structuralFormulaForwardCorrections(op, beforeFormulaForwardOps, afterFormulaForwardOps)
        const forwardOps = [appliedForwardOp, ...structuralFormulaForwardOps]
        args.state.undoStack.push({
          forward:
            forwardOps.length === 1
              ? createLazySingleOpTransactionRecord(forwardOps[0]!, potentialNewCells)
              : createOpsTransactionRecord(forwardOps, potentialNewCells),
          inverse,
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    if (
      options.returnUndoOps === false &&
      ops.length === 1 &&
      options.preparedCellAddressesByOpIndex === undefined &&
      ops[0]?.kind === 'renameSheet'
    ) {
      const op = ops[0]
      applyForward(potentialNewCells === undefined ? { kind: 'single-op', op } : { kind: 'single-op', op, potentialNewCells })
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazySingleOpTransactionRecord(
            { kind: 'renameSheet', oldName: op.oldName, newName: op.newName },
            potentialNewCells,
          ),
          inverse: createLazySingleOpTransactionRecord({ kind: 'renameSheet', oldName: op.newName, newName: op.oldName }),
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    if (
      options.returnUndoOps === false &&
      ops.length === 1 &&
      options.preparedCellAddressesByOpIndex === undefined &&
      ops[0]?.kind === 'upsertDefinedName'
    ) {
      const op = ops[0]
      const existing = args.state.workbook.getDefinedName(op.name)
      const inverseOp: EngineOp =
        existing === undefined
          ? { kind: 'deleteDefinedName', name: op.name }
          : { kind: 'upsertDefinedName', name: existing.name, value: structuredClone(existing.value) }
      applyForward(potentialNewCells === undefined ? { kind: 'single-op', op } : { kind: 'single-op', op, potentialNewCells })
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazySingleOpTransactionRecord({
            kind: 'upsertDefinedName',
            name: op.name,
            value: structuredClone(op.value),
          }),
          inverse: createLazySingleOpTransactionRecord(inverseOp),
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    const forward = createOpsTransactionRecord(ops, potentialNewCells, options.preparedCellAddressesByOpIndex)
    const baseFastHistoryArgs: Parameters<typeof tryBuildFastMutationHistory>[0] =
      potentialNewCells === undefined
        ? {
            workbook: args.state.workbook,
            getCellByIndex: args.getCellByIndex,
            ops,
            cloneForwardOps: options.reuseForwardOps !== true,
          }
        : {
            workbook: args.state.workbook,
            getCellByIndex: args.getCellByIndex,
            ops,
            potentialNewCells,
            includeUndoOps: options.returnUndoOps !== false,
            cloneForwardOps: options.reuseForwardOps !== true,
          }
    const fastHistoryArgs: Parameters<typeof tryBuildFastMutationHistory>[0] = options.preparedCellAddressesByOpIndex
      ? {
          ...baseFastHistoryArgs,
          preparedCellAddressesByOpIndex: options.preparedCellAddressesByOpIndex,
        }
      : baseFastHistoryArgs
    const fastHistory =
      tryBuildSingleCellOpHistoryWithoutSnapshot(
        ops,
        potentialNewCells,
        options.returnUndoOps !== false,
        options.reuseForwardOps !== true,
      ) ?? tryBuildFastMutationHistory(fastHistoryArgs)
    const inverse: TransactionRecord = fastHistory?.inverse ?? buildInverseRecord(ops, options.returnUndoOps !== false)
    applyForward(forward)
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward:
          fastHistory?.forward ??
          createOpsTransactionRecord(canonicalizeForwardOps(ops), potentialNewCells, options.preparedCellAddressesByOpIndex),
        inverse,
      })
      args.state.redoStack.length = 0
    }
    if (options.returnUndoOps === false) {
      return null
    }
    return fastHistory?.undoOps ?? cloneTransactionRecordOps(args.state.workbook, inverse)
  }

  const executeLocalCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
    options: {
      returnUndoOps?: boolean
      reuseRefs?: boolean
    } = {},
  ): readonly EngineOp[] | null => {
    if (refs.length === 0) {
      return null
    }
    const nextRefs = options.reuseRefs ? refs : refs.map((ref) => cloneCellMutationRef(ref))
    const nextPotentialNewCells = potentialNewCells ?? countPotentialNewCellsForMutationRefs(nextRefs)
    const shouldCreateBatch = shouldCreateLocalBatch()
    if (options.returnUndoOps === false) {
      const inverse =
        tryCreateSingleExistingNumericInverseCellMutationRecord(nextRefs) ??
        tryCreateExistingNumericInverseCellMutationRecord(nextRefs) ??
        createLazyInverseCellMutationRecord(nextRefs)
      const batch = shouldCreateBatch
        ? createBatch(
            args.state.replicaState,
            nextRefs.map((ref) => cellMutationRefToEngineOp(args.state.workbook, ref)),
          )
        : null
      args.applyCellMutationsAtBatchNow(nextRefs, batch, 'local', nextPotentialNewCells)
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazyCellMutationTransactionRecord(nextRefs, nextPotentialNewCells),
          inverse,
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    const fastHistory = buildFastMutationHistoryFromRefs(nextRefs, nextPotentialNewCells)
    const inverse: TransactionRecord = fastHistory?.inverse ?? {
      kind: 'ops',
      ops: buildInverseOps(transactionRecordOps(args.state.workbook, fastHistory.forward)),
      potentialNewCells: transactionRecordOps(args.state.workbook, fastHistory.forward).length,
    }
    const batch = shouldCreateBatch
      ? createBatch(args.state.replicaState, [...transactionRecordOps(args.state.workbook, fastHistory.forward)])
      : null
    args.applyCellMutationsAtBatchNow(nextRefs, batch, 'local', nextPotentialNewCells)
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: fastHistory?.forward ?? {
          kind: 'ops',
          ops: canonicalizeForwardOps([...transactionRecordOps(args.state.workbook, fastHistory.forward)]),
          potentialNewCells: nextPotentialNewCells,
        },
        inverse,
      })
      args.state.redoStack.length = 0
    }
    return fastHistory?.undoOps ?? cloneTransactionRecordOps(args.state.workbook, inverse)
  }

  const executeLocalExistingNumericCellMutationAtNow = (
    request: EngineExistingNumericCellMutationRef,
    options: {
      returnUndoOps?: boolean
    } = {},
  ): EngineExistingNumericCellMutationResult | null => {
    if (options.returnUndoOps !== false || shouldCreateLocalBatch()) {
      return null
    }
    const cellStore = args.state.workbook.cellStore
    const cellIndex = request.cellIndex
    const oldNumericValue =
      request.trustedExistingNumericLiteral && request.oldNumericValue !== undefined
        ? request.oldNumericValue
        : (cellStore.numbers[cellIndex] ?? 0)
    const result = args.applyExistingNumericCellMutationAtNow?.(request)
    if (!result) {
      return null
    }
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: createSingleExistingNumericCellMutationTransactionRecord(request, 0),
        inverse: createSingleExistingNumericCellMutationTransactionRecord(
          {
            sheetId: request.sheetId,
            cellIndex,
            row: request.row,
            col: request.col,
            value: oldNumericValue,
          },
          1,
        ),
      })
      args.state.redoStack.length = 0
    }
    return result
  }

  const executeLocalExistingLiteralCellMutationAtNow = (
    request: EngineExistingLiteralCellMutationRef,
    options: {
      returnUndoOps?: boolean
    } = {},
  ): EngineExistingNumericCellMutationResult | null => {
    if (typeof request.value === 'number') {
      return executeLocalExistingNumericCellMutationAtNow(
        {
          sheetId: request.sheetId,
          row: request.row,
          col: request.col,
          cellIndex: request.cellIndex,
          value: request.value,
          ...(request.emitTracked === undefined ? {} : { emitTracked: request.emitTracked }),
        },
        options,
      )
    }
    if (options.returnUndoOps !== false || shouldCreateLocalBatch()) {
      return null
    }
    const ref: EngineCellMutationRef = {
      sheetId: request.sheetId,
      cellIndex: request.cellIndex,
      mutation: {
        kind: 'setCellValue',
        row: request.row,
        col: request.col,
        value: request.value,
      },
    }
    const inverse = createLazyInverseCellMutationRecord([ref])
    const result = args.applyExistingLiteralCellMutationAtNow?.(request)
    if (!result) {
      return null
    }
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: createSingleExistingLiteralCellMutationTransactionRecord(request, 0),
        inverse,
      })
      args.state.redoStack.length = 0
    }
    return result
  }

  const applyCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    } = {},
  ): readonly EngineOp[] | null => {
    const source = options.source ?? 'restore'
    const captureUndo = options.captureUndo ?? source === 'local'
    if (captureUndo) {
      const executeOptions: {
        returnUndoOps?: boolean
        reuseRefs?: boolean
      } = {}
      if (options.returnUndoOps !== undefined) {
        executeOptions.returnUndoOps = options.returnUndoOps
      }
      if (options.reuseRefs !== undefined) {
        executeOptions.reuseRefs = options.reuseRefs
      }
      return executeLocalCellMutationsAtNow(refs, options.potentialNewCells, {
        ...executeOptions,
      })
    }
    if (refs.length === 0) {
      return null
    }
    const nextRefs = options.reuseRefs ? refs : refs.map((ref) => cloneCellMutationRef(ref))
    const nextPotentialNewCells = options.potentialNewCells ?? countPotentialNewCellsForMutationRefs(nextRefs)
    const forwardOps = source === 'restore' ? emptyBatchOps : nextRefs.map((ref) => cellMutationRefToEngineOp(args.state.workbook, ref))
    const batch =
      source === 'local' && shouldCreateLocalBatch()
        ? createBatch(args.state.replicaState, forwardOps)
        : source === 'restore'
          ? null
          : createBatch(args.state.replicaState, forwardOps)
    args.applyCellMutationsAtBatchNow(nextRefs, batch, source, nextPotentialNewCells)
    return null
  }

  const executeLocalNowPublic = (
    ops: EngineOp[],
    potentialNewCells?: number,
    options: { readonly returnUndoOps?: boolean; readonly emitTracked?: boolean } = {},
  ): readonly EngineOp[] | null => {
    if (!shouldCreateLocalBatch()) {
      const refs = tryCellMutationRefsFromOps(ops)
      if (refs !== null) {
        return executeLocalCellMutationsAtNow(refs, potentialNewCells, {
          returnUndoOps: options.returnUndoOps ?? true,
          reuseRefs: true,
        })
      }
    }
    return executeLocalNowWithCustomApply(
      ops,
      potentialNewCells,
      (forward) => {
        executeTransactionNow(forward, 'local', options.emitTracked === undefined ? {} : { emitTracked: options.emitTracked })
      },
      { returnUndoOps: options.returnUndoOps ?? true, reuseForwardOps: false },
    )
  }

  const executeLocal = (
    ops: EngineOp[],
    potentialNewCells?: number,
    options: { readonly returnUndoOps?: boolean } = {},
  ): Effect.Effect<readonly EngineOp[] | null, EngineMutationError> =>
    Effect.try({
      try: () => executeLocalNowPublic(ops, potentialNewCells, options),
      catch: (cause) =>
        new EngineMutationError({
          message: 'Failed to execute local transaction',
          cause,
        }),
    })

  const rangeOperations = createMutationRangeOperations({
    workbook: args.state.workbook,
    getCellByIndex: args.getCellByIndex,
    readRangeCells: args.readRangeCells,
    toCellStateOps: args.toCellStateOps,
    executeLocal,
    ...(args.settleImportedFormulas === undefined ? {} : { settleImportedFormulas: args.settleImportedFormulas }),
  })

  return {
    executeTransactionNow: executeTransactionNow,
    executeTransaction(record, source) {
      return Effect.try({
        try: () => {
          executeTransactionNow(record, source)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: `Failed to execute ${source} transaction`,
            cause,
          }),
      })
    },
    executeLocalNow: executeLocalNowPublic,
    executeLocalCellMutationsAtNow(refs, potentialNewCells) {
      return executeLocalCellMutationsAtNow(refs, potentialNewCells)
    },
    executeLocalExistingNumericCellMutationAtNow(request, options = {}) {
      return executeLocalExistingNumericCellMutationAtNow(request, options)
    },
    executeLocalExistingLiteralCellMutationAtNow(request, options = {}) {
      return executeLocalExistingLiteralCellMutationAtNow(request, options)
    },
    applyCellMutationsAtNow(refs, options = {}) {
      return applyCellMutationsAtNow(refs, options)
    },
    applyCellMutationsAt(refs, options = {}) {
      return Effect.try({
        try: () => applyCellMutationsAtNow(refs, options),
        catch: (cause) =>
          new EngineMutationError({
            message: 'Failed to apply cell mutations',
            cause,
          }),
      })
    },
    executeLocal,
    applyOpsNow(ops, options = {}) {
      const nextOps = options.trusted ? Array.from(ops) : structuredClone([...ops])
      if (nextOps.length === 0) {
        return null
      }
      if (options.captureUndo) {
        return options.returnUndoOps === undefined
          ? this.executeLocalNow(nextOps, options.potentialNewCells)
          : this.executeLocalNow(nextOps, options.potentialNewCells, { returnUndoOps: options.returnUndoOps })
      }
      executeTransactionNow(
        options.potentialNewCells === undefined
          ? { kind: 'ops', ops: nextOps }
          : { kind: 'ops', ops: nextOps, potentialNewCells: options.potentialNewCells },
        options.source ?? 'restore',
      )
      return null
    },
    applyOps(ops, options = {}) {
      return Effect.try({
        try: () => this.applyOpsNow(ops, options),
        catch: (cause) =>
          new EngineMutationError({
            message: 'Failed to apply engine operations',
            cause,
          }),
      })
    },
    captureUndoOps(mutate) {
      return Effect.try({
        try: () => {
          const previousUndoDepth = args.state.undoStack.length
          const result = mutate()
          if (args.state.undoStack.length === previousUndoDepth) {
            return {
              result,
              undoOps: null,
            }
          }
          if (args.state.undoStack.length === previousUndoDepth + 1) {
            const inverse = args.state.undoStack.at(-1)!.inverse
            return {
              result,
              undoOps: cloneTransactionRecordOps(args.state.workbook, inverse),
            }
          }
          throw new Error('Expected a single local transaction while capturing undo ops')
        },
        catch: (cause) =>
          new EngineMutationError({
            message: 'Failed to capture undo ops',
            cause,
          }),
      })
    },
    setRangeValues: rangeOperations.setRangeValues,
    setRangeFormulas: rangeOperations.setRangeFormulas,
    clearRange: rangeOperations.clearRange,
    fillRange: rangeOperations.fillRange,
    copyRange: rangeOperations.copyRange,
    moveRange: rangeOperations.moveRange,
    importSheetCsv: rangeOperations.importSheetCsv,
    renderCommit(ops) {
      return Effect.flatMap(
        Effect.try({
          try: () => {
            if (
              tryExecuteMutationRenderCommitFastPath({
                state: args.state,
                ops,
                hasExternallyVisibleBatchRequirement,
                restoreCellOpFromRef,
                executeLocalNowWithCustomApply,
                executeTransactionNow,
                applyCellMutationsAtNow,
              })
            ) {
              return null
            }
            return normalizeRenderCommitOps(ops)
          },
          catch: (cause) =>
            new EngineMutationError({
              message: 'Failed to normalize render commit operations',
              cause,
            }),
        }),
        (normalized) => {
          if (normalized === null) {
            return Effect.void
          }
          const { engineOps, potentialNewCells, preparedCellAddressesByOpIndex } = normalized
          return Effect.try({
            try: () => {
              executeLocalNowWithCustomApply(
                engineOps,
                potentialNewCells,
                (forward) => {
                  const batchOps = [...transactionRecordOps(args.state.workbook, forward)]
                  args.applyBatchNow(
                    createBatch(args.state.replicaState, batchOps),
                    'local',
                    forward.potentialNewCells,
                    preparedCellAddressesByOpIndex,
                  )
                },
                {
                  returnUndoOps: false,
                  reuseForwardOps: true,
                  preparedCellAddressesByOpIndex,
                },
              )
            },
            catch: (cause) =>
              new EngineMutationError({
                message: 'Failed to execute render commit transaction',
                cause,
              }),
          })
        },
      ).pipe(Effect.asVoid)
    },
  }
}
