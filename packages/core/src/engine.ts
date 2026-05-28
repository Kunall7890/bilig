import {
  ErrorCode,
  ValueTag,
  type CellRangeRef,
  type CellSnapshot,
  type CellValue,
  type DependencySnapshot,
  type EngineEvent,
  type ExplainCellSnapshot,
  type LiteralInput,
  type RecalcMetrics,
  type SelectionState,
  type SyncState,
  type WorkbookSnapshot,
} from '@bilig/protocol'
import type { CsvParseOptions } from './csv.js'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook'
import type {
  EngineCellMutationRef,
  EngineExistingLiteralCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationsRef,
  EngineExistingNumericCellMutationResult,
  EngineFormulaSourceRefs,
} from './cell-mutations-at.js'
import { CellFlags, type CellStore } from './cell-store.js'
import { addEngineCounter, cloneEngineCounters, resetEngineCounters, type EngineCounters } from './perf/engine-counters.js'
import type { CommitOp, EngineReplicaSnapshot, EngineSyncClient, SpreadsheetEngineOptions } from './engine/runtime-state.js'
import { runEngineEffect, runEngineEffectPromise } from './engine/live.js'
import { SpreadsheetEngineWorkbookFacadeBase } from './engine/engine-workbook-facade-base.js'
import {
  createTrustedExistingNumericDirectScalarFastPath,
  type TrustedExistingNumericDirectScalarFastPath,
} from './engine/services/trusted-existing-numeric-direct-scalar-fast-path.js'
import {
  createLazyCellMutationTransactionRecord,
  createExistingNumericCellMutationsTransactionRecord,
  createLazyReversedExistingNumericCellMutationsTransactionRecord,
} from './engine/services/mutation-transaction-records.js'
import {
  directScalarCellNumber,
  evaluateDirectScalarWithReplacementNumbers,
  writeSingleInputAffineDirectScalar,
} from './engine/services/direct-scalar-helpers.js'
import { tagTrustedPhysicalTrackedChanges } from './engine/services/operation-change-helpers.js'
import { buildDirectScalarDescriptor } from './engine/services/formula-binding-direct-scalar.js'
import { directScalarDependencyCellsEqual } from './engine/services/formula-binding-shape-helpers.js'
import { isWorkbookTableHeaderCell } from './engine/services/operation-table-header-rename.js'
import { tryCompileSimpleDirectScalarFormula } from './formula/simple-direct-scalar-compile.js'
import type { EngineSnapshotExportOptions } from './engine/services/snapshot-service.js'

export type {
  CommitOp,
  EngineReplicaSnapshot,
  EngineSyncClient,
  EngineSyncClientConnection,
  SpreadsheetEngineOptions,
} from './engine/runtime-state.js'
export { selectors } from './engine-selectors.js'

const DIRECT_SCALAR_FORMULA_REPLACEMENT_LIMIT = 4_096
const DIRECT_SCALAR_FORMULA_REPLACEMENT_INITIAL_CAPACITY = 2_048
const DIRECT_SCALAR_FORMULA_REPLACEMENT_BLOCKED_FLAGS = CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput

function cellStoreValueToLiteralInput(
  cellStore: {
    readonly tags: ArrayLike<ValueTag | undefined>
    readonly numbers: ArrayLike<number | undefined>
    readonly stringIds: ArrayLike<number | undefined>
  },
  strings: { get(stringId: number): string },
  cellIndex: number,
): LiteralInput | undefined {
  switch (cellStore.tags[cellIndex] ?? ValueTag.Empty) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return cellStore.numbers[cellIndex] ?? 0
    case ValueTag.Boolean:
      return (cellStore.numbers[cellIndex] ?? 0) !== 0
    case ValueTag.String: {
      const stringId = cellStore.stringIds[cellIndex] ?? 0
      return stringId === 0 ? '' : strings.get(stringId)
    }
    case ValueTag.Error:
      return undefined
  }
}

function cellStoreCanRewriteDirectScalarFormula(
  cellStore: {
    readonly tags: ArrayLike<ValueTag | undefined>
    readonly stringIds: ArrayLike<number | undefined>
    readonly errors: ArrayLike<ErrorCode | undefined>
    readonly flags: ArrayLike<number | undefined>
  },
  cellIndex: number,
): boolean {
  return (
    cellStore.tags[cellIndex] === ValueTag.Number &&
    (cellStore.stringIds[cellIndex] ?? 0) === 0 &&
    (cellStore.errors[cellIndex] ?? ErrorCode.None) === ErrorCode.None &&
    ((cellStore.flags[cellIndex] ?? 0) & DIRECT_SCALAR_FORMULA_REPLACEMENT_BLOCKED_FLAGS) === 0
  )
}

function writeTrustedDirectScalarFormulaNumber(
  cellStore: Pick<CellStore, 'flags' | 'versions' | 'stringIds' | 'tags' | 'numbers' | 'errors'>,
  cellIndex: number,
  value: number,
): void {
  cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  cellStore.stringIds[cellIndex] = 0
  cellStore.tags[cellIndex] = ValueTag.Number
  cellStore.numbers[cellIndex] = value
  cellStore.errors[cellIndex] = ErrorCode.None
}

function isStrictlyIncreasing(values: Uint32Array): boolean {
  for (let index = 1; index < values.length; index += 1) {
    if (values[index]! <= values[index - 1]!) {
      return false
    }
  }
  return true
}

export class SpreadsheetEngine extends SpreadsheetEngineWorkbookFacadeBase {
  private trustedExistingNumericDirectScalarFastPath: TrustedExistingNumericDirectScalarFastPath | undefined

  async ready(): Promise<void> {
    await this.wasm.init()
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    return runEngineEffect(this.runtime.events.subscribe(listener))
  }

  subscribeCell(sheetName: string, address: string, listener: () => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeCell(sheetName, address, listener))
  }

  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeCells(sheetName, addresses, listener))
  }

  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void {
    return runEngineEffect(this.runtime.events.subscribeBatches(listener))
  }

  subscribeSelection(listener: () => void): () => void {
    return runEngineEffect(this.runtime.selection.subscribe(listener))
  }

  getSelectionState(): SelectionState {
    return runEngineEffect(this.runtime.selection.getSelectionState())
  }

  setSelection(
    sheetName: string,
    address: string | null,
    options: {
      anchorAddress?: string | null
      range?: { startAddress: string; endAddress: string } | null
      editMode?: SelectionState['editMode']
    } = {},
  ): void {
    runEngineEffect(this.runtime.selection.setSelection(sheetName, address, options))
  }

  getLastMetrics(): RecalcMetrics {
    return this.lastMetrics
  }

  getPerformanceCounters(): EngineCounters {
    return cloneEngineCounters(this.performanceCounters)
  }

  resetPerformanceCounters(): void {
    resetEngineCounters(this.performanceCounters)
  }

  resetForReuse(options: SpreadsheetEngineOptions = {}): void {
    runEngineEffect(this.runtime.maintenance.resetWorkbook(options.workbookName ?? 'Workbook'))
    this.setUseColumnIndexEnabled(options.useColumnIndex ?? true)
    this.setEvaluationTimeoutMs(options.evaluationTimeoutMs)
    this.resetPerformanceCounters()
  }

  setUseColumnIndexEnabled(enabled: boolean): void {
    this.state.setUseColumnIndex(enabled)
  }

  getSyncState(): SyncState {
    return this.syncState
  }

  async connectSyncClient(client: EngineSyncClient): Promise<void> {
    if (!this.state.trackReplicaVersions) {
      throw new Error('Sync is unavailable when trackReplicaVersions is disabled; construct the engine with trackReplicaVersions enabled.')
    }
    await runEngineEffectPromise(this.runtime.sync.connectClient(client))
  }

  async disconnectSyncClient(): Promise<void> {
    await runEngineEffectPromise(this.runtime.sync.disconnectClient())
  }

  createSheet(name: string): void {
    this.executeLocalTransaction([{ kind: 'upsertSheet', name, order: this.workbook.sheetsByName.size }])
  }

  createSheetForInitialization(name: string): number {
    return this.workbook.createSheet(name, this.workbook.sheetsByName.size).id
  }

  moveSheet(name: string, order: number): void {
    if (!this.workbook.getSheet(name)) {
      return
    }
    const nextOrder = Number.isFinite(order) ? Math.max(0, Math.trunc(order)) : 0
    this.executeLocalTransaction([{ kind: 'upsertSheet', name, order: nextOrder }])
  }

  renameSheet(oldName: string, newName: string): void {
    const trimmedName = newName.trim()
    if (trimmedName.length === 0 || oldName === trimmedName) {
      return
    }
    if (this.workbook.getSheet(trimmedName)) {
      return
    }
    this.executeLocalTransaction([{ kind: 'renameSheet', oldName, newName: trimmedName }])
  }

  deleteSheet(name: string): void {
    this.executeLocalTransaction([{ kind: 'deleteSheet', name }])
  }

  setCellValue(sheetName: string, address: string, value: LiteralInput): CellValue {
    this.executeLocalTransaction([{ kind: 'setCellValue', sheetName, address, value }])
    return this.getCellValue(sheetName, address)
  }

  setCellValueAt(sheetId: number, row: number, col: number, value: LiteralInput): CellValue {
    const sheetName = this.workbook.getSheetById(sheetId)?.name
    if (!sheetName) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    const address = formatAddress(row, col)
    this.runtime.mutation.executeLocalCellMutationsAtNow([{ sheetId, mutation: { kind: 'setCellValue', row, col, value } }], 1, {
      returnUndoOps: false,
    })
    return this.getCellValue(sheetName, address)
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellValue {
    if (this.getCell(sheetName, address).formula === formula) {
      return this.getCellValue(sheetName, address)
    }
    this.executeLocalTransaction([{ kind: 'setCellFormula', sheetName, address, formula }])
    return this.getCellValue(sheetName, address)
  }

  setCellFormulaAt(sheetId: number, row: number, col: number, formula: string): CellValue {
    const sheetName = this.workbook.getSheetById(sheetId)?.name
    if (!sheetName) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    const address = formatAddress(row, col)
    if (this.getCell(sheetName, address).formula === formula) {
      return this.getCellValue(sheetName, address)
    }
    this.runtime.mutation.executeLocalCellMutationsAtNow([{ sheetId, mutation: { kind: 'setCellFormula', row, col, formula } }], 1, {
      returnUndoOps: false,
    })
    return this.getCellValue(sheetName, address)
  }

  tryApplyExistingDirectScalarFormulaMutationAt(request: {
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly cellIndex: number
    readonly formula: string
  }): boolean {
    if (
      this.state.trackReplicaVersions ||
      this.batchListeners.size > 0 ||
      this.syncClientConnection !== null ||
      this.events.hasListeners() ||
      this.events.hasCellListeners() ||
      this.runtime.hasVolatileFormulas() ||
      this.runtime.hasRegionFormulaSubscriptions() ||
      this.reverseExactLookupColumnEdges.size > 0 ||
      this.reverseSortedLookupColumnEdges.size > 0 ||
      this.reverseAggregateColumnEdges.size > 0
    ) {
      return false
    }
    const sheet = this.workbook.getSheetById(request.sheetId)
    const cellStore = this.workbook.cellStore
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      cellStore.sheetIds[request.cellIndex] !== request.sheetId ||
      cellStore.rows[request.cellIndex] !== request.row ||
      cellStore.cols[request.cellIndex] !== request.col ||
      isWorkbookTableHeaderCell(this.workbook, sheet.name, request.row, request.col) ||
      !cellStoreCanRewriteDirectScalarFormula(cellStore, request.cellIndex)
    ) {
      return false
    }
    const existingFormula = this.formulas.get(request.cellIndex)
    if (
      !existingFormula ||
      existingFormula.directScalar === undefined ||
      existingFormula.directLookup !== undefined ||
      existingFormula.directAggregate !== undefined ||
      existingFormula.directCriteria !== undefined ||
      existingFormula.compiled.volatile ||
      existingFormula.compiled.producesSpill
    ) {
      return false
    }
    if (existingFormula.source === request.formula) {
      return true
    }
    const oldFormulaSource = existingFormula.source
    const compiled = tryCompileSimpleDirectScalarFormula(request.formula)
    if (
      compiled === undefined ||
      compiled.volatile ||
      compiled.producesSpill ||
      compiled.symbolicNames.length !== 0 ||
      compiled.symbolicTables.length !== 0 ||
      compiled.symbolicSpills.length !== 0 ||
      compiled.symbolicRanges.length !== 0
    ) {
      return false
    }
    const resolveExistingCell = (sheetName: string, address: string): number => {
      const ownerSheet = this.workbook.getSheet(sheetName)
      if (!ownerSheet) {
        return -1
      }
      try {
        const parsed = parseCellAddress(address, sheetName)
        return this.workbook.getCellIndexAt(ownerSheet.id, parsed.row, parsed.col) ?? -1
      } catch {
        return -1
      }
    }
    const replacementDirectScalar = buildDirectScalarDescriptor({
      compiled,
      ownerSheetName: sheet.name,
      ownerSheetId: sheet.id,
      workbook: this.workbook,
      ensureCellTracked: resolveExistingCell,
      ensureCellTrackedByCoords: (sheetId, refRow, refCol) => this.workbook.getCellIndexAt(sheetId, refRow, refCol) ?? -1,
    })
    if (!replacementDirectScalar || !directScalarDependencyCellsEqual(existingFormula.directScalar, replacementDirectScalar)) {
      return false
    }

    let changed = new Uint32Array(DIRECT_SCALAR_FORMULA_REPLACEMENT_INITIAL_CAPACITY)
    let values = new Float64Array(DIRECT_SCALAR_FORMULA_REPLACEMENT_INITIAL_CAPACITY)
    let changedCount = 0
    const appendChanged = (cellIndex: number, value: number): void => {
      if (changedCount >= changed.length) {
        const nextChanged = new Uint32Array(changed.length * 2)
        nextChanged.set(changed)
        changed = nextChanged
        const nextValues = new Float64Array(values.length * 2)
        nextValues.set(values)
        values = nextValues
      }
      changed[changedCount] = cellIndex
      values[changedCount] = value
      changedCount += 1
    }
    const readStagedCellNumber = (cellIndex: number): number | undefined => {
      for (let index = changedCount - 1; index >= 0; index -= 1) {
        if (changed[index] === cellIndex) {
          return values[index]
        }
      }
      return directScalarCellNumber(cellStore, cellIndex)
    }
    const nextRootNumber = evaluateDirectScalarWithReplacementNumbers(replacementDirectScalar, -1, 0, readStagedCellNumber)
    if (nextRootNumber === undefined || !Number.isFinite(nextRootNumber)) {
      return false
    }
    appendChanged(request.cellIndex, nextRootNumber)

    let currentCellIndex = request.cellIndex
    const affine = { scale: 0, offset: 0 }
    for (;;) {
      if (changedCount >= DIRECT_SCALAR_FORMULA_REPLACEMENT_LIMIT) {
        return false
      }
      const formulaCellIndex = this.runtime.traversal.getSingleCellDependentNow(currentCellIndex)
      if (formulaCellIndex === -1) {
        break
      }
      if (formulaCellIndex < 0) {
        return false
      }
      const dependentFormula = this.formulas.get(formulaCellIndex)
      if (
        formulaCellIndex === request.cellIndex ||
        !dependentFormula ||
        dependentFormula.directScalar === undefined ||
        dependentFormula.compiled.volatile ||
        dependentFormula.compiled.producesSpill ||
        dependentFormula.directLookup !== undefined ||
        dependentFormula.directAggregate !== undefined ||
        dependentFormula.directCriteria !== undefined ||
        !cellStoreCanRewriteDirectScalarFormula(cellStore, formulaCellIndex)
      ) {
        return false
      }
      const previousNumber = values[changedCount - 1]!
      const nextFormulaNumber = writeSingleInputAffineDirectScalar(dependentFormula.directScalar, currentCellIndex, affine)
        ? previousNumber * affine.scale + affine.offset
        : evaluateDirectScalarWithReplacementNumbers(dependentFormula.directScalar, currentCellIndex, previousNumber, readStagedCellNumber)
      if (nextFormulaNumber === undefined || !Number.isFinite(nextFormulaNumber)) {
        return false
      }
      appendChanged(formulaCellIndex, nextFormulaNumber)
      currentCellIndex = formulaCellIndex
    }

    if (!this.runtime.binding.rewriteFormulaCompiledPreservingBindingNow(request.cellIndex, request.formula, compiled)) {
      return false
    }
    const changedCellIndices = changed.subarray(0, changedCount)
    const nextNumbers = values.subarray(0, changedCount)
    for (let index = 0; index < changedCount; index += 1) {
      writeTrustedDirectScalarFormulaNumber(cellStore, changedCellIndices[index]!, nextNumbers[index]!)
    }
    addEngineCounter(this.performanceCounters, 'directScalarDeltaApplications', Math.max(0, changedCount - 1))
    addEngineCounter(this.performanceCounters, 'directScalarDeltaOnlyRecalcSkips')
    const metrics = {
      batchId: this.lastMetrics.batchId + 1,
      changedInputCount: 1,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
    }
    this.lastMetrics = metrics
    if (
      changedCellIndices.length > 1 &&
      isStrictlyIncreasing(changedCellIndices) &&
      cellStore.sheetIds[changedCellIndices[0]!] === request.sheetId
    ) {
      let sameSheet = true
      for (let index = 1; index < changedCellIndices.length; index += 1) {
        if (cellStore.sheetIds[changedCellIndices[index]!] !== request.sheetId) {
          sameSheet = false
          break
        }
      }
      if (sameSheet) {
        tagTrustedPhysicalTrackedChanges(changedCellIndices, request.sheetId, 1)
      }
    }
    if (this.events.hasTrackedListeners()) {
      this.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics,
        explicitChangedCount: 1,
      })
    }
    if (this.transactionReplayDepth === 0) {
      this.undoStack.push({
        forward: createLazyCellMutationTransactionRecord(
          [
            {
              sheetId: request.sheetId,
              cellIndex: request.cellIndex,
              mutation: { kind: 'setCellFormula', row: request.row, col: request.col, formula: request.formula },
            },
          ],
          0,
        ),
        inverse: createLazyCellMutationTransactionRecord(
          [
            {
              sheetId: request.sheetId,
              cellIndex: request.cellIndex,
              mutation: { kind: 'setCellFormula', row: request.row, col: request.col, formula: oldFormulaSource },
            },
          ],
          0,
        ),
      })
      this.redoStack.length = 0
    }
    return true
  }

  setCellFormat(sheetName: string, address: string, format: string | null): void {
    this.executeLocalTransaction([{ kind: 'setCellFormat', sheetName, address, format }])
  }

  clearCellAt(sheetId: number, row: number, col: number): void {
    this.runtime.mutation.executeLocalCellMutationsAtNow([{ sheetId, mutation: { kind: 'clearCell', row, col } }], 0, {
      returnUndoOps: false,
    })
  }

  applyCellMutationsAt(refs: readonly EngineCellMutationRef[], potentialNewCells?: number): readonly EngineOp[] | null {
    return this.runtime.mutation.executeLocalCellMutationsAtNow(refs, potentialNewCells)
  }

  applyCellMutationsAtWithOptions(
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    } = {},
  ): readonly EngineOp[] | null {
    const nextOptions = options.source === undefined ? { ...options, source: 'local' as const } : options
    return this.runtime.mutation.applyCellMutationsAtNow(refs, nextOptions)
  }

  tryApplyExistingNumericCellMutationAt(request: EngineExistingNumericCellMutationRef): EngineExistingNumericCellMutationResult | null {
    if (this.state.trackReplicaVersions || this.batchListeners.size > 0 || this.syncClientConnection !== null) {
      return null
    }
    const cellStore = this.workbook.cellStore
    const cellIndex = request.cellIndex
    const oldNumericValue =
      request.trustedExistingNumericLiteral && request.oldNumericValue !== undefined
        ? request.oldNumericValue
        : (cellStore.numbers[cellIndex] ?? 0)
    const result =
      this.getTrustedExistingNumericDirectScalarFastPath().tryApply(request, oldNumericValue) ??
      this.runtime.operations.applyExistingNumericCellMutationAtNow(request)
    if (!result) {
      return null
    }
    if (this.transactionReplayDepth === 0) {
      this.undoStack.push({
        forward: {
          kind: 'single-existing-numeric-cell-mutation',
          sheetId: request.sheetId,
          row: request.row,
          col: request.col,
          cellIndex,
          value: request.value,
          potentialNewCells: 0,
        },
        inverse: {
          kind: 'single-existing-numeric-cell-mutation',
          sheetId: request.sheetId,
          row: request.row,
          col: request.col,
          cellIndex,
          value: oldNumericValue,
          potentialNewCells: 1,
        },
      })
      this.redoStack.length = 0
    }
    return result
  }

  private getTrustedExistingNumericDirectScalarFastPath(): TrustedExistingNumericDirectScalarFastPath {
    return (this.trustedExistingNumericDirectScalarFastPath ??= createTrustedExistingNumericDirectScalarFastPath({
      workbook: this.workbook,
      formulas: this.formulas,
      events: this.events,
      counters: this.performanceCounters,
      traversal: this.runtime.traversal,
      directScalarDeltaOutputCellIndicesByInput: this.state.directScalarDeltaOutputCellIndicesByInput,
      deferKernelSync: this.runtime.deferKernelSync,
      hasVolatileFormulas: this.runtime.hasVolatileFormulas,
      hasRegionFormulaSubscriptions: this.runtime.hasRegionFormulaSubscriptions,
      reverseExactLookupColumnEdges: this.reverseExactLookupColumnEdges,
      reverseSortedLookupColumnEdges: this.reverseSortedLookupColumnEdges,
      reverseAggregateColumnEdges: this.reverseAggregateColumnEdges,
      getLastMetrics: () => this.lastMetrics,
      setLastMetrics: (metrics) => {
        this.lastMetrics = metrics
      },
    }))
  }

  tryApplyExistingNumericCellMutationsAt(request: EngineExistingNumericCellMutationsRef): boolean {
    if (this.state.trackReplicaVersions || this.batchListeners.size > 0 || this.syncClientConnection !== null) {
      return false
    }
    const oldNumbers = this.getTrustedExistingNumericDirectScalarFastPath().tryApplyBatch(request)
    if (oldNumbers !== null) {
      if (this.transactionReplayDepth === 0) {
        this.undoStack.push({
          forward: createExistingNumericCellMutationsTransactionRecord(
            {
              sheetIds: request.sheetIds,
              cellIndexPlusOnes: request.cellIndexPlusOnes,
              rows: request.rows,
              cols: request.cols,
              numbers: request.numbers,
            },
            0,
          ),
          inverse: createLazyReversedExistingNumericCellMutationsTransactionRecord(
            {
              sheetIds: request.sheetIds,
              cellIndexPlusOnes: request.cellIndexPlusOnes,
              rows: request.rows,
              cols: request.cols,
            },
            oldNumbers,
            1,
          ),
        })
        this.redoStack.length = 0
      }
      return true
    }
    return this.runtime.mutation.executeLocalExistingNumericCellMutationsAtNow(request, { returnUndoOps: false })
  }

  tryApplyExistingLiteralCellMutationAt(request: EngineExistingLiteralCellMutationRef): EngineExistingNumericCellMutationResult | null {
    if (typeof request.value === 'number') {
      return this.tryApplyExistingNumericCellMutationAt({
        sheetId: request.sheetId,
        row: request.row,
        col: request.col,
        cellIndex: request.cellIndex,
        value: request.value,
        ...(request.emitTracked === undefined ? {} : { emitTracked: request.emitTracked }),
      })
    }
    if (this.state.trackReplicaVersions || this.batchListeners.size > 0 || this.syncClientConnection !== null) {
      return null
    }
    const cellStore = this.workbook.cellStore
    const oldLiteralValue = cellStoreValueToLiteralInput(cellStore, this.strings, request.cellIndex)
    if (oldLiteralValue === undefined) {
      return null
    }
    const result = this.runtime.operations.applyExistingLiteralCellMutationAtNow(request)
    if (!result) {
      return null
    }
    if (this.transactionReplayDepth === 0) {
      this.undoStack.push({
        forward: {
          kind: 'single-existing-literal-cell-mutation',
          sheetId: request.sheetId,
          row: request.row,
          col: request.col,
          cellIndex: request.cellIndex,
          value: request.value,
          potentialNewCells: 0,
        },
        inverse: {
          kind: 'single-existing-literal-cell-mutation',
          sheetId: request.sheetId,
          row: request.row,
          col: request.col,
          cellIndex: request.cellIndex,
          value: oldLiteralValue,
          potentialNewCells: 1,
        },
      })
      this.redoStack.length = 0
    }
    return result
  }

  initializeCellFormulasAt(refs: readonly EngineCellMutationRef[], potentialNewCells?: number): void {
    runEngineEffect(this.runtime.formulaInitialization.initializeCellFormulasAt(refs, potentialNewCells))
  }

  initializeCellFormulasAtNow(refs: readonly EngineCellMutationRef[], potentialNewCells?: number): void {
    this.runtime.formulaInitialization.initializeCellFormulasAtNow(refs, potentialNewCells)
  }

  initializeFormulaSourcesAtNow(refs: EngineFormulaSourceRefs, potentialNewCells?: number): void {
    this.runtime.formulaInitialization.initializeFormulaSourcesAtNow(refs, potentialNewCells)
  }

  recalculateNow(): number[] {
    return this.runtime.recalc.recalculateAllNowSync()
  }

  recalculateChangedValuesNowForRebuild(): number[] {
    return this.runtime.recalc.recalculateChangedValuesNowForRebuildSync()
  }

  recalculateDifferential(): { js: CellSnapshot[]; wasm: CellSnapshot[]; drift: string[] } {
    return runEngineEffect(this.runtime.recalc.recalculateDifferential())
  }

  recalculateDirty(
    dirtyRegions: Array<{
      sheetName: string
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    }>,
  ): number[] {
    return runEngineEffect(this.runtime.recalc.recalculateDirty(dirtyRegions))
  }

  clearCell(sheetName: string, address: string): void {
    this.executeLocalTransaction([{ kind: 'clearCell', sheetName, address }])
  }

  setRangeValues(range: CellRangeRef, values: readonly (readonly LiteralInput[])[]): void {
    runEngineEffect(this.runtime.mutation.setRangeValues(range, values))
  }

  setRangeFormulas(range: CellRangeRef, formulas: readonly (readonly string[])[]): void {
    runEngineEffect(this.runtime.mutation.setRangeFormulas(range, formulas))
  }

  clearRange(range: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.clearRange(range))
  }

  fillRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.fillRange(source, target))
  }

  copyRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.copyRange(source, target))
  }

  moveRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.moveRange(source, target))
  }

  pasteRange(source: CellRangeRef, target: CellRangeRef): void {
    runEngineEffect(this.runtime.mutation.copyRange(source, target))
  }

  undo(): boolean {
    return runEngineEffect(this.runtime.history.undo())
  }

  redo(): boolean {
    return runEngineEffect(this.runtime.history.redo())
  }

  canUndo(): boolean {
    return this.undoStack.length > 0
  }

  canRedo(): boolean {
    return this.redoStack.length > 0
  }

  clearHistory(): void {
    this.undoStack.length = 0
    this.redoStack.length = 0
  }

  exportSheetCsv(sheetName: string): string {
    return runEngineEffect(this.runtime.read.exportSheetCsv(sheetName))
  }

  importSheetCsv(sheetName: string, csv: string, options?: CsvParseOptions): void {
    runEngineEffect(this.runtime.mutation.importSheetCsv(sheetName, csv, options))
  }

  getCellValue(sheetName: string, address: string): CellValue {
    return runEngineEffect(this.runtime.read.getCellValue(sheetName, address))
  }

  getRangeValues(range: CellRangeRef): CellValue[][] {
    return runEngineEffect(this.runtime.read.getRangeValues(range))
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    return runEngineEffect(this.runtime.read.getCell(sheetName, address))
  }

  getCellByIndex(cellIndex: number): CellSnapshot {
    return runEngineEffect(this.runtime.read.getCellByIndex(cellIndex))
  }

  getDependencies(sheetName: string, address: string): DependencySnapshot {
    return runEngineEffect(this.runtime.read.getDependencies(sheetName, address))
  }

  getDependents(sheetName: string, address: string): DependencySnapshot {
    return runEngineEffect(this.runtime.read.getDependents(sheetName, address))
  }

  explainCell(sheetName: string, address: string): ExplainCellSnapshot {
    return runEngineEffect(this.runtime.read.explainCell(sheetName, address))
  }

  exportSnapshot(options?: EngineSnapshotExportOptions): WorkbookSnapshot {
    return runEngineEffect(this.runtime.snapshot.exportWorkbook(options))
  }

  importSnapshot(snapshot: WorkbookSnapshot): void {
    runEngineEffect(this.runtime.snapshot.importWorkbook(snapshot))
  }

  exportReplicaSnapshot(): EngineReplicaSnapshot {
    return runEngineEffect(this.runtime.snapshot.exportReplica())
  }

  importReplicaSnapshot(snapshot: EngineReplicaSnapshot): void {
    runEngineEffect(this.runtime.snapshot.importReplica(snapshot))
  }

  renderCommit(ops: CommitOp[]): void {
    runEngineEffect(this.runtime.mutation.renderCommit(ops))
  }

  applyRemoteBatch(batch: EngineOpBatch): boolean {
    return runEngineEffect(this.runtime.sync.applyRemoteBatch(batch))
  }

  captureUndoOps<T>(mutate: () => T): {
    result: T
    undoOps: readonly EngineOp[] | null
  } {
    return runEngineEffect(this.runtime.mutation.captureUndoOps(mutate))
  }

  applyOps(
    ops: readonly EngineOp[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      trusted?: boolean
    } = {},
  ): readonly EngineOp[] | null {
    const nextOptions = options.source === undefined ? { ...options, source: 'local' as const } : options
    return this.runtime.mutation.applyOpsNow(ops, nextOptions)
  }
}
