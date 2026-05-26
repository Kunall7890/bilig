import { ErrorCode, ValueTag, type RecalcMetrics } from '@bilig/protocol'
import { formatAddress } from '@bilig/formula'
import type { EngineExistingNumericCellMutationRef, EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { makeCellEntity } from '../../entity-ids.js'
import type { EngineEventBus } from '../../events.js'
import type { FormulaTable } from '../../formula-table.js'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'
import type { SheetRecord, WorkbookStore } from '../../workbook-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { EngineTraversalService } from './traversal-service.js'
import { directScalarCellNumber, directScalarDeltaFromNumbers } from './direct-scalar-helpers.js'
import { makeCompactExistingNumericMutationResult, makeExistingNumericMutationResult } from './operation-change-helpers.js'
import { assertProtectionAllowsOp } from './operation-protection-helpers.js'
import { writeTrustedOperationExistingNumericLiteralToCell } from './operation-literal-write-helpers.js'
import { isWorkbookTableHeaderCell } from './operation-table-header-rename.js'

const DIRECT_SCALAR_EXISTING_NUMERIC_FAST_PATH_LIMIT = 4_096
const TRUSTED_EXISTING_NUMERIC_BLOCKED_FLAGS =
  CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput
const DIRECT_FORMULA_OUTPUT_FLAGS = CellFlags.SpillChild | CellFlags.PivotOutput

function appendChangedCellIndex(buffer: Uint32Array<ArrayBuffer>, count: number, cellIndex: number): Uint32Array<ArrayBuffer> {
  if (count < buffer.length) {
    buffer[count] = cellIndex
    return buffer
  }
  const next = new Uint32Array(buffer.length * 2)
  next.set(buffer)
  next[count] = cellIndex
  return next
}

function growFloat64(buffer: Float64Array<ArrayBuffer>, count: number, value: number): Float64Array<ArrayBuffer> {
  if (count < buffer.length) {
    buffer[count] = value
    return buffer
  }
  const next = new Float64Array(buffer.length * 2)
  next.set(buffer)
  next[count] = value
  return next
}

interface TrackedColumnDependencyState {
  readonly size: number
}

export interface TrustedExistingNumericDirectScalarFastPath {
  readonly tryApply: (
    request: EngineExistingNumericCellMutationRef,
    oldNumericValue: number,
  ) => EngineExistingNumericCellMutationResult | null
}

export function createTrustedExistingNumericDirectScalarFastPath(args: {
  readonly workbook: WorkbookStore
  readonly formulas: FormulaTable<RuntimeFormula>
  readonly events: EngineEventBus
  readonly counters: EngineCounters
  readonly traversal: Pick<EngineTraversalService, 'getSingleEntityDependentNow' | 'getEntityDependentsNow' | 'getSingleCellDependentNow'>
  readonly deferKernelSync: (cellIndices: readonly number[] | Uint32Array) => void
  readonly hasVolatileFormulas: () => boolean
  readonly hasRegionFormulaSubscriptions: () => boolean
  readonly reverseExactLookupColumnEdges: TrackedColumnDependencyState
  readonly reverseSortedLookupColumnEdges: TrackedColumnDependencyState
  readonly reverseAggregateColumnEdges: TrackedColumnDependencyState
  readonly getLastMetrics: () => RecalcMetrics
  readonly setLastMetrics: (metrics: RecalcMetrics) => void
}): TrustedExistingNumericDirectScalarFastPath {
  const deferSingleKernelSync = (cellIndex: number): void => {
    args.deferKernelSync(Uint32Array.of(cellIndex))
  }

  const setSingleLiteralSkipMetrics = (): void => {
    const previousMetrics = args.getLastMetrics()
    args.setLastMetrics({
      batchId: previousMetrics.batchId + 1,
      changedInputCount: 1,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
    })
  }

  const canApplyTrustedDirectScalarFormulaDelta = (formulaCellIndex: number): boolean => {
    const formula = args.formulas.get(formulaCellIndex)
    if (
      formula?.directScalar === undefined ||
      formula.compiled.volatile ||
      formula.compiled.producesSpill ||
      formula.directLookup !== undefined ||
      formula.directAggregate !== undefined ||
      formula.directCriteria !== undefined
    ) {
      return false
    }
    const cellStore = args.workbook.cellStore
    return (
      cellStore.tags[formulaCellIndex] === ValueTag.Number &&
      (cellStore.stringIds[formulaCellIndex] ?? 0) === 0 &&
      ((cellStore.errors[formulaCellIndex] as ErrorCode | undefined) ?? ErrorCode.None) === ErrorCode.None &&
      ((cellStore.flags[formulaCellIndex] ?? 0) & (CellFlags.InCycle | DIRECT_FORMULA_OUTPUT_FLAGS)) === 0
    )
  }

  const applyTrustedDirectScalarFormulaDelta = (cellIndex: number, delta: number): void => {
    const cellStore = args.workbook.cellStore
    cellStore.numbers[cellIndex] = (cellStore.numbers[cellIndex] ?? 0) + delta
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  }

  const writeTrustedDirectScalarFormulaNumber = (cellIndex: number, value: number): void => {
    const cellStore = args.workbook.cellStore
    cellStore.numbers[cellIndex] = value
    cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  }

  const tryApplyFanoutMutation = (
    request: EngineExistingNumericCellMutationRef,
    sheet: SheetRecord,
    oldNumber: number,
    rootEntity: number,
  ): EngineExistingNumericCellMutationResult | null => {
    const dependents = args.traversal.getEntityDependentsNow(rootEntity)
    if (dependents.length === 0) {
      return null
    }

    const cellStore = args.workbook.cellStore
    const deltas = new Float64Array(dependents.length)
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (
        formulaCellIndex < 0 ||
        args.traversal.getSingleCellDependentNow(formulaCellIndex) !== -1 ||
        !canApplyTrustedDirectScalarFormulaDelta(formulaCellIndex)
      ) {
        return null
      }
      const formula = args.formulas.get(formulaCellIndex)!
      const delta = directScalarDeltaFromNumbers(
        formula.directScalar!,
        request.cellIndex,
        oldNumber,
        request.value,
        (dependencyCellIndex) => directScalarCellNumber(cellStore, dependencyCellIndex),
      )
      if (delta === undefined) {
        return null
      }
      deltas[index] = delta
    }

    writeTrustedOperationExistingNumericLiteralToCell({
      cellStore,
      cellIndex: request.cellIndex,
      sheet,
      col: request.col,
      value: request.value,
    })

    const changed = new Uint32Array(dependents.length + 1)
    changed[0] = request.cellIndex
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      applyTrustedDirectScalarFormulaDelta(formulaCellIndex, deltas[index]!)
      changed[index + 1] = formulaCellIndex
    }
    addEngineCounter(args.counters, 'directScalarDeltaApplications', dependents.length)
    addEngineCounter(args.counters, 'directScalarDeltaOnlyRecalcSkips')
    deferSingleKernelSync(request.cellIndex)
    setSingleLiteralSkipMetrics()
    return makeExistingNumericMutationResult(changed, 1)
  }

  const tryApplyChainMutation = (
    request: EngineExistingNumericCellMutationRef,
    sheet: SheetRecord,
    oldRootNumber: number,
    firstFormulaCellIndex: number,
  ): EngineExistingNumericCellMutationResult | null => {
    const cellStore = args.workbook.cellStore
    let changed = new Uint32Array(16)
    let nextValues = new Float64Array(16)
    let changedCount = 0
    let currentCellIndex = request.cellIndex
    let currentOldNumber = oldRootNumber
    let currentNewNumber = request.value
    let formulaCellIndex = firstFormulaCellIndex

    for (;;) {
      if (changedCount >= DIRECT_SCALAR_EXISTING_NUMERIC_FAST_PATH_LIMIT) {
        return null
      }
      if (formulaCellIndex < 0 || formulaCellIndex === request.cellIndex || !canApplyTrustedDirectScalarFormulaDelta(formulaCellIndex)) {
        return null
      }
      const formula = args.formulas.get(formulaCellIndex)!
      const delta = directScalarDeltaFromNumbers(
        formula.directScalar!,
        currentCellIndex,
        currentOldNumber,
        currentNewNumber,
        (dependencyCellIndex) => directScalarCellNumber(cellStore, dependencyCellIndex),
      )
      if (delta === undefined) {
        return null
      }
      const formulaOldNumber = cellStore.numbers[formulaCellIndex] ?? 0
      const formulaNewNumber = formulaOldNumber + delta
      changed = appendChangedCellIndex(changed, changedCount, formulaCellIndex)
      nextValues = growFloat64(nextValues, changedCount, formulaNewNumber)
      changedCount += 1

      const nextDependent = args.traversal.getSingleCellDependentNow(formulaCellIndex)
      if (nextDependent === -1) {
        break
      }
      currentCellIndex = formulaCellIndex
      currentOldNumber = formulaOldNumber
      currentNewNumber = formulaNewNumber
      formulaCellIndex = nextDependent
    }

    writeTrustedOperationExistingNumericLiteralToCell({
      cellStore,
      cellIndex: request.cellIndex,
      sheet,
      col: request.col,
      value: request.value,
    })

    const changedCellIndices = new Uint32Array(changedCount + 1)
    changedCellIndices[0] = request.cellIndex
    for (let index = 0; index < changedCount; index += 1) {
      const changedCellIndex = changed[index]!
      writeTrustedDirectScalarFormulaNumber(changedCellIndex, nextValues[index]!)
      changedCellIndices[index + 1] = changedCellIndex
    }
    addEngineCounter(args.counters, 'directScalarDeltaApplications', changedCount)
    addEngineCounter(args.counters, 'directScalarDeltaOnlyRecalcSkips')
    deferSingleKernelSync(request.cellIndex)
    setSingleLiteralSkipMetrics()
    if (changedCount === 1) {
      const onlyFormulaCellIndex = changed[0]!
      return makeCompactExistingNumericMutationResult(
        request.cellIndex,
        onlyFormulaCellIndex,
        1,
        cellStore.numbers[onlyFormulaCellIndex],
        cellStore.rows[onlyFormulaCellIndex] ?? 0,
        cellStore.cols[onlyFormulaCellIndex] ?? 0,
      )
    }
    return makeExistingNumericMutationResult(changedCellIndices, 1)
  }

  const tryApply = (
    request: EngineExistingNumericCellMutationRef,
    oldNumericValue: number,
  ): EngineExistingNumericCellMutationResult | null => {
    if (
      request.trustedExistingNumericLiteral !== true ||
      request.emitTracked !== false ||
      Object.is(request.value, -0) ||
      args.workbook.hasPivots() ||
      args.events.hasListeners() ||
      args.events.hasTrackedListeners() ||
      args.events.hasCellListeners() ||
      args.hasVolatileFormulas() ||
      args.hasRegionFormulaSubscriptions() ||
      args.reverseExactLookupColumnEdges.size > 0 ||
      args.reverseSortedLookupColumnEdges.size > 0 ||
      args.reverseAggregateColumnEdges.size > 0
    ) {
      return null
    }

    const sheet = args.workbook.getSheetById(request.sheetId)
    const cellStore = args.workbook.cellStore
    const cellIndex = request.cellIndex
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      cellStore.sheetIds[cellIndex] !== request.sheetId ||
      cellStore.rows[cellIndex] !== request.row ||
      cellStore.cols[cellIndex] !== request.col ||
      cellStore.tags[cellIndex] !== ValueTag.Number ||
      args.formulas.get(cellIndex) !== undefined ||
      ((cellStore.flags[cellIndex] ?? 0) & TRUSTED_EXISTING_NUMERIC_BLOCKED_FLAGS) !== 0 ||
      isWorkbookTableHeaderCell(args.workbook, sheet.name, request.row, request.col)
    ) {
      return null
    }

    assertProtectionAllowsOp(args.workbook, {
      kind: 'setCellValue',
      sheetName: sheet.name,
      address: formatAddress(request.row, request.col),
      value: request.value,
    })

    const rootEntity = makeCellEntity(cellIndex)
    const singleDependent = args.traversal.getSingleEntityDependentNow(rootEntity)
    if (singleDependent === -1) {
      writeTrustedOperationExistingNumericLiteralToCell({
        cellStore,
        cellIndex,
        sheet,
        col: request.col,
        value: request.value,
      })
      deferSingleKernelSync(cellIndex)
      setSingleLiteralSkipMetrics()
      return makeCompactExistingNumericMutationResult(cellIndex, undefined, 1)
    }

    const oldNumber = Object.is(oldNumericValue, -0) ? 0 : oldNumericValue
    if (singleDependent >= 0) {
      return tryApplyChainMutation(request, sheet, oldNumber, singleDependent)
    }
    return tryApplyFanoutMutation(request, sheet, oldNumber, rootEntity)
  }

  return { tryApply }
}
