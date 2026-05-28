import { ErrorCode, ValueTag, type RecalcMetrics } from '@bilig/protocol'
import { formatAddress } from '@bilig/formula'
import type {
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
  EngineExistingNumericCellMutationsRef,
} from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { makeCellEntity } from '../../entity-ids.js'
import type { EngineEventBus } from '../../events.js'
import type { FormulaTable } from '../../formula-table.js'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'
import type { SheetRecord, WorkbookStore } from '../../workbook-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { EngineTraversalService } from './traversal-service.js'
import {
  directScalarCellNumber,
  directScalarDeltaFromNumbers,
  evaluateRowPairDirectScalarCode,
  rowPairDirectScalarCode,
  writeSingleInputAffineDirectScalar,
} from './direct-scalar-helpers.js'
import {
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  tagTrustedPhysicalTrackedChanges,
} from './operation-change-helpers.js'
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
  readonly tryApplyBatch: (request: EngineExistingNumericCellMutationsRef) => Float64Array | null
}

export function createTrustedExistingNumericDirectScalarFastPath(args: {
  readonly workbook: WorkbookStore
  readonly formulas: FormulaTable<RuntimeFormula>
  readonly events: EngineEventBus
  readonly counters: EngineCounters
  readonly traversal: Pick<EngineTraversalService, 'getSingleEntityDependentNow' | 'getEntityDependentsNow' | 'getSingleCellDependentNow'>
  readonly directScalarDeltaOutputCellIndicesByInput?: ArrayLike<number | undefined>
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

  const setLiteralSkipMetrics = (changedInputCount: number): RecalcMetrics => {
    const previousMetrics = args.getLastMetrics()
    const nextMetrics = {
      batchId: previousMetrics.batchId + 1,
      changedInputCount,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0,
    }
    args.setLastMetrics(nextMetrics)
    return nextMetrics
  }

  const singleDirectScalarDependent = (inputCellIndex: number): number => {
    const indexedDependent = args.directScalarDeltaOutputCellIndicesByInput?.[inputCellIndex]
    return indexedDependent !== undefined && indexedDependent >= 0
      ? indexedDependent
      : args.traversal.getSingleEntityDependentNow(makeCellEntity(inputCellIndex))
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
    setLiteralSkipMetrics(1)
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
    setLiteralSkipMetrics(1)
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
      setLiteralSkipMetrics(1)
      return makeCompactExistingNumericMutationResult(cellIndex, undefined, 1)
    }

    const oldNumber = Object.is(oldNumericValue, -0) ? 0 : oldNumericValue
    if (singleDependent >= 0) {
      return tryApplyChainMutation(request, sheet, oldNumber, singleDependent)
    }
    return tryApplyFanoutMutation(request, sheet, oldNumber, rootEntity)
  }

  const tryApplyRowPairBatch = (request: EngineExistingNumericCellMutationsRef, count: number): Float64Array | null => {
    if (count % 2 !== 0) {
      return null
    }
    const firstSheetId = request.sheetIds[0]!
    const firstRow = request.rows[0]!
    const firstCol = request.cols[0]!
    const secondSheetId = request.sheetIds[1]!
    const secondRow = request.rows[1]!
    const secondCol = request.cols[1]!
    if (firstSheetId !== secondSheetId || firstRow !== secondRow || firstCol >= secondCol) {
      return null
    }
    const sheet = args.workbook.getSheetById(firstSheetId)
    if (!sheet || sheet.structureVersion !== 1) {
      return null
    }

    const hasProtectionMetadata = args.workbook.hasProtectionMetadataForSheet(sheet.name)
    const cellStore = args.workbook.cellStore
    const changedCellIndices = new Uint32Array(count * 2)
    const inputCellIndices = changedCellIndices.subarray(0, count)
    const formulaCellIndices = changedCellIndices.subarray(count)
    const formulaNumericResults = new Float64Array(count)
    const oldNumbers = new Float64Array(count)
    let formulaCount = 0
    let previousRow = firstRow - 1
    let previousFormulaRow = -1
    let previousFormulaCol = -1

    for (let refIndex = 0; refIndex < count; refIndex += 2) {
      const leftSheetId = request.sheetIds[refIndex]!
      const rightSheetId = request.sheetIds[refIndex + 1]!
      const leftRow = request.rows[refIndex]!
      const rightRow = request.rows[refIndex + 1]!
      const leftCol = request.cols[refIndex]!
      const rightCol = request.cols[refIndex + 1]!
      const leftValue = request.numbers[refIndex]!
      const rightValue = request.numbers[refIndex + 1]!
      const leftCellIndexPlusOne = request.cellIndexPlusOnes[refIndex]!
      const rightCellIndexPlusOne = request.cellIndexPlusOnes[refIndex + 1]!
      if (
        leftSheetId !== firstSheetId ||
        rightSheetId !== firstSheetId ||
        leftRow !== rightRow ||
        leftRow <= previousRow ||
        leftCol !== firstCol ||
        rightCol !== secondCol ||
        Object.is(leftValue, -0) ||
        Object.is(rightValue, -0) ||
        leftCellIndexPlusOne === 0 ||
        rightCellIndexPlusOne === 0
      ) {
        return null
      }
      previousRow = leftRow
      const leftIndex = leftCellIndexPlusOne - 1
      const rightIndex = rightCellIndexPlusOne - 1
      if (
        cellStore.sheetIds[leftIndex] !== leftSheetId ||
        cellStore.rows[leftIndex] !== leftRow ||
        cellStore.cols[leftIndex] !== leftCol ||
        cellStore.sheetIds[rightIndex] !== rightSheetId ||
        cellStore.rows[rightIndex] !== rightRow ||
        cellStore.cols[rightIndex] !== rightCol ||
        cellStore.tags[leftIndex] !== ValueTag.Number ||
        cellStore.tags[rightIndex] !== ValueTag.Number ||
        args.formulas.get(leftIndex) !== undefined ||
        args.formulas.get(rightIndex) !== undefined ||
        ((cellStore.flags[leftIndex] ?? 0) & TRUSTED_EXISTING_NUMERIC_BLOCKED_FLAGS) !== 0 ||
        ((cellStore.flags[rightIndex] ?? 0) & TRUSTED_EXISTING_NUMERIC_BLOCKED_FLAGS) !== 0 ||
        isWorkbookTableHeaderCell(args.workbook, sheet.name, leftRow, leftCol) ||
        isWorkbookTableHeaderCell(args.workbook, sheet.name, rightRow, rightCol)
      ) {
        return null
      }
      if (hasProtectionMetadata) {
        assertProtectionAllowsOp(args.workbook, {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address: formatAddress(leftRow, leftCol),
          value: leftValue,
        })
        assertProtectionAllowsOp(args.workbook, {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address: formatAddress(rightRow, rightCol),
          value: rightValue,
        })
      }

      const rowFormulaStart = formulaCount
      const considerDependent = (formulaCellIndex: number): boolean => {
        for (let index = rowFormulaStart; index < formulaCount; index += 1) {
          if (formulaCellIndices[index] === formulaCellIndex) {
            return true
          }
        }
        const formula = args.formulas.get(formulaCellIndex)
        if (
          formula?.directScalar === undefined ||
          !canApplyTrustedDirectScalarFormulaDelta(formulaCellIndex) ||
          args.traversal.getSingleCellDependentNow(formulaCellIndex) !== -1 ||
          cellStore.sheetIds[formulaCellIndex] !== firstSheetId ||
          cellStore.rows[formulaCellIndex] !== leftRow
        ) {
          return false
        }
        const formulaRow = cellStore.rows[formulaCellIndex] ?? 0
        const formulaCol = cellStore.cols[formulaCellIndex] ?? 0
        if (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol)) {
          return false
        }
        const code = rowPairDirectScalarCode(formula.directScalar, leftIndex, rightIndex)
        const result = code === 0 ? undefined : evaluateRowPairDirectScalarCode(code, leftValue, rightValue)
        if (result === undefined || formulaCount >= formulaCellIndices.length) {
          return false
        }
        formulaCellIndices[formulaCount] = formulaCellIndex
        formulaNumericResults[formulaCount] = result
        formulaCount += 1
        previousFormulaRow = formulaRow
        previousFormulaCol = formulaCol
        return true
      }

      const leftDependents = args.traversal.getEntityDependentsNow(makeCellEntity(leftIndex))
      const rightDependents = args.traversal.getEntityDependentsNow(makeCellEntity(rightIndex))
      if (leftDependents.length === 0 || rightDependents.length === 0) {
        return null
      }
      for (let index = 0; index < leftDependents.length; index += 1) {
        if (!considerDependent(leftDependents[index]!)) {
          return null
        }
      }
      if (!sameDependentOrder(leftDependents, rightDependents)) {
        for (let index = 0; index < rightDependents.length; index += 1) {
          if (!considerDependent(rightDependents[index]!)) {
            return null
          }
        }
      }
      if (formulaCount !== rowFormulaStart + 2) {
        return null
      }
      inputCellIndices[refIndex] = leftIndex
      inputCellIndices[refIndex + 1] = rightIndex
      oldNumbers[refIndex] = cellStore.numbers[leftIndex] ?? 0
      oldNumbers[refIndex + 1] = cellStore.numbers[rightIndex] ?? 0
    }

    if (formulaCount !== count) {
      return null
    }

    const flags = cellStore.flags
    const versions = cellStore.versions
    const stringIds = cellStore.stringIds
    const tags = cellStore.tags
    const numbers = cellStore.numbers
    const errors = cellStore.errors
    const clearAuthoredBlankFlag = ~CellFlags.AuthoredBlank
    for (let index = 0; index < count; index += 1) {
      const cellIndex = inputCellIndices[index]!
      const currentFlags = flags[cellIndex] ?? 0
      if ((currentFlags & CellFlags.AuthoredBlank) !== 0) {
        flags[cellIndex] = currentFlags & clearAuthoredBlankFlag
      }
      tags[cellIndex] = ValueTag.Number
      errors[cellIndex] = ErrorCode.None
      stringIds[cellIndex] = 0
      numbers[cellIndex] = request.numbers[index]!
      versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
    }
    for (let index = 0; index < formulaCount; index += 1) {
      const formulaCellIndex = formulaCellIndices[index]!
      numbers[formulaCellIndex] = formulaNumericResults[index]!
      versions[formulaCellIndex] = (versions[formulaCellIndex] ?? 0) + 1
    }

    args.workbook.notifyColumnPairWritten(firstSheetId, firstCol, secondCol)
    args.deferKernelSync(inputCellIndices)
    addEngineCounter(args.counters, 'directScalarDeltaApplications', formulaCount)
    addEngineCounter(args.counters, 'directScalarDeltaOnlyRecalcSkips')
    const metrics = setLiteralSkipMetrics(count)
    if (args.events.hasTrackedListeners()) {
      tagTrustedPhysicalTrackedChanges(changedCellIndices, firstSheetId, count)
      args.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics,
        explicitChangedCount: count,
      })
    }
    return oldNumbers
  }

  const tryApplyBatch = (request: EngineExistingNumericCellMutationsRef): Float64Array | null => {
    const count = request.sheetIds.length
    if (
      count < 32 ||
      request.cellIndexPlusOnes.length !== count ||
      request.rows.length !== count ||
      request.cols.length !== count ||
      request.numbers.length !== count ||
      (request.potentialNewCells ?? 0) !== 0 ||
      args.workbook.hasPivots() ||
      args.events.hasListeners() ||
      args.events.hasCellListeners() ||
      args.hasVolatileFormulas() ||
      args.hasRegionFormulaSubscriptions() ||
      args.reverseExactLookupColumnEdges.size > 0 ||
      args.reverseSortedLookupColumnEdges.size > 0 ||
      args.reverseAggregateColumnEdges.size > 0
    ) {
      return null
    }
    const hasTrackedEventListeners = args.events.hasTrackedListeners()

    const firstSheetId = request.sheetIds[0]!
    const firstRow = request.rows[0]!
    const firstCol = request.cols[0]!
    const secondRow = request.rows[1]!
    const secondCol = request.cols[1]!
    if (secondRow === firstRow && firstCol < secondCol) {
      return tryApplyRowPairBatch(request, count)
    }
    if (secondRow <= firstRow) {
      return null
    }
    const sheet = args.workbook.getSheetById(firstSheetId)
    if (!sheet || sheet.structureVersion !== 1) {
      return null
    }
    const hasProtectionMetadata = args.workbook.hasProtectionMetadataForSheet(sheet.name)

    const cellStore = args.workbook.cellStore
    const changedCellIndices = new Uint32Array(count * 2)
    const inputCellIndices = changedCellIndices.subarray(0, count)
    const formulaCellIndices = changedCellIndices.subarray(count)
    const oldNumbers = new Float64Array(count)
    let previousRow = firstRow
    let affineScale: number | undefined
    let affineOffset: number | undefined
    const affine = { scale: 0, offset: 0 }

    for (let index = 0; index < count; index += 1) {
      const sheetId = request.sheetIds[index]!
      const row = request.rows[index]!
      const col = request.cols[index]!
      const cellIndexPlusOne = request.cellIndexPlusOnes[index]!
      const value = request.numbers[index]!
      if (
        sheetId !== firstSheetId ||
        col !== firstCol ||
        cellIndexPlusOne === 0 ||
        Object.is(value, -0) ||
        (index > 0 && row <= previousRow)
      ) {
        return null
      }
      previousRow = row
      const cellIndex = cellIndexPlusOne - 1
      if (
        cellStore.sheetIds[cellIndex] !== sheetId ||
        cellStore.rows[cellIndex] !== row ||
        cellStore.cols[cellIndex] !== col ||
        cellStore.tags[cellIndex] !== ValueTag.Number ||
        args.formulas.get(cellIndex) !== undefined ||
        ((cellStore.flags[cellIndex] ?? 0) & TRUSTED_EXISTING_NUMERIC_BLOCKED_FLAGS) !== 0 ||
        isWorkbookTableHeaderCell(args.workbook, sheet.name, row, col)
      ) {
        return null
      }
      if (hasProtectionMetadata) {
        assertProtectionAllowsOp(args.workbook, {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address: formatAddress(row, col),
          value,
        })
      }

      const formulaCellIndex = singleDirectScalarDependent(cellIndex)
      if (formulaCellIndex < 0 || args.traversal.getSingleCellDependentNow(formulaCellIndex) !== -1) {
        return null
      }
      const formula = args.formulas.get(formulaCellIndex)
      if (
        formula?.directScalar === undefined ||
        !canApplyTrustedDirectScalarFormulaDelta(formulaCellIndex) ||
        cellStore.sheetIds[formulaCellIndex] !== firstSheetId ||
        cellStore.rows[formulaCellIndex] !== row ||
        !writeSingleInputAffineDirectScalar(formula.directScalar, cellIndex, affine)
      ) {
        return null
      }
      if (affineScale === undefined) {
        affineScale = affine.scale
        affineOffset = affine.offset
      } else if (!Object.is(affineScale, affine.scale) || !Object.is(affineOffset, affine.offset)) {
        return null
      }

      inputCellIndices[index] = cellIndex
      formulaCellIndices[index] = formulaCellIndex
      oldNumbers[index] = cellStore.numbers[cellIndex] ?? 0
    }

    if (affineScale === undefined || affineOffset === undefined) {
      return null
    }

    const flags = cellStore.flags
    const versions = cellStore.versions
    const stringIds = cellStore.stringIds
    const tags = cellStore.tags
    const numbers = cellStore.numbers
    const errors = cellStore.errors
    const formulaOutputFlags = CellFlags.SpillChild | CellFlags.PivotOutput
    const clearFormulaOutputFlags = ~formulaOutputFlags
    for (let index = 0; index < count; index += 1) {
      const inputCellIndex = inputCellIndices[index]!
      const value = request.numbers[index]!
      const inputFlags = flags[inputCellIndex] ?? 0
      if ((inputFlags & formulaOutputFlags) !== 0) {
        flags[inputCellIndex] = inputFlags & clearFormulaOutputFlags
      }
      versions[inputCellIndex] = (versions[inputCellIndex] ?? 0) + 1
      stringIds[inputCellIndex] = 0
      tags[inputCellIndex] = ValueTag.Number
      numbers[inputCellIndex] = value
      errors[inputCellIndex] = ErrorCode.None

      const formulaCellIndex = formulaCellIndices[index]!
      const formulaFlags = flags[formulaCellIndex] ?? 0
      if ((formulaFlags & formulaOutputFlags) !== 0) {
        flags[formulaCellIndex] = formulaFlags & clearFormulaOutputFlags
      }
      versions[formulaCellIndex] = (versions[formulaCellIndex] ?? 0) + 1
      stringIds[formulaCellIndex] = 0
      tags[formulaCellIndex] = ValueTag.Number
      numbers[formulaCellIndex] = value * affineScale + affineOffset
      errors[formulaCellIndex] = ErrorCode.None
    }
    args.workbook.notifyColumnWritten(firstSheetId, firstCol)
    args.deferKernelSync(inputCellIndices)
    addEngineCounter(args.counters, 'directScalarDeltaApplications', count)
    addEngineCounter(args.counters, 'directScalarDeltaOnlyRecalcSkips')
    const metrics = setLiteralSkipMetrics(count)
    if (hasTrackedEventListeners) {
      tagTrustedPhysicalTrackedChanges(changedCellIndices, firstSheetId, count)
      args.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics,
        explicitChangedCount: count,
      })
    }
    return oldNumbers
  }

  return { tryApply, tryApplyBatch }
}

function sameDependentOrder(left: Uint32Array, right: Uint32Array): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}
