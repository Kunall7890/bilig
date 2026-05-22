import { ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook'
import type {
  EngineCellMutationRef,
  EngineExistingLiteralCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
} from '../../cell-mutations-at.js'
import { writeLiteralToCellStore } from '../../engine-value-utils.js'
import { isRangeEntity, makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { U32 } from '../runtime-state.js'
import { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import {
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  hasCompleteDirectFormulaDeltas,
} from './direct-formula-recalc-helpers.js'
import { directScalarLiteralNumericValue } from './direct-scalar-helpers.js'
import { hasOperationCompactedRangeDependencies } from './operation-cell-lifecycle-helpers.js'
import {
  canTrustPhysicalTrackedChangeSplit,
  makeExistingNumericMutationResult,
  tagTrustedPhysicalTrackedChanges,
} from './operation-change-helpers.js'
import { recordKernelSyncOnlyLiteralChange } from './operation-kernel-sync-only-literal-change.js'
import { applyOperationLookupNumericWriteTailPatches, planOperationLookupNumericWrites } from './operation-lookup-write-plans.js'
import { tryApplySinglePostRecalcDirectFormula, type DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import type { MutationSource, OperationSingleExistingLiteralFastPathArgs } from './operation-single-existing-literal-fast-path-types.js'
import { isWorkbookTableHeaderCell } from './operation-table-header-rename.js'

const DIRECT_RANGE_POST_RECALC_LIMIT = 16_384
const EMPTY_CHANGED_CELLS = new Uint32Array(0)
const FORMULA_LEAF_DEPENDENCY_SCAN_LIMIT = 128

export function createOperationSingleExistingLiteralFastPath(args: OperationSingleExistingLiteralFastPathArgs): {
  readonly tryApplySingleExistingDirectLiteralMutation: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: MutationSource,
  ) => boolean
  readonly applyExistingNumericCellMutationAtNow: (
    request: EngineExistingNumericCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly applyExistingLiteralCellMutationAtNow: (
    request: EngineExistingLiteralCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
} {
  const {
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    canSkipApproximateLookupNewNumericColumnWrite,
    writeNumericLiteralToExistingCell,
    deferSingleCellKernelSync,
    makeSingleLiteralSkipMetrics,
    canFastPathLiteralOverwrite,
    directScalarCellNumericValue,
    tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation,
    tryApplyTrustedColumnDirectAggregateExistingNumericMutation,
    tryApplyTrustedDirectScalarClosureExistingNumericMutation,
    tryApplyTrustedFormulaLeafExistingNumericMutation,
    tryApplyFormulaLeafExistingLiteralMutation,
    tryApplySingleDirectAggregateLiteralMutationFastPath,
    planExactLookupNumericColumnWrite,
    planApproximateLookupNumericColumnWrite,
    patchUniformLookupTailWrites,
    tryApplySingleKernelSyncOnlyLiteralMutationFastPath,
    tryApplySingleDirectFormulaLiteralMutationWithoutEvents,
    tryApplySingleDirectScalarLiteralMutationWithoutEvents,
    tryApplySingleDirectScalarLiteralMutationWithoutEventsAndReturnChanged,
    tryApplySingleDirectLookupOperandMutationFastPath,
    markPostRecalcDirectScalarNumericDependents,
    tryMarkDirectScalarLinearDeltaClosure,
    collectSingleAffectedDirectRangeDependent,
    collectAffectedDirectRangeDependents,
    applyDirectFormulaCurrentResult,
    applyDirectFormulaNumericDelta,
    applyDirectScalarCurrentValue,
    tryApplyDirectScalarDeltas,
    tryApplyDirectFormulaDeltas,
    countPostRecalcDirectFormulaMetric,
    hasDynamicFormulaDependents,
  } = args

  const hasKnownDynamicFormulaDependents = (
    sheetId: number,
    row: number,
    col: number,
    existingIndex: number,
    singleExistingCellDependent: number,
  ): boolean => {
    if (singleExistingCellDependent === -2) {
      return hasDynamicFormulaDependents(existingIndex)
    }
    if (singleExistingCellDependent >= 0 && !isRangeEntity(singleExistingCellDependent) && args.state.formulas.size === 1) {
      const formula = args.state.formulas.get(singleExistingCellDependent)
      return formula !== undefined && hasOperationCompactedRangeDependencies(formula)
    }
    const singleRegionFormulaDependent = args.collectSingleRegionFormulaDependentForCellAt?.(sheetId, row, col)
    if (singleRegionFormulaDependent === undefined || singleRegionFormulaDependent === -2) {
      return hasDynamicFormulaDependents(existingIndex)
    }

    let firstFormulaCellIndex = -1
    let secondFormulaCellIndex = -1
    const pushFormulaCellIndex = (candidate: number): void => {
      if (candidate < 0 || isRangeEntity(candidate)) {
        return
      }
      if (firstFormulaCellIndex === -1) {
        firstFormulaCellIndex = candidate
        return
      }
      if (candidate !== firstFormulaCellIndex) {
        secondFormulaCellIndex = candidate
      }
    }

    pushFormulaCellIndex(singleExistingCellDependent)
    pushFormulaCellIndex(singleRegionFormulaDependent)
    if (firstFormulaCellIndex === -1) {
      return false
    }
    const firstFormula = args.state.formulas.get(firstFormulaCellIndex)
    if (firstFormula !== undefined && hasOperationCompactedRangeDependencies(firstFormula)) {
      return true
    }
    const secondFormula = secondFormulaCellIndex === -1 ? undefined : args.state.formulas.get(secondFormulaCellIndex)
    return secondFormula !== undefined && hasOperationCompactedRangeDependencies(secondFormula)
  }

  const collectSingleFormulaLeafRangeDependent = (
    existingIndex: number,
    singleExistingCellDependent: number,
    sheetId: number,
    sheetName: string,
    row: number,
    col: number,
  ): number => {
    if (isRangeEntity(singleExistingCellDependent)) {
      const rangeFormulaDependent = args.getSingleEntityDependent(singleExistingCellDependent)
      if (rangeFormulaDependent !== -1) {
        return rangeFormulaDependent
      }
    }
    const indexedSingle =
      args.collectSingleRegionFormulaDependentForCellAt?.(sheetId, row, col) ??
      args.collectSingleRegionFormulaDependentForCell(sheetName, row, col)
    if (indexedSingle >= 0 && !isRangeEntity(indexedSingle)) {
      return indexedSingle
    }
    const dependents = args.collectRegionFormulaDependentsForCell(sheetName, row, col)
    let singleFormulaCellIndex = -1
    for (let index = 0; index < dependents.length; index += 1) {
      const candidate = dependents[index]!
      if (candidate < 0 || isRangeEntity(candidate)) {
        continue
      }
      if (singleFormulaCellIndex !== -1 && singleFormulaCellIndex !== candidate) {
        return -2
      }
      singleFormulaCellIndex = candidate
    }
    if (singleFormulaCellIndex !== -1 || args.state.formulas.size > FORMULA_LEAF_DEPENDENCY_SCAN_LIMIT) {
      return singleFormulaCellIndex
    }
    for (const [formulaCellIndex, formula] of args.state.formulas.entries()) {
      const dependencyIndices = formula.dependencyIndices
      for (let index = 0; index < dependencyIndices.length; index += 1) {
        if (dependencyIndices[index] !== existingIndex) {
          continue
        }
        if (singleFormulaCellIndex !== -1 && singleFormulaCellIndex !== formulaCellIndex) {
          return -2
        }
        singleFormulaCellIndex = formulaCellIndex
        break
      }
    }
    return singleFormulaCellIndex !== -1 ? singleFormulaCellIndex : indexedSingle === -2 ? -2 : -1
  }

  const tryApplySingleExistingDirectLiteralMutation = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
  ): boolean => {
    if (
      source !== 'local' ||
      batch !== null ||
      refs.length !== 1 ||
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners()
    ) {
      return false
    }
    if (args.hasVolatileFormulas?.()) {
      return false
    }
    const ref = refs[0]!
    const mutation = ref.mutation
    if (mutation.kind !== 'setCellValue' || mutation.value === null) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(ref.sheetId)
    if (!sheet || sheet.structureVersion !== 1) {
      return false
    }
    if (isWorkbookTableHeaderCell(args.state.workbook, sheet.name, mutation.row, mutation.col)) {
      return false
    }
    const existingIndex =
      ref.cellIndex !== undefined &&
      args.state.workbook.cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
      args.state.workbook.cellStore.rows[ref.cellIndex] === mutation.row &&
      args.state.workbook.cellStore.cols[ref.cellIndex] === mutation.col
        ? ref.cellIndex
        : sheet.grid.getPhysical(mutation.row, mutation.col)
    const sheetName = sheet.name
    const hasTrackedColumnDependents = args.hasTrackedColumnDependentsAnywhere()
    const hasExactLookupDependents = hasTrackedColumnDependents && hasTrackedExactLookupDependents(ref.sheetId, mutation.col)
    const hasSortedLookupDependents = hasTrackedColumnDependents && hasTrackedSortedLookupDependents(ref.sheetId, mutation.col)
    const hasAggregateDependents = hasTrackedColumnDependents && hasTrackedDirectRangeDependents(ref.sheetId, mutation.col)
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    if (existingIndex === -1) {
      if (
        args.state.trackReplicaVersions ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0) ||
        hasExactLookupDependents ||
        hasAggregateDependents ||
        (hasSortedLookupDependents && !canSkipApproximateLookupNewNumericColumnWrite(ref.sheetId, mutation.col, mutation.row))
      ) {
        return false
      }
      const cellIndex = args.state.workbook.ensureCellAt(ref.sheetId, mutation.row, mutation.col).cellIndex
      writeNumericLiteralToExistingCell(cellIndex, mutation.value)
      recordKernelSyncOnlyLiteralChange({
        state: args.state,
        cellIndex,
        deferSingleCellKernelSync,
        makeSingleLiteralSkipMetrics,
        emitTracked: hasTrackedEventListeners,
      })
      return true
    }
    if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
      return false
    }
    const singleExistingCellDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
    const oldNumber = directScalarCellNumericValue(existingIndex)
    const newNumber = directScalarLiteralNumericValue(mutation.value)
    if (oldNumber === undefined || newNumber === undefined) {
      if (hasKnownDynamicFormulaDependents(ref.sheetId, mutation.row, mutation.col, existingIndex, singleExistingCellDependent)) {
        return false
      }
      const directLookupResult =
        !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents
          ? tryApplySingleDirectLookupOperandMutationFastPath({
              existingIndex,
              formulaCellIndex: singleExistingCellDependent,
              value: mutation.value,
              exactLookupValue: undefined,
              approximateLookupValue: undefined,
              emitTracked: hasTrackedEventListeners,
              lookupSheetHint: sheet,
            })
          : null
      if (directLookupResult) {
        return true
      }
      const formulaLeafResult =
        !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents
          ? tryApplyFormulaLeafExistingLiteralMutation({
              existingIndex,
              formulaCellIndex: singleExistingCellDependent,
              value: mutation.value,
              hasTrackedEventListeners,
            })
          : null
      return formulaLeafResult !== null
    }

    if (
      hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      (singleExistingCellDependent === -1 || isRangeEntity(singleExistingCellDependent)) &&
      tryApplySingleDirectAggregateLiteralMutationFastPath({
        existingIndex,
        sheetId: ref.sheetId,
        sheetName,
        row: mutation.row,
        col: mutation.col,
        value: mutation.value,
        delta: newNumber - oldNumber,
        emitTracked: hasTrackedEventListeners,
        ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
      })
    ) {
      return true
    }
    const existingTag = (args.state.workbook.cellStore.tags[existingIndex] as ValueTag | undefined) ?? ValueTag.Empty
    const mutationIsNumber = typeof mutation.value === 'number'
    const directLookupExactMutationNumber = mutationIsNumber ? newNumber : undefined
    const directLookupApproximateMutationNumber = newNumber
    const oldExactLookupNumber = hasExactLookupDependents && existingTag === ValueTag.Number ? oldNumber : undefined
    const newExactLookupNumber = hasExactLookupDependents && mutationIsNumber ? newNumber : undefined
    const oldApproximateLookupNumber = hasSortedLookupDependents ? oldNumber : undefined
    const newApproximateLookupNumber = hasSortedLookupDependents ? newNumber : undefined
    const lookupWritePlans = planOperationLookupNumericWrites({
      isRestore: false,
      hasAggregateDependents,
      hasExactLookupDependents,
      hasSortedLookupDependents,
      sheetId: ref.sheetId,
      sheetName,
      row: mutation.row,
      col: mutation.col,
      oldExactLookupNumber,
      newExactLookupNumber,
      oldApproximateLookupNumber,
      newApproximateLookupNumber,
      planner: {
        planExactLookupNumericColumnWrite,
        planApproximateLookupNumericColumnWrite,
      },
    })
    const exactLookupDependentsHandled = hasExactLookupDependents && lookupWritePlans.exactHandled
    const sortedLookupDependentsHandled = hasSortedLookupDependents && lookupWritePlans.sortedHandled
    if ((hasExactLookupDependents && !exactLookupDependentsHandled) || (hasSortedLookupDependents && !sortedLookupDependentsHandled)) {
      return false
    }

    const lookupDependentsHandled =
      (hasExactLookupDependents && exactLookupDependentsHandled) || (hasSortedLookupDependents && sortedLookupDependentsHandled)
    const canUseNumericLookupWriteFastPath = lookupDependentsHandled && existingTag === ValueTag.Number && mutationIsNumber
    if (!hasAggregateDependents && (hasExactLookupDependents || hasSortedLookupDependents) && singleExistingCellDependent === -1) {
      if (canUseNumericLookupWriteFastPath) {
        writeNumericLiteralToExistingCell(existingIndex, newNumber)
        const columnVersionAfterWrite = sheet.columnVersions[mutation.col] ?? 0
        const patchedLookupTailTargets = applyOperationLookupNumericWriteTailPatches(lookupWritePlans, {
          row: mutation.row,
          oldExactLookupNumber,
          newExactLookupNumber,
          oldApproximateLookupNumber,
          newApproximateLookupNumber,
          columnVersionAfterWrite,
        })
        const needsExactPatch = hasExactLookupDependents && exactLookupDependentsHandled && !patchedLookupTailTargets.exact
        const needsSortedPatch = hasSortedLookupDependents && sortedLookupDependentsHandled && !patchedLookupTailTargets.sorted
        const patchedLookupOwners =
          needsExactPatch || needsSortedPatch
            ? patchUniformLookupTailWrites({
                sheetId: ref.sheetId,
                col: mutation.col,
                row: mutation.row,
                oldNumeric: oldNumber,
                newNumeric: newNumber,
                exact: needsExactPatch,
                sorted: needsSortedPatch,
              })
            : { exact: true, sorted: true }
        if (needsExactPatch && !patchedLookupOwners.exact) {
          args.invalidateExactLookupColumn({ sheetName, col: mutation.col })
        }
        if (needsSortedPatch && !patchedLookupOwners.sorted) {
          args.invalidateSortedLookupColumn({ sheetName, col: mutation.col })
        }
        recordKernelSyncOnlyLiteralChange({
          state: args.state,
          cellIndex: existingIndex,
          deferSingleCellKernelSync,
          makeSingleLiteralSkipMetrics,
          emitTracked: hasTrackedEventListeners,
        })
        return true
      }
      if (
        tryApplySingleKernelSyncOnlyLiteralMutationFastPath({
          existingIndex,
          value: mutation.value,
          emitTracked: hasTrackedEventListeners,
        })
      ) {
        return true
      }
    }

    if (hasKnownDynamicFormulaDependents(ref.sheetId, mutation.row, mutation.col, existingIndex, singleExistingCellDependent)) {
      return false
    }

    if (!hasTrackedEventListeners && !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents) {
      if (
        tryApplySingleDirectFormulaLiteralMutationWithoutEvents({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          value: mutation.value,
          oldNumber,
          newNumber,
          exactLookupValue: directLookupExactMutationNumber,
          approximateLookupValue: directLookupApproximateMutationNumber,
        }) ||
        tryApplySingleDirectScalarLiteralMutationWithoutEvents({
          existingIndex,
          value: mutation.value,
          oldNumber,
          newNumber,
        })
      ) {
        return true
      }
    }
    if (!hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents) {
      if (
        tryApplySingleDirectLookupOperandMutationFastPath({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          value: mutation.value,
          exactLookupValue: directLookupExactMutationNumber,
          approximateLookupValue: directLookupApproximateMutationNumber,
          emitTracked: hasTrackedEventListeners,
          lookupSheetHint: sheet,
        })
      ) {
        return true
      }
    }
    const oldValue: CellValue = { tag: ValueTag.Number, value: oldNumber }
    const newValue: CellValue = { tag: ValueTag.Number, value: newNumber }
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
    let directDependentsHandled = markPostRecalcDirectScalarNumericDependents(
      existingIndex,
      oldNumber,
      newNumber,
      postRecalcDirectFormulaIndices,
      directLookupExactMutationNumber,
      directLookupApproximateMutationNumber,
    )
    if (
      !directDependentsHandled &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      !hasAggregateDependents &&
      tryMarkDirectScalarLinearDeltaClosure(existingIndex, oldValue, newValue, postRecalcDirectFormulaIndices)
    ) {
      directDependentsHandled = true
    }
    if (
      hasAggregateDependents &&
      directDependentsHandled &&
      postRecalcDirectFormulaIndices.size === 0 &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      tryApplySingleDirectAggregateLiteralMutationFastPath({
        existingIndex,
        sheetId: ref.sheetId,
        sheetName,
        row: mutation.row,
        col: mutation.col,
        value: mutation.value,
        delta: newNumber - oldNumber,
        emitTracked: hasTrackedEventListeners,
        ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
      })
    ) {
      return true
    }
    let shouldNoteAggregateLiteralWrite = false
    if (hasAggregateDependents) {
      if (!directDependentsHandled) {
        directDependentsHandled = true
      }
      const singleAffected = collectSingleAffectedDirectRangeDependent({
        sheetName,
        sheetId: ref.sheetId,
        row: mutation.row,
        col: mutation.col,
      })
      if (singleAffected >= 0) {
        const formula = args.state.formulas.get(singleAffected)
        if (
          !formula ||
          formula.directAggregate?.aggregateKind !== 'sum' ||
          formula.dependencyIndices.length !== 0 ||
          args.getSingleEntityDependent(makeCellEntity(singleAffected)) !== -1
        ) {
          return false
        }
        postRecalcDirectFormulaIndices.addDelta(singleAffected, newNumber - oldNumber)
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(existingIndex)
        shouldNoteAggregateLiteralWrite = true
      } else if (singleAffected === -2) {
        const affected = collectAffectedDirectRangeDependents({
          sheetName,
          row: mutation.row,
          col: mutation.col,
        })
        if (affected.length === 0 || affected.length > DIRECT_RANGE_POST_RECALC_LIMIT) {
          return false
        }
        for (let index = 0; index < affected.length; index += 1) {
          const formulaCellIndex = affected[index]!
          const formula = args.state.formulas.get(formulaCellIndex)
          if (
            !formula ||
            formula.directAggregate?.aggregateKind !== 'sum' ||
            formula.dependencyIndices.length !== 0 ||
            args.getSingleEntityDependent(makeCellEntity(formulaCellIndex)) !== -1
          ) {
            return false
          }
        }
        postRecalcDirectFormulaIndices.appendConstantDelta(affected, newNumber - oldNumber)
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(existingIndex)
        shouldNoteAggregateLiteralWrite = true
      }
    }
    if (
      !directDependentsHandled ||
      (postRecalcDirectFormulaIndices.size > 1 && !hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices))
    ) {
      return false
    }

    const explicitChangedCount = hasTrackedEventListeners ? 1 : 0
    const postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts = {
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    }
    let recalculated: U32 = EMPTY_CHANGED_CELLS
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      writeLiteralToCellStore(args.state.workbook.cellStore, existingIndex, mutation.value, args.state.strings)
      args.state.workbook.notifyCellValueWritten(existingIndex)
      if (shouldNoteAggregateLiteralWrite) {
        args.noteAggregateLiteralWrite({
          sheetName,
          row: mutation.row,
          col: mutation.col,
          oldValue,
          newValue,
        })
      }
      if (postRecalcDirectFormulaIndices.size > 0) {
        if (hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices)) {
          countDirectFormulaDeltaSkip(args.state.formulas, postRecalcDirectFormulaIndices, args.state.counters)
        } else if (canEvaluatePostRecalcDirectFormulasWithoutKernel(args.state.formulas, postRecalcDirectFormulaIndices)) {
          addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
        }
        const directChanged =
          tryApplyDirectScalarDeltas(postRecalcDirectFormulaIndices, hasTrackedEventListeners) ??
          tryApplyDirectFormulaDeltas(postRecalcDirectFormulaIndices, hasTrackedEventListeners) ??
          tryApplySinglePostRecalcDirectFormula(
            {
              state: args.state,
              collection: postRecalcDirectFormulaIndices,
              recalculated: EMPTY_CHANGED_CELLS,
              didRunRecalc: false,
              metrics: postRecalcDirectFormulaMetrics,
              applyDirectFormulaCurrentResult,
              applyDirectFormulaNumericDelta,
              applyDirectScalarCurrentValue,
              tryApplyDirectScalarDeltas,
              tryApplyDirectFormulaDeltas,
              countPostRecalcDirectFormulaMetric,
              evaluateDirectFormula: args.evaluateDirectFormula,
            },
            hasTrackedEventListeners,
          )
        if (directChanged === undefined) {
          throw new Error('Failed to apply single direct literal mutation fast path')
        }
        recalculated = directChanged
      } else if (hasExactLookupDependents || hasSortedLookupDependents) {
        addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
      }
    })

    deferSingleCellKernelSync(existingIndex)
    const previousMetrics = args.state.getLastMetrics()
    const lastMetrics = {
      ...previousMetrics,
      dirtyFormulaCount: 0,
      wasmFormulaCount: postRecalcDirectFormulaMetrics.wasmFormulaCount,
      jsFormulaCount: postRecalcDirectFormulaMetrics.jsFormulaCount,
      rangeNodeVisits: 0,
      recalcMs: 0,
      batchId: previousMetrics.batchId + 1,
      changedInputCount: 1,
      compileMs: 0,
    }
    args.state.setLastMetrics(lastMetrics)
    if (hasTrackedEventListeners) {
      const changed = composeSingleDisjointExplicitEventChanges(existingIndex, recalculated)
      if (changed.length > 4 && canTrustPhysicalTrackedChangeSplit(changed, ref.sheetId, explicitChangedCount, args.state.workbook)) {
        tagTrustedPhysicalTrackedChanges(changed, ref.sheetId, explicitChangedCount)
      }
      args.state.events.emitTracked({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: changed,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: lastMetrics,
        explicitChangedCount,
      })
    }
    return true
  }

  const applyExistingNumericCellMutationAtNow = (
    request: EngineExistingNumericCellMutationRef,
  ): EngineExistingNumericCellMutationResult | null => {
    if (
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners() ||
      args.hasVolatileFormulas?.()
    ) {
      return null
    }
    const sheet = args.state.workbook.getSheetById(request.sheetId)
    const cellStore = args.state.workbook.cellStore
    const existingIndex = request.cellIndex
    const trustedExistingNumericLiteral = request.trustedExistingNumericLiteral === true
    if (sheet && isWorkbookTableHeaderCell(args.state.workbook, sheet.name, request.row, request.col)) {
      return null
    }
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      (!trustedExistingNumericLiteral &&
        (cellStore.sheetIds[existingIndex] !== request.sheetId ||
          cellStore.rows[existingIndex] !== request.row ||
          cellStore.cols[existingIndex] !== request.col ||
          !canFastPathLiteralOverwrite(existingIndex)))
    ) {
      return null
    }
    const oldNumber = trustedExistingNumericLiteral
      ? request.oldNumericValue === undefined || Object.is(request.oldNumericValue, -0)
        ? 0
        : request.oldNumericValue
      : directScalarCellNumericValue(existingIndex)
    if (oldNumber === undefined || Object.is(request.value, -0)) {
      return null
    }
    const sheetName = sheet.name
    const singleExistingCellDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
    const hasTrackedColumnDependents = args.hasTrackedColumnDependentsAnywhere()
    const hasExactLookupDependents = hasTrackedColumnDependents && hasTrackedExactLookupDependents(request.sheetId, request.col)
    const hasSortedLookupDependents = hasTrackedColumnDependents && hasTrackedSortedLookupDependents(request.sheetId, request.col)
    const hasTrackedEventListeners = request.emitTracked !== false && args.state.events.hasTrackedListeners()
    const hasAggregateDependents =
      isRangeEntity(singleExistingCellDependent) ||
      (hasTrackedColumnDependents && hasTrackedDirectRangeDependents(request.sheetId, request.col))
    if (trustedExistingNumericLiteral && request.emitTracked === false && isRangeEntity(singleExistingCellDependent)) {
      const trustedAggregateResult = tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation({
        existingIndex,
        rangeEntityDependent: singleExistingCellDependent,
        sheet,
        sheetId: request.sheetId,
        col: request.col,
        value: request.value,
        delta: request.value - oldNumber,
        hasExactLookupDependents,
        hasSortedLookupDependents,
      })
      if (trustedAggregateResult) {
        return trustedAggregateResult
      }
    }
    if (trustedExistingNumericLiteral && request.emitTracked === false && hasAggregateDependents && singleExistingCellDependent === -1) {
      const trustedAggregateResult = tryApplyTrustedColumnDirectAggregateExistingNumericMutation({
        existingIndex,
        sheet,
        sheetId: request.sheetId,
        sheetName,
        row: request.row,
        col: request.col,
        value: request.value,
        delta: request.value - oldNumber,
        hasExactLookupDependents,
        hasSortedLookupDependents,
      })
      if (trustedAggregateResult) {
        return trustedAggregateResult
      }
    }
    const aggregateFastPathResult =
      hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      (singleExistingCellDependent === -1 || isRangeEntity(singleExistingCellDependent))
        ? tryApplySingleDirectAggregateLiteralMutationFastPath({
            existingIndex,
            sheetId: request.sheetId,
            sheetName,
            row: request.row,
            col: request.col,
            value: request.value,
            delta: request.value - oldNumber,
            emitTracked: hasTrackedEventListeners,
            ...(isRangeEntity(singleExistingCellDependent) ? { singleRangeEntityDependent: singleExistingCellDependent } : {}),
          })
        : null
    if (aggregateFastPathResult) {
      return aggregateFastPathResult
    }
    let triedDirectLookupFastPath = false
    if (
      trustedExistingNumericLiteral &&
      !hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      singleExistingCellDependent >= 0
    ) {
      if (args.state.formulas.get(singleExistingCellDependent)?.directLookup !== undefined) {
        triedDirectLookupFastPath = true
        const directLookupFastPathResult = tryApplySingleDirectLookupOperandMutationFastPath({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          value: request.value,
          exactLookupValue: request.value,
          approximateLookupValue: request.value,
          emitTracked: hasTrackedEventListeners,
          lookupSheetHint: sheet,
          trustedInputSheet: sheet,
          trustedInputCol: request.col,
        })
        if (directLookupFastPathResult) {
          return directLookupFastPathResult
        }
      }
    }
    if (!hasAggregateDependents && (hasExactLookupDependents || hasSortedLookupDependents) && singleExistingCellDependent === -1) {
      const lookupWritePlans = planOperationLookupNumericWrites({
        isRestore: false,
        hasAggregateDependents,
        hasExactLookupDependents,
        hasSortedLookupDependents,
        sheetId: request.sheetId,
        sheetName,
        row: request.row,
        col: request.col,
        oldExactLookupNumber: hasExactLookupDependents ? oldNumber : undefined,
        newExactLookupNumber: hasExactLookupDependents ? request.value : undefined,
        oldApproximateLookupNumber: hasSortedLookupDependents ? oldNumber : undefined,
        newApproximateLookupNumber: hasSortedLookupDependents ? request.value : undefined,
        planner: {
          planExactLookupNumericColumnWrite,
          planApproximateLookupNumericColumnWrite,
        },
      })
      if (
        (!hasExactLookupDependents || lookupWritePlans.exactHandled) &&
        (!hasSortedLookupDependents || lookupWritePlans.sortedHandled) &&
        (lookupWritePlans.exact === undefined ||
          lookupWritePlans.exact.tailPatchTarget === undefined ||
          lookupWritePlans.exact.tailPatchTarget.tailPatch === undefined) &&
        (lookupWritePlans.sorted === undefined ||
          lookupWritePlans.sorted.tailPatchTarget === undefined ||
          lookupWritePlans.sorted.tailPatchTarget.tailPatch === undefined)
      ) {
        writeNumericLiteralToExistingCell(existingIndex, request.value)
        applyOperationLookupNumericWriteTailPatches(lookupWritePlans, {
          row: request.row,
          oldExactLookupNumber: hasExactLookupDependents ? oldNumber : undefined,
          newExactLookupNumber: hasExactLookupDependents ? request.value : undefined,
          oldApproximateLookupNumber: hasSortedLookupDependents ? oldNumber : undefined,
          newApproximateLookupNumber: hasSortedLookupDependents ? request.value : undefined,
          columnVersionAfterWrite: sheet.columnVersions[request.col] ?? 0,
        })
        const { changedCellIndices } = recordKernelSyncOnlyLiteralChange({
          state: args.state,
          cellIndex: existingIndex,
          deferSingleCellKernelSync,
          makeSingleLiteralSkipMetrics,
          emitTracked: hasTrackedEventListeners,
        })
        return makeExistingNumericMutationResult(changedCellIndices, 1)
      }
    }
    const singleRegionFormulaLeafDependent =
      !hasExactLookupDependents && !hasSortedLookupDependents
        ? collectSingleFormulaLeafRangeDependent(
            existingIndex,
            singleExistingCellDependent,
            request.sheetId,
            sheetName,
            request.row,
            request.col,
          )
        : -1
    if (singleRegionFormulaLeafDependent >= 0) {
      const singleRegionFormula = args.state.formulas.get(singleRegionFormulaLeafDependent)
      if (
        singleRegionFormula !== undefined &&
        singleRegionFormula.directLookup === undefined &&
        singleRegionFormula.directAggregate === undefined &&
        singleRegionFormula.directCriteria === undefined &&
        singleRegionFormula.directScalar === undefined
      ) {
        const formulaLeafResult = tryApplyTrustedFormulaLeafExistingNumericMutation({
          existingIndex,
          formulaCellIndex: singleRegionFormulaLeafDependent,
          sheet,
          col: request.col,
          value: request.value,
          oldNumber,
          hasTrackedEventListeners,
        })
        if (formulaLeafResult) {
          return formulaLeafResult
        }
      }
    }
    if (hasKnownDynamicFormulaDependents(request.sheetId, request.row, request.col, existingIndex, singleExistingCellDependent)) {
      return null
    }
    if (
      trustedExistingNumericLiteral &&
      request.emitTracked === false &&
      !hasAggregateDependents &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      singleExistingCellDependent < -1
    ) {
      const recalculated = tryApplySingleDirectScalarLiteralMutationWithoutEventsAndReturnChanged({
        existingIndex,
        value: request.value,
        oldNumber,
        newNumber: request.value,
      })
      if (recalculated !== null) {
        const changed = composeSingleDisjointExplicitEventChanges(existingIndex, recalculated)
        if (changed.length > 4 && canTrustPhysicalTrackedChangeSplit(changed, request.sheetId, 1, args.state.workbook)) {
          tagTrustedPhysicalTrackedChanges(changed, request.sheetId, 1)
        }
        return makeExistingNumericMutationResult(changed, 1)
      }
    }
    if (trustedExistingNumericLiteral && !hasExactLookupDependents && !hasSortedLookupDependents && singleExistingCellDependent >= 0) {
      const singleDependentFormula = args.state.formulas.get(singleExistingCellDependent)
      if (!hasAggregateDependents && singleDependentFormula?.directScalar !== undefined) {
        const scalarClosureResult = tryApplyTrustedDirectScalarClosureExistingNumericMutation({
          existingIndex,
          sheet,
          sheetId: request.sheetId,
          col: request.col,
          value: request.value,
          oldNumber,
          hasTrackedEventListeners,
        })
        if (scalarClosureResult) {
          return scalarClosureResult
        }
      }
      if (
        singleDependentFormula !== undefined &&
        singleDependentFormula.directLookup === undefined &&
        singleDependentFormula.directAggregate === undefined &&
        singleDependentFormula.directCriteria === undefined &&
        singleDependentFormula.directScalar === undefined
      ) {
        const formulaLeafResult = tryApplyTrustedFormulaLeafExistingNumericMutation({
          existingIndex,
          formulaCellIndex: singleExistingCellDependent,
          sheet,
          col: request.col,
          value: request.value,
          oldNumber,
          hasTrackedEventListeners,
        })
        if (formulaLeafResult) {
          return formulaLeafResult
        }
      }
    }
    const directLookupFastPathResult =
      !triedDirectLookupFastPath && !hasAggregateDependents && !hasExactLookupDependents && !hasSortedLookupDependents
        ? tryApplySingleDirectLookupOperandMutationFastPath({
            existingIndex,
            formulaCellIndex: singleExistingCellDependent,
            value: request.value,
            exactLookupValue: request.value,
            approximateLookupValue: request.value,
            emitTracked: hasTrackedEventListeners,
            lookupSheetHint: sheet,
            ...(trustedExistingNumericLiteral ? { trustedInputSheet: sheet, trustedInputCol: request.col } : {}),
          })
        : null
    if (directLookupFastPathResult) {
      return directLookupFastPathResult
    }
    return null
  }

  const applyExistingLiteralCellMutationAtNow = (
    request: EngineExistingLiteralCellMutationRef,
  ): EngineExistingNumericCellMutationResult | null => {
    if (typeof request.value === 'number') {
      return applyExistingNumericCellMutationAtNow({
        sheetId: request.sheetId,
        row: request.row,
        col: request.col,
        cellIndex: request.cellIndex,
        value: request.value,
        ...(request.emitTracked === undefined ? {} : { emitTracked: request.emitTracked }),
      })
    }
    if (
      args.state.workbook.hasPivots() ||
      args.state.events.hasListeners() ||
      args.state.events.hasCellListeners() ||
      args.hasVolatileFormulas?.()
    ) {
      return null
    }
    const sheet = args.state.workbook.getSheetById(request.sheetId)
    const cellStore = args.state.workbook.cellStore
    const existingIndex = request.cellIndex
    if (sheet && isWorkbookTableHeaderCell(args.state.workbook, sheet.name, request.row, request.col)) {
      return null
    }
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      cellStore.sheetIds[existingIndex] !== request.sheetId ||
      cellStore.rows[existingIndex] !== request.row ||
      cellStore.cols[existingIndex] !== request.col ||
      !canFastPathLiteralOverwrite(existingIndex)
    ) {
      return null
    }
    const formulaCellIndex = args.getSingleEntityDependent(makeCellEntity(existingIndex))
    if (hasKnownDynamicFormulaDependents(request.sheetId, request.row, request.col, existingIndex, formulaCellIndex)) {
      return null
    }
    if (
      args.hasTrackedColumnDependentsAnywhere() &&
      (hasTrackedDirectRangeDependents(request.sheetId, request.col) ||
        hasTrackedExactLookupDependents(request.sheetId, request.col) ||
        hasTrackedSortedLookupDependents(request.sheetId, request.col))
    ) {
      return null
    }
    const hasTrackedEventListeners = request.emitTracked !== false && args.state.events.hasTrackedListeners()
    const directLookupResult = tryApplySingleDirectLookupOperandMutationFastPath({
      existingIndex,
      formulaCellIndex,
      value: request.value,
      exactLookupValue: undefined,
      approximateLookupValue: undefined,
      emitTracked: hasTrackedEventListeners,
      lookupSheetHint: sheet,
    })
    if (directLookupResult) {
      return directLookupResult
    }
    return tryApplyFormulaLeafExistingLiteralMutation({
      existingIndex,
      formulaCellIndex,
      value: request.value,
      hasTrackedEventListeners,
    })
  }

  return { tryApplySingleExistingDirectLiteralMutation, applyExistingNumericCellMutationAtNow, applyExistingLiteralCellMutationAtNow }
}
