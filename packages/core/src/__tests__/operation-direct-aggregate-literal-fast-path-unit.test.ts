import { describe, expect, it, vi } from 'vitest'
import { makeRangeEntity } from '../entity-ids.js'
import { createInitialRecalcMetrics } from '../engine/runtime-state.js'
import {
  tryApplySingleDirectAggregateLiteralMutationFastPath,
  tryApplyTrustedColumnDirectAggregateExistingNumericMutation,
  tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation,
  type OperationDirectAggregateLiteralFastPathArgs,
} from '../engine/services/operation-direct-aggregate-literal-fast-path.js'
import { createWorkbookSheetRecord } from '../workbook-sheet-record.js'

function testSheet() {
  return createWorkbookSheetRecord({ id: 1, name: 'Sheet1', order: 0, counters: undefined })
}

function directAggregateArgs(
  overrides: Partial<OperationDirectAggregateLiteralFastPathArgs> = {},
): OperationDirectAggregateLiteralFastPathArgs {
  const cellStore = {
    sheetIds: [] as number[],
    rows: [] as number[],
    cols: [] as number[],
  }
  cellStore.sheetIds[10] = 1
  cellStore.rows[10] = 4
  cellStore.cols[10] = 6
  cellStore.sheetIds[11] = 1
  cellStore.rows[11] = 5
  cellStore.cols[11] = 7
  cellStore.sheetIds[12] = 1
  cellStore.rows[12] = 6
  cellStore.cols[12] = 8

  return {
    state: {
      workbook: { cellStore },
      counters: {
        directAggregateDeltaApplications: 0,
        directAggregateDeltaOnlyRecalcSkips: 0,
      },
      events: {
        emitTracked: vi.fn(),
      },
      setLastMetrics: vi.fn(),
    },
    directRangePostRecalcLimit: 8,
    getSingleEntityDependent: () => -1,
    collectSingleAffectedDirectRangeDependent: () => -1,
    collectAffectedDirectRangeDependents: () => [],
    collectSingleApplicableDirectAggregateDependent: () => -1,
    canApplyDirectAggregateLiteralDeltaForRequest: () => true,
    canApplyDirectAggregateLiteralDelta: () => true,
    writeFastPathLiteralToExistingCell: vi.fn(),
    writeTrustedExistingNumericLiteralToCell: vi.fn(),
    applyTerminalDirectFormulaNumericDeltaAndReturn: vi.fn(() => 22),
    applyDirectFormulaNumericDelta: vi.fn(() => true),
    applyDirectFormulaNumericDeltaBatch: vi.fn(() => true),
    cellsShareVersionColumn: () => false,
    withOptionalColumnVersionBatch: (_enabled, apply) => apply(),
    deferSingleCellKernelSync: vi.fn(),
    makeSingleLiteralSkipMetrics: () => createInitialRecalcMetrics(),
    ...overrides,
  }
}

describe('direct aggregate literal fast path units', () => {
  it('writes a literal without aggregate deltas when no range dependent applies', () => {
    const args = directAggregateArgs()

    const result = tryApplySingleDirectAggregateLiteralMutationFastPath(args, {
      existingIndex: 1,
      sheetId: 1,
      sheetName: 'Sheet1',
      row: 20,
      col: 0,
      value: 9,
      delta: 8,
      emitTracked: false,
    })

    expect(result).toMatchObject({ firstChangedCellIndex: 1, changedCellCount: 1 })
    expect(args.writeFastPathLiteralToExistingCell).toHaveBeenCalledWith(1, 9)
    expect(args.deferSingleCellKernelSync).toHaveBeenCalledWith(1)
    expect(args.state.counters.directAggregateDeltaApplications).toBe(0)
  })

  it('emits tracked literal writes when no aggregate dependent applies', () => {
    const args = directAggregateArgs()

    const result = tryApplySingleDirectAggregateLiteralMutationFastPath(args, {
      existingIndex: 2,
      sheetId: 1,
      sheetName: 'Sheet1',
      row: 21,
      col: 0,
      value: 10,
      delta: 9,
      emitTracked: true,
    })

    expect(result).toEqual({
      changedCellIndices: new Uint32Array([2]),
      explicitChangedCount: 1,
    })
    expect(Reflect.get(args.state.events, 'emitTracked')).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([2]),
        explicitChangedCount: 1,
      }),
    )
  })

  it('applies a terminal aggregate delta without batching when the aggregate is on another column', () => {
    const args = directAggregateArgs({
      collectSingleAffectedDirectRangeDependent: () => 10,
    })

    const result = tryApplySingleDirectAggregateLiteralMutationFastPath(args, {
      existingIndex: 3,
      sheetId: 1,
      sheetName: 'Sheet1',
      row: 4,
      col: 0,
      value: 12,
      delta: 11,
      emitTracked: false,
    })

    expect(result).toMatchObject({
      firstChangedCellIndex: 3,
      secondChangedCellIndex: 10,
      secondChangedNumericValue: 22,
      secondChangedRow: 4,
      secondChangedCol: 6,
    })
    expect(args.applyTerminalDirectFormulaNumericDeltaAndReturn).toHaveBeenCalledWith(10, 11)
    expect(args.state.counters.directAggregateDeltaApplications).toBe(1)
  })

  it('batches column versions while applying the same delta to multiple aggregates', () => {
    const batched: boolean[] = []
    const args = directAggregateArgs({
      collectSingleAffectedDirectRangeDependent: () => -2,
      collectAffectedDirectRangeDependents: () => [10, 11, 12],
      withOptionalColumnVersionBatch: (enabled, apply) => {
        batched.push(enabled)
        apply()
      },
    })

    const result = tryApplySingleDirectAggregateLiteralMutationFastPath(args, {
      existingIndex: 4,
      sheetId: 1,
      sheetName: 'Sheet1',
      row: 1,
      col: 0,
      value: 15,
      delta: 14,
      emitTracked: false,
    })

    expect(result).toEqual({
      changedCellIndices: new Uint32Array([4, 10, 11, 12]),
      explicitChangedCount: 1,
    })
    expect(batched).toEqual([true])
    expect(args.applyDirectFormulaNumericDeltaBatch).toHaveBeenCalledWith([10, 11, 12], 14)
    expect(args.state.counters.directAggregateDeltaApplications).toBe(3)
  })

  it('uses the trusted range aggregate mutation path when range and indexed dependents agree', () => {
    const rangeEntity = makeRangeEntity(5)
    const sheet = testSheet()
    const args = directAggregateArgs({
      getSingleEntityDependent: (entityId) => (entityId === rangeEntity ? 10 : -1),
      collectSingleAffectedDirectRangeDependent: () => 10,
      collectAffectedDirectRangeDependents: () => [10],
      collectSingleApplicableDirectAggregateDependent: () => 10,
    })

    const result = tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation(args, {
      existingIndex: 5,
      rangeEntityDependent: rangeEntity,
      sheet,
      sheetId: 1,
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      value: 16,
      delta: 15,
      hasExactLookupDependents: false,
      hasSortedLookupDependents: false,
    })

    expect(result).toMatchObject({
      firstChangedCellIndex: 5,
      secondChangedCellIndex: 10,
      secondChangedNumericValue: 22,
    })
    expect(args.writeTrustedExistingNumericLiteralToCell).toHaveBeenCalledWith(5, sheet, 0, 16)
  })

  it('uses the trusted column aggregate mutation path for indexed aggregate columns', () => {
    const sheet = testSheet()
    const args = directAggregateArgs({
      collectSingleApplicableDirectAggregateDependent: () => 12,
    })

    const result = tryApplyTrustedColumnDirectAggregateExistingNumericMutation(args, {
      existingIndex: 6,
      sheet,
      sheetId: 1,
      sheetName: 'Sheet1',
      row: 2,
      col: 0,
      value: 17,
      delta: 16,
      hasExactLookupDependents: false,
      hasSortedLookupDependents: false,
    })

    expect(result).toMatchObject({
      firstChangedCellIndex: 6,
      secondChangedCellIndex: 12,
      secondChangedNumericValue: 22,
      secondChangedRow: 6,
      secondChangedCol: 8,
    })
    expect(args.writeTrustedExistingNumericLiteralToCell).toHaveBeenCalledWith(6, sheet, 0, 17)
  })
})
