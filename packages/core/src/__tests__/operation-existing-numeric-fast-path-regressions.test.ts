import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../cell-mutations-at.js'
import { SpreadsheetEngine } from '../engine.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectCriteriaDescriptor } from '../engine/runtime-state.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import { aggregateColumnDependencyKey } from '../engine/services/direct-formula-recalc-helpers.js'
import type { OperationDirectAggregateLiteralMutationRequest } from '../engine/services/operation-direct-aggregate-literal-fast-path.js'
import { createOperationDirectRangeDependentService } from '../engine/services/operation-direct-range-dependents.js'

function operationHook<Args extends readonly unknown[], Result>(engine: SpreadsheetEngine, name: string): (...args: Args) => Result {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (typeof operations !== 'object' || operations === null) {
    throw new TypeError('Expected operation service')
  }
  const hooks = Reflect.get(operations, '__testHooks')
  if (typeof hooks !== 'object' || hooks === null) {
    throw new TypeError('Expected operation hooks')
  }
  const hook = Reflect.get(hooks, name)
  if (typeof hook !== 'function') {
    throw new TypeError(`Expected operation hook ${name}`)
  }
  return (...args: Args): Result => Reflect.apply(hook, hooks, args)
}

function directAggregateDescriptor(rowStart: number, rowEnd: number): RuntimeDirectAggregateDescriptor {
  return {
    regionId: rowStart + 1,
    aggregateKind: 'sum',
    sheetName: 'Sheet1',
    rowStart,
    rowEnd,
    col: 0,
    colEnd: 0,
    length: rowEnd - rowStart + 1,
  }
}

function directCriteriaDescriptor(rowStart: number, rowEnd: number): RuntimeDirectCriteriaDescriptor {
  return {
    aggregateKind: 'sum',
    aggregateRange: {
      regionId: rowStart + 20,
      sheetName: 'Sheet1',
      rowStart,
      rowEnd,
      col: 0,
      length: rowEnd - rowStart + 1,
    },
    criteriaPairs: [],
  }
}

describe('existing numeric mutation fast path regressions', () => {
  it('collects multiple aggregate-column dependents into post-recalc deltas', () => {
    const formulas = new Map([
      [
        10,
        {
          directAggregate: directAggregateDescriptor(0, 4),
          directCriteria: undefined,
          dependencyIndices: { length: 0 },
        },
      ],
      [
        11,
        {
          directAggregate: directAggregateDescriptor(1, 5),
          directCriteria: undefined,
          dependencyIndices: { length: 0 },
        },
      ],
    ])
    const marked: number[] = []
    const service = createOperationDirectRangeDependentService({
      workbook: {
        getSheet: (sheetName) => (sheetName === 'Sheet1' ? { id: 1 } : undefined),
      },
      formulas,
      reverseAggregateColumnEdges: new Map([[aggregateColumnDependencyKey(1, 0), new Set([10, 11])]]),
      collectRegionFormulaDependentsForCell: () => new Uint32Array(),
      collectSingleRegionFormulaDependentForCell: () => -1,
      hasNoCellDependents: () => true,
      getSingleEntityDependent: () => -1,
      markFormulaChanged: (cellIndex, count) => {
        marked.push(cellIndex)
        return count + 1
      },
      tryDirectCriteriaSumDelta: () => undefined,
      postRecalcLimit: 8,
    })
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()

    expect(
      service.markAffectedDirectRangeDependents(
        {
          sheetName: 'Sheet1',
          sheetId: 1,
          row: 2,
          col: 0,
          oldValue: { tag: ValueTag.Number, value: 2 },
          newValue: { tag: ValueTag.Number, value: 5 },
          inputCellIndex: 99,
        },
        7,
        postRecalcDirectFormulaIndices,
      ),
    ).toBe(7)

    expect(service.collectAffectedDirectRangeDependents({ sheetName: 'Sheet1', sheetId: 1, row: 2, col: 0 })).toEqual([10, 11])
    expect(service.collectSingleAffectedDirectRangeDependent({ sheetName: 'Sheet1', sheetId: 1, row: 2, col: 0 })).toBe(-2)
    expect(postRecalcDirectFormulaIndices.size).toBe(2)
    expect(postRecalcDirectFormulaIndices.getDelta(10)).toBe(3)
    expect(postRecalcDirectFormulaIndices.getDelta(11)).toBe(3)
    expect(postRecalcDirectFormulaIndices.hasCoveredDirectRangeInput(99)).toBe(true)
    expect(marked).toEqual([])
  })

  it('rejects single aggregate fast-path selection when reverse column dependents conflict', () => {
    const formulas = new Map([
      [
        20,
        {
          directAggregate: directAggregateDescriptor(0, 4),
          directCriteria: undefined,
          dependencyIndices: { length: 0 },
        },
      ],
      [
        21,
        {
          directAggregate: directAggregateDescriptor(1, 5),
          directCriteria: undefined,
          dependencyIndices: { length: 0 },
        },
      ],
    ])
    const service = createOperationDirectRangeDependentService({
      workbook: {
        getSheet: (sheetName) => (sheetName === 'Sheet1' ? { id: 1 } : undefined),
      },
      formulas,
      reverseAggregateColumnEdges: new Map([[aggregateColumnDependencyKey(1, 0), new Set([20, 21])]]),
      collectRegionFormulaDependentsForCell: () => new Uint32Array(),
      collectSingleRegionFormulaDependentForCell: () => -1,
      hasNoCellDependents: () => true,
      getSingleEntityDependent: () => -1,
      markFormulaChanged: (_cellIndex, count) => count + 1,
      tryDirectCriteriaSumDelta: () => undefined,
      postRecalcLimit: 8,
    })

    expect(service.collectSingleApplicableDirectAggregateDependent({ sheetName: 'Sheet1', sheetId: 1, row: 2, col: 0 })).toBe(-2)
  })

  it('collects direct criteria dependents into shared post-recalc deltas', () => {
    const formulas = new Map([
      [
        30,
        {
          directAggregate: undefined,
          directCriteria: directCriteriaDescriptor(0, 4),
          dependencyIndices: { length: 0 },
        },
      ],
      [
        31,
        {
          directAggregate: undefined,
          directCriteria: directCriteriaDescriptor(1, 5),
          dependencyIndices: { length: 0 },
        },
      ],
    ])
    const marked: number[] = []
    const service = createOperationDirectRangeDependentService({
      workbook: {
        getSheet: (sheetName) => (sheetName === 'Sheet1' ? { id: 1 } : undefined),
      },
      formulas,
      reverseAggregateColumnEdges: new Map([[aggregateColumnDependencyKey(1, 0), new Set([30, 31])]]),
      collectRegionFormulaDependentsForCell: () => new Uint32Array(),
      collectSingleRegionFormulaDependentForCell: () => -1,
      hasNoCellDependents: () => true,
      getSingleEntityDependent: () => -1,
      markFormulaChanged: (cellIndex, count) => {
        marked.push(cellIndex)
        return count + 1
      },
      tryDirectCriteriaSumDelta: () => 4,
      postRecalcLimit: 8,
    })
    const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()

    expect(
      service.markAffectedDirectRangeDependents(
        {
          sheetName: 'Sheet1',
          sheetId: 1,
          row: 2,
          col: 0,
          oldValue: { tag: ValueTag.Number, value: 2 },
          newValue: { tag: ValueTag.Number, value: 9 },
          inputCellIndex: 100,
        },
        5,
        postRecalcDirectFormulaIndices,
      ),
    ).toBe(5)

    expect(postRecalcDirectFormulaIndices.size).toBe(2)
    expect(postRecalcDirectFormulaIndices.getDelta(30)).toBe(4)
    expect(postRecalcDirectFormulaIndices.getDelta(31)).toBe(4)
    expect(postRecalcDirectFormulaIndices.hasCoveredDirectRangeInput(100)).toBe(true)
    expect(marked).toEqual([])
  })

  it('deduplicates indexed direct formula collections', () => {
    const collection = new DirectFormulaIndexCollection()

    for (let offset = 0; offset < 17; offset += 1) {
      collection.add(100 + offset)
    }
    collection.add(100)
    collection.add(150)

    expect(collection.size).toBe(18)
    expect(collection.has(150)).toBe(true)
    expect(collection.has(999)).toBe(false)
    expect(collection.hasAny([], 0)).toBe(false)
    expect(collection.hasAny([999, 150], 2)).toBe(true)

    const seen: number[] = []
    collection.forEach((cellIndex) => seen.push(cellIndex))
    expect(seen.slice(0, 3)).toEqual([100, 101, 102])
    expect(seen.at(-1)).toBe(150)

    const indexed: Array<[number, number]> = []
    collection.forEachIndexed((cellIndex, index) => indexed.push([cellIndex, index]))
    expect(indexed.at(-1)).toEqual([150, 17])
    expect(collection.getCellIndexAt(17)).toBe(150)
    expect(collection.getCellIndicesForRead()).toHaveLength(18)
  })

  it('uses set-backed coverage checks for large direct formula and range inputs', () => {
    const collection = new DirectFormulaIndexCollection()

    for (let offset = 0; offset < 17; offset += 1) {
      collection.markDirectFormulaInputCovered(200 + offset)
      collection.markDirectRangeInputCovered(300 + offset)
    }
    collection.markDirectFormulaInputCovered(200)
    collection.markDirectFormulaInputCovered(260)
    collection.markDirectRangeInputCovered(300)
    collection.markDirectRangeInputCovered(360)

    expect(collection.hasCoveredDirectFormulaInput(260)).toBe(true)
    expect(collection.hasCoveredDirectFormulaInput(999)).toBe(false)
    expect(collection.hasCoveredDirectRangeInput(360)).toBe(true)
    expect(collection.hasCoveredDirectRangeInput(999)).toBe(false)
  })

  it('merges prepared scalar deltas without preserving stale scalar completeness', () => {
    const collection = new DirectFormulaIndexCollection()
    const baseCells = Array.from({ length: 17 }, (_, index) => 400 + index)

    collection.appendConstantDelta(baseCells, 2, 'scalar')
    expect(collection.getConstantScalarDelta()).toBe(2)
    expect(collection.hasCompleteScalarDeltas()).toBe(true)

    collection.addDelta(400, 3)
    expect(collection.getDelta(400)).toBe(5)
    expect(collection.getScalarDeltaAt(0)).toBeUndefined()
    expect(collection.hasCompleteScalarDeltas()).toBe(false)

    collection.appendDeltas([500, 501], [7, 8], 'scalar')
    expect(collection.size).toBe(19)
    expect(collection.getDelta(500)).toBe(7)
    expect(collection.getScalarDeltaAt(18)).toBe(8)
    expect(collection.hasCompleteDeltas()).toBe(true)
  })

  it('assigns deltas when direct formula collections cross the indexed threshold', () => {
    const thresholdCollection = new DirectFormulaIndexCollection()
    for (let offset = 0; offset < 16; offset += 1) {
      thresholdCollection.add(600 + offset)
    }

    thresholdCollection.addScalarDelta(700, 9)
    expect(thresholdCollection.size).toBe(17)
    expect(thresholdCollection.getDelta(700)).toBe(9)
    expect(thresholdCollection.getScalarDeltaAt(16)).toBe(9)
    expect(thresholdCollection.has(700)).toBe(true)

    const indexedCollection = new DirectFormulaIndexCollection()
    for (let offset = 0; offset < 17; offset += 1) {
      indexedCollection.add(800 + offset)
    }

    indexedCollection.addScalarDelta(900, 11)
    expect(indexedCollection.size).toBe(18)
    expect(indexedCollection.getDelta(900)).toBe(11)
    expect(indexedCollection.getScalarDeltaAt(17)).toBe(11)
    expect(indexedCollection.has(900)).toBe(true)
  })

  it('applies prepared bulk deltas to existing indexed formulas without duplicates', () => {
    const collection = new DirectFormulaIndexCollection()
    for (let offset = 0; offset < 17; offset += 1) {
      collection.add(1000 + offset)
    }

    collection.appendDeltas([1000, 1100], [4, 5], 'scalar')
    expect(collection.size).toBe(18)
    expect(collection.getDelta(1000)).toBe(4)
    expect(collection.getDelta(1100)).toBe(5)
    expect(collection.getScalarDeltaAt(0)).toBe(4)
    expect(collection.getScalarDeltaAt(17)).toBe(5)

    collection.appendConstantDelta([1000, 1200], 2)
    expect(collection.size).toBe(19)
    expect(collection.getDelta(1000)).toBe(6)
    expect(collection.getDelta(1200)).toBe(2)
    expect(collection.hasCompleteScalarDeltas()).toBe(false)
  })

  it('rejects collected direct criteria dependents before aggregate delta application', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-direct-criteria-not-aggregate-delta',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setCellValue('Data', 'A1', 'Deposit')
    engine.setCellValue('Data', 'B1', 1.235)
    engine.setCellValue('Data', 'A2', 'Deposit')
    engine.setCellValue('Data', 'B2', 0)
    engine.setCellFormula('Summary', 'A1', 'ROUND(SUMIFS(Data!$B$1:$B$2,Data!$A$1:$A$2,"Deposit"),2)')

    const sheet = engine.workbook.getSheet('Data')!
    const inputIndex = engine.workbook.getCellIndex('Data', 'B1')!
    const fastPath = operationHook<[OperationDirectAggregateLiteralMutationRequest], EngineExistingNumericCellMutationResult | null>(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )
    const result = fastPath({
      existingIndex: inputIndex,
      sheetId: sheet.id,
      sheetName: 'Data',
      row: 0,
      col: 1,
      value: 1.236,
      delta: 0.001,
      emitTracked: false,
    })

    expect(result).toBeNull()
    expect(engine.getCellValue('Data', 'B1')).toEqual({ tag: ValueTag.Number, value: 1.235 })
    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 1.24 })
  })

  it('applies terminal aggregate deltas for valid off-column collected dependents', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-terminal-aggregate-delta',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A2)')

    const sheet = engine.workbook.getSheet('Sheet1')!
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const fastPath = operationHook<[OperationDirectAggregateLiteralMutationRequest], EngineExistingNumericCellMutationResult | null>(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )
    const result = fastPath({
      existingIndex: inputIndex,
      sheetId: sheet.id,
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      value: 5,
      delta: 4,
      emitTracked: false,
    })

    expect(result).toMatchObject({
      firstChangedCellIndex: inputIndex,
      changedCellCount: 2,
      explicitChangedCount: 1,
      secondChangedNumericValue: 7,
      secondChangedRow: 0,
      secondChangedCol: 2,
    })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })
  })

  it('recalculates every aggregate touching a range-entity input after text replacement', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-overlapping-aggregate-text-replacement',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      for (let col = 0; col < 5; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row + 1}`, row * 10 + col + 1)
      }
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')
    engine.insertColumns('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:B2)')

    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 13 })

    engine.setCellValue('Sheet1', 'B2', 'north')

    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 94 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('falls back when a numeric input has both aggregate and scalar formula dependents', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-mixed-dependent-fallback',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, row * 10 + 1)
    }
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!

    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 0,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
    })

    expect(result).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })

    engine.setCellValue('Sheet1', 'A1', 0)

    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 104 })
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 2 })
  })
})
