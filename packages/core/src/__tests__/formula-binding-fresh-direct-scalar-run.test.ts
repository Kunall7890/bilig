import { describe, expect, it } from 'vitest'
import { compileFormula } from '@bilig/formula'
import { SpreadsheetEngine } from '../engine.js'
import { makeCellEntity } from '../entity-ids.js'
import type { EngineFormulaBindingService } from '../engine/services/formula-binding-service.js'
import { getFormulaBindingReverseEdgeSlice, type FormulaBindingReverseEdgeState } from '../engine/services/formula-binding-reverse-edges.js'

interface TestEdgeArena {
  read(slice: { ptr: number; len: number; cap: number }): Uint32Array
}

interface TestEdgeSlice {
  ptr: number
  len: number
  cap: number
}

function isEngineFormulaBindingService(value: unknown): value is EngineFormulaBindingService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'bindFreshDirectScalarFormulaRunNow') === 'function'
}

function isTestEdgeArena(value: unknown): value is TestEdgeArena {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'read') === 'function'
}

function getBindingService(engine: SpreadsheetEngine): EngineFormulaBindingService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const binding = Reflect.get(runtime, 'binding')
  if (!isEngineFormulaBindingService(binding)) {
    throw new TypeError('Expected engine formula binding service')
  }
  return binding
}

function readRuntimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const formulas = Reflect.get(engine, 'formulas')
  if (typeof formulas !== 'object' || formulas === null || typeof Reflect.get(formulas, 'get') !== 'function') {
    throw new TypeError('Expected internal formulas store')
  }
  return Reflect.get(formulas, 'get').call(formulas, cellIndex)
}

function readRuntimeFormulaProperty(formula: unknown, property: string): unknown {
  if (typeof formula !== 'object' || formula === null) {
    throw new TypeError('Expected runtime formula')
  }
  return Reflect.get(formula, property)
}

function readRuntimeDirectScalar(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const formula = readRuntimeFormula(engine, cellIndex)
  return typeof formula === 'object' && formula !== null ? Reflect.get(formula, 'directScalar') : undefined
}

function readRuntimeDependencyEntitiesSlice(engine: SpreadsheetEngine, cellIndex: number): TestEdgeSlice {
  const slice = readRuntimeFormulaProperty(readRuntimeFormula(engine, cellIndex), 'dependencyEntities')
  const ptr = typeof slice === 'object' && slice !== null ? Reflect.get(slice, 'ptr') : undefined
  const len = typeof slice === 'object' && slice !== null ? Reflect.get(slice, 'len') : undefined
  const cap = typeof slice === 'object' && slice !== null ? Reflect.get(slice, 'cap') : undefined
  if (typeof slice !== 'object' || slice === null || typeof ptr !== 'number' || typeof len !== 'number' || typeof cap !== 'number') {
    throw new TypeError('Expected dependency edge slice')
  }
  return {
    ptr,
    len,
    cap,
  }
}

function readReverseDependents(engine: SpreadsheetEngine, dependencyCellIndex: number): number[] {
  const reverseState = {
    reverseCellEdges: Reflect.get(engine, 'reverseCellEdges'),
    reverseRangeEdges: Reflect.get(engine, 'reverseRangeEdges'),
    reverseDefinedNameEdges: Reflect.get(engine, 'reverseDefinedNameEdges'),
    reverseTableEdges: Reflect.get(engine, 'reverseTableEdges'),
    reverseSpillEdges: Reflect.get(engine, 'reverseSpillEdges'),
    reverseAggregateColumnEdges: Reflect.get(engine, 'reverseAggregateColumnEdges'),
    reverseExactLookupColumnEdges: Reflect.get(engine, 'reverseExactLookupColumnEdges'),
    reverseSortedLookupColumnEdges: Reflect.get(engine, 'reverseSortedLookupColumnEdges'),
  } as FormulaBindingReverseEdgeState
  const edgeArena = Reflect.get(engine, 'edgeArena')
  if (!isTestEdgeArena(edgeArena)) {
    throw new TypeError('Expected engine edge arena')
  }
  const slice = getFormulaBindingReverseEdgeSlice(reverseState, makeCellEntity(dependencyCellIndex))
  return slice ? [...edgeArena.read(slice)] : []
}

describe('fresh direct scalar formula run binding', () => {
  it('bulk-binds runs with packed dependencies and reverse slices', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-fresh-direct-scalar-run-bulk' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    expect(sheetId).toBeDefined()

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellValue('Sheet1', 'C1', 0)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'B2', 4)
    engine.setCellValue('Sheet1', 'C2', 0)

    const c1Index = engine.workbook.getCellIndex('Sheet1', 'C1')!
    const c2Index = engine.workbook.getCellIndex('Sheet1', 'C2')!
    getBindingService(engine).bindFreshDirectScalarFormulaRunNow({
      sheetId: sheetId!,
      ownerSheetName: 'Sheet1',
      cellIndices: new Uint32Array([c1Index, c2Index]),
      members: [
        { row: 0, col: 2, source: 'A1+B1', compiled: compileFormula('A1+B1'), templateId: 1 },
        { row: 1, col: 2, source: 'A2+B2', compiled: compileFormula('A2+B2'), templateId: 1 },
      ],
    })

    expect(readRuntimeDirectScalar(engine, c1Index)).toBeDefined()
    expect(readRuntimeDirectScalar(engine, c2Index)).toBeDefined()
    const c1Dependencies = readRuntimeDependencyEntitiesSlice(engine, c1Index)
    const c2Dependencies = readRuntimeDependencyEntitiesSlice(engine, c2Index)
    expect(c1Dependencies).toMatchObject({ len: 2, cap: 2 })
    expect(c2Dependencies).toMatchObject({ ptr: c1Dependencies.ptr + 2, len: 2, cap: 2 })
    expect(readReverseDependents(engine, engine.workbook.getCellIndex('Sheet1', 'A1')!)).toEqual([c1Index])
    expect(readReverseDependents(engine, engine.workbook.getCellIndex('Sheet1', 'B2')!)).toEqual([c2Index])
    expect(engine.getPerformanceCounters()).toMatchObject({
      freshDirectScalarBulkRunBindings: 1,
      freshDirectScalarBulkMembers: 2,
      freshDirectScalarBulkFallbacks: 0,
      freshDirectScalarBulkReverseEdgeSlices: 4,
      freshDirectScalarFormulaObjectsMaterialized: 2,
    })
  })

  it('keeps duplicate references unique when a run cannot pack reverse edges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-fresh-direct-scalar-run-duplicate-refs' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    expect(sheetId).toBeDefined()

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 0)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'B2', 0)

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')!
    const b2Index = engine.workbook.getCellIndex('Sheet1', 'B2')!
    getBindingService(engine).bindFreshDirectScalarFormulaRunNow({
      sheetId: sheetId!,
      ownerSheetName: 'Sheet1',
      cellIndices: new Uint32Array([b1Index, b2Index]),
      members: [
        { row: 0, col: 1, source: 'A1+A1', compiled: compileFormula('A1+A1'), templateId: 1 },
        { row: 1, col: 1, source: 'A2+A2', compiled: compileFormula('A2+A2'), templateId: 1 },
      ],
    })

    expect(readReverseDependents(engine, engine.workbook.getCellIndex('Sheet1', 'A1')!)).toEqual([b1Index])
    expect(readReverseDependents(engine, engine.workbook.getCellIndex('Sheet1', 'A2')!)).toEqual([b2Index])
    expect(engine.getPerformanceCounters()).toMatchObject({
      freshDirectScalarBulkRunBindings: 1,
      freshDirectScalarBulkMembers: 2,
      freshDirectScalarBulkReverseEdgeSlices: 0,
      freshDirectScalarFormulaObjectsMaterialized: 2,
    })
  })
})
