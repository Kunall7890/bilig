import { describe, expect, it } from 'vitest'
import { ErrorCode, FormulaMode, ValueTag, type CellValue } from '@bilig/protocol'
import type { CompiledFormula, FormulaNode } from '@bilig/formula'
import type {
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectScalarDescriptor,
  RuntimeFormula,
} from '../engine/runtime-state.js'
import { readRuntimeDirectCriteriaOperandValue } from '../engine/services/direct-criteria-operands.js'
import {
  numericLikeValueInView,
  strictNumericAggregateCandidateInView,
} from '../engine/services/formula-evaluation-direct-criteria-transforms.js'
import { tryEvaluateDirectAggregate } from '../engine/services/formula-evaluation-direct-aggregate.js'
import { tryEvaluateDirectIndexExactMatch, tryEvaluateDirectIndexOffset } from '../engine/services/formula-evaluation-direct-index.js'
import { tryEvaluateDirectScalar } from '../engine/services/formula-evaluation-direct-scalar.js'
import type { RangeAggregateCacheService } from '../engine/services/range-aggregate-cache-service.js'
import { createEngineCounters } from '../perf/engine-counters.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from '../engine/services/runtime-column-store-service.js'

const emptyCell: CellValue = { tag: ValueTag.Empty }
const valueCell = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const booleanCell = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const stringCell = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const errorCell = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const numberAst: FormulaNode = { kind: 'NumberLiteral', value: 0 }

function makeColumnView(cells: readonly CellValue[]): RuntimeColumnView {
  const owner = {
    sheetName: 'Sheet1',
    col: 0,
    columnVersion: 1,
    structureVersion: 1,
    sheetColumnVersions: new Uint32Array([1]),
    pages: new Map(),
  }
  const readCellValueAt = (offset: number): CellValue => cells[offset] ?? emptyCell
  const readNumberAt = (offset: number): number => {
    const value = readCellValueAt(offset)
    if (value.tag === ValueTag.Number) {
      return value.value
    }
    if (value.tag === ValueTag.Boolean) {
      return value.value ? 1 : 0
    }
    return 0
  }
  return {
    owner,
    sheetName: owner.sheetName,
    rowStart: 0,
    rowEnd: Math.max(0, cells.length - 1),
    col: owner.col,
    length: cells.length,
    columnVersion: owner.columnVersion,
    structureVersion: owner.structureVersion,
    sheetColumnVersions: owner.sheetColumnVersions,
    readTagAt: (offset) => readCellValueAt(offset).tag,
    readNumberAt,
    readStringIdAt: () => 0,
    readErrorAt: (offset) => {
      const value = readCellValueAt(offset)
      return value.tag === ValueTag.Error ? value.code : ErrorCode.None
    },
    readCellValueAt,
  }
}

function makeColumnStore(view: RuntimeColumnView): EngineRuntimeColumnStoreService {
  return {
    getColumnOwner: () => view.owner,
    getColumnView: () => view,
    getColumnSlice: () => ({
      sheetName: view.sheetName,
      rowStart: view.rowStart,
      rowEnd: view.rowEnd,
      col: view.col,
      length: view.length,
      columnVersion: view.columnVersion,
      structureVersion: view.structureVersion,
      sheetColumnVersions: view.sheetColumnVersions,
      tags: new Uint8Array(view.length),
      numbers: new Float64Array(view.length),
      stringIds: new Uint32Array(view.length),
      errors: new Uint16Array(view.length),
    }),
    readCellValue: (_sheetName, row) => view.readCellValueAt(row - view.rowStart),
    readRangeValues: () => [],
    readRangeValueMatrix: () => [],
    findMaxResidentRowInColumns: () => view.rowEnd,
    normalizeStringId: (stringId) => String(stringId),
    normalizeLookupText: (value) => value.value.toUpperCase(),
  }
}

function directIndexCriteria(offsetValue: number): RuntimeDirectCriteriaDescriptor {
  return {
    aggregateKind: 'first',
    aggregateRange: {
      regionId: 1,
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
      length: 3,
    },
    offsetOperand: { kind: 'literal-number', value: offsetValue },
    criteriaPairs: [],
  }
}

function compiledFormula(): CompiledFormula {
  return {
    id: 0,
    source: '0',
    mode: FormulaMode.JsOnly,
    depsPtr: 0,
    depsLen: 0,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    maxStackDepth: 0,
    ast: numberAst,
    optimizedAst: numberAst,
    deps: [],
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    volatile: false,
    randCallCount: 0,
    producesSpill: false,
    jsPlan: [],
    program: new Uint32Array(),
    constants: new Float64Array(),
    symbolicRefs: [],
    symbolicRanges: [],
    symbolicStrings: [],
  }
}

function runtimeFormula(
  overrides: Pick<Partial<RuntimeFormula>, 'directAggregate' | 'directScalar' | 'dependencyIndices'> = {},
): RuntimeFormula {
  const compiled = compiledFormula()
  return {
    cellIndex: 0,
    formulaSlotId: 0,
    planId: 0,
    templateId: undefined,
    source: compiled.source,
    compiled,
    plan: { id: 0, source: compiled.source, compiled },
    dependencyIndices: new Uint32Array(),
    dependencyEntities: { ptr: -1, len: 0, cap: 0 },
    rangeDependencies: new Uint32Array(),
    graphRangeDependencies: new Uint32Array(),
    runtimeProgram: new Uint32Array(),
    constants: new Float64Array(),
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar: undefined,
    directCriteria: undefined,
    ...overrides,
  }
}

function runtimeFormulaWithDirectScalar(directScalar: RuntimeDirectScalarDescriptor | undefined): RuntimeFormula {
  return runtimeFormula({ directScalar })
}

function runtimeFormulaWithDirectAggregate(directAggregate: RuntimeDirectAggregateDescriptor): RuntimeFormula {
  return runtimeFormula({ directAggregate, dependencyIndices: new Uint32Array() })
}

function stringifyCriteriaValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.Empty:
      return ''
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return `#${value.code}`
  }
}

describe('direct formula evaluation helpers', () => {
  it('classifies numeric-like column values for direct criteria aggregation', () => {
    const view = makeColumnView([
      valueCell(4),
      booleanCell(true),
      booleanCell(false),
      emptyCell,
      stringCell('4'),
      errorCell(ErrorCode.Value),
    ])

    expect(numericLikeValueInView(view, 0)).toBe(4)
    expect(numericLikeValueInView(view, 1)).toBe(1)
    expect(numericLikeValueInView(view, 2)).toBe(0)
    expect(numericLikeValueInView(view, 3)).toBe(0)
    expect(numericLikeValueInView(view, 4)).toBeUndefined()
    expect(numericLikeValueInView(view, 5)).toBeUndefined()
    expect(strictNumericAggregateCandidateInView(view, 0)).toBe(4)
    expect(strictNumericAggregateCandidateInView(view, 1)).toBeUndefined()
  })

  it('evaluates direct INDEX offsets and rejects invalid offsets', () => {
    const runtimeColumnStore = makeColumnStore(makeColumnView([valueCell(10), valueCell(20), valueCell(30)]))
    const readCellValueByIndex = (): CellValue => emptyCell

    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: directIndexCriteria(2),
        runtimeColumnStore,
        readCellValueByIndex,
      }),
    ).toEqual(valueCell(20))
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: directIndexCriteria(0),
        runtimeColumnStore,
        readCellValueByIndex,
      }),
    ).toEqual(errorCell(ErrorCode.Ref))
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: directIndexCriteria(Number.POSITIVE_INFINITY),
        runtimeColumnStore,
        readCellValueByIndex,
      }),
    ).toEqual(errorCell(ErrorCode.Value))
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: directIndexCriteria(4),
        runtimeColumnStore,
        readCellValueByIndex,
      }),
    ).toEqual(errorCell(ErrorCode.Ref))
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: directIndexCriteria(0),
        runtimeColumnStore,
        readCellValueByIndex,
        ownerRow: 1,
      }),
    ).toEqual(valueCell(20))
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: directIndexCriteria(0),
        runtimeColumnStore,
        readCellValueByIndex,
        ownerRow: 9,
      }),
    ).toEqual(emptyCell)
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: {
          ...directIndexCriteria(1),
          offsetOperand: { kind: 'error', code: ErrorCode.NA },
        },
        runtimeColumnStore,
        readCellValueByIndex,
      }),
    ).toEqual(errorCell(ErrorCode.NA))
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: {
          ...directIndexCriteria(1),
          offsetOperand: undefined,
        },
        runtimeColumnStore,
        readCellValueByIndex,
      }),
    ).toBeUndefined()
    expect(
      tryEvaluateDirectIndexOffset({
        directCriteria: {
          ...directIndexCriteria(1),
          aggregateRange: undefined,
        },
        runtimeColumnStore,
        readCellValueByIndex,
      }),
    ).toBeUndefined()
  })

  it('coerces direct INDEX offset cell operands with spreadsheet semantics', () => {
    const runtimeColumnStore = makeColumnStore(makeColumnView([valueCell(10), valueCell(20), valueCell(30)]))
    const directCriteria: RuntimeDirectCriteriaDescriptor = {
      ...directIndexCriteria(1),
      offsetOperand: { kind: 'cell', cellIndex: 7 },
    }

    const evaluateWithOffsetCell = (value: CellValue) =>
      tryEvaluateDirectIndexOffset({
        directCriteria,
        runtimeColumnStore,
        readCellValueByIndex: () => value,
      })

    expect(evaluateWithOffsetCell(valueCell(3))).toEqual(valueCell(30))
    expect(evaluateWithOffsetCell(booleanCell(true))).toEqual(valueCell(10))
    expect(evaluateWithOffsetCell(booleanCell(false))).toEqual(errorCell(ErrorCode.Ref))
    expect(evaluateWithOffsetCell(emptyCell)).toEqual(errorCell(ErrorCode.Ref))
    expect(evaluateWithOffsetCell(errorCell(ErrorCode.Name))).toEqual(errorCell(ErrorCode.Name))
    expect(evaluateWithOffsetCell(stringCell('2'))).toEqual(errorCell(ErrorCode.Value))
  })

  it('only handles direct INDEX exact matches when the lookup index can answer', () => {
    const runtimeColumnStore = makeColumnStore(makeColumnView([valueCell(10), valueCell(20), valueCell(30)]))
    const directCriteria: RuntimeDirectCriteriaDescriptor = {
      aggregateKind: 'first',
      firstMatchMode: 'exact-lookup',
      aggregateRange: {
        regionId: 1,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 2,
        col: 0,
        length: 3,
      },
      criteriaPairs: [
        {
          range: {
            regionId: 2,
            sheetName: 'Sheet1',
            rowStart: 0,
            rowEnd: 2,
            col: 1,
            length: 3,
          },
          criterion: { kind: 'literal', value: valueCell(20) },
        },
      ],
    }
    const prepared = {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 1,
      length: 3,
      columnVersion: 1,
      structureVersion: 1,
      sheetColumnVersions: new Uint32Array([1]),
      comparableKind: 'numeric' as const,
      uniformStart: undefined,
      uniformStep: undefined,
      firstPositions: new Map<string, number>(),
      lastPositions: new Map<string, number>(),
      firstNumericPositions: undefined,
      lastNumericPositions: undefined,
      firstTextPositions: undefined,
      lastTextPositions: undefined,
    }

    expect(
      tryEvaluateDirectIndexExactMatch({
        directCriteria,
        runtimeColumnStore,
        lookupValue: valueCell(20),
        exactLookup: {
          prepareVectorLookup: () => prepared,
          findPreparedVectorMatch: () => ({ handled: false }),
        },
      }),
    ).toBeUndefined()
    expect(
      tryEvaluateDirectIndexExactMatch({
        directCriteria,
        runtimeColumnStore,
        lookupValue: valueCell(20),
        exactLookup: {
          prepareVectorLookup: () => prepared,
          findPreparedVectorMatch: () => ({ handled: true, position: undefined }),
        },
      }),
    ).toEqual(errorCell(ErrorCode.NA))
    expect(
      tryEvaluateDirectIndexExactMatch({
        directCriteria,
        runtimeColumnStore,
        lookupValue: valueCell(20),
        exactLookup: {
          prepareVectorLookup: () => prepared,
          findPreparedVectorMatch: () => ({ handled: true, position: 3 }),
        },
      }),
    ).toEqual(valueCell(30))
  })

  it('evaluates direct scalar descriptors with numeric coercion and error propagation', () => {
    const readCellValueByIndex = (cellIndex: number | undefined): CellValue => {
      switch (cellIndex) {
        case undefined:
          return emptyCell
        case 1:
          return stringCell('6')
        case 2:
          return stringCell('not numeric')
        case 3:
          return booleanCell(true)
        case 4:
          return emptyCell
        case 5:
          return errorCell(ErrorCode.Div0)
        default:
          return emptyCell
      }
    }

    expect(tryEvaluateDirectScalar(runtimeFormulaWithDirectScalar(undefined), readCellValueByIndex)).toBeUndefined()
    expect(
      tryEvaluateDirectScalar(
        runtimeFormulaWithDirectScalar({ kind: 'abs', operand: { kind: 'literal-number', value: -4 } }),
        readCellValueByIndex,
      ),
    ).toEqual(valueCell(4))
    expect(
      tryEvaluateDirectScalar(
        runtimeFormulaWithDirectScalar({ kind: 'abs', operand: { kind: 'error', code: ErrorCode.NA } }),
        readCellValueByIndex,
      ),
    ).toEqual(errorCell(ErrorCode.NA))
    expect(
      tryEvaluateDirectScalar(
        runtimeFormulaWithDirectScalar({
          kind: 'binary',
          operator: '+',
          left: { kind: 'cell', cellIndex: 1 },
          right: { kind: 'cell', cellIndex: 3 },
          resultOffset: 2,
        }),
        readCellValueByIndex,
      ),
    ).toEqual(valueCell(9))
    expect(
      tryEvaluateDirectScalar(
        runtimeFormulaWithDirectScalar({
          kind: 'binary',
          operator: '*',
          left: { kind: 'cell', cellIndex: 4 },
          right: { kind: 'literal-number', value: 8 },
        }),
        readCellValueByIndex,
      ),
    ).toEqual(valueCell(0))
    expect(
      tryEvaluateDirectScalar(
        runtimeFormulaWithDirectScalar({
          kind: 'binary',
          operator: '-',
          left: { kind: 'cell', cellIndex: 2 },
          right: { kind: 'literal-number', value: 1 },
        }),
        readCellValueByIndex,
      ),
    ).toEqual(errorCell(ErrorCode.Value))
    expect(
      tryEvaluateDirectScalar(
        runtimeFormulaWithDirectScalar({
          kind: 'binary',
          operator: '+',
          left: { kind: 'literal-number', value: 1 },
          right: { kind: 'cell', cellIndex: 5 },
        }),
        readCellValueByIndex,
      ),
    ).toEqual(errorCell(ErrorCode.Div0))
    expect(
      tryEvaluateDirectScalar(
        runtimeFormulaWithDirectScalar({
          kind: 'binary',
          operator: '/',
          left: { kind: 'literal-number', value: 1 },
          right: { kind: 'literal-number', value: 0 },
        }),
        readCellValueByIndex,
      ),
    ).toEqual(errorCell(ErrorCode.Div0))
  })

  it('uses page summaries for large direct aggregate windows without scanning cells', () => {
    const counters = createEngineCounters()
    const aggregateCache: RangeAggregateCacheService = {
      getOrBuildPrefix: () => {
        throw new Error('unexpected prefix build')
      },
      getOrBuildColumnPrefix: () => {
        throw new Error('unexpected column prefix build')
      },
      hasReusableColumnPrefix: () => false,
      summarizeColumnWindow: (request) => ({
        sum: request.col === 0 ? 40 : 20,
        count: request.col === 0 ? 4 : 2,
        averageCount: request.col === 0 ? 4 : 2,
        minimum: request.col === 0 ? 1 : 2,
        maximum: request.col === 0 ? 20 : 12,
        errorCount: 0,
        errorCode: ErrorCode.None,
        fullPageHits: 2,
        edgeCellScans: 0,
      }),
    }

    expect(
      tryEvaluateDirectAggregate({
        formula: runtimeFormulaWithDirectAggregate({
          regionId: 1,
          aggregateKind: 'average',
          sheetName: 'Sheet1',
          rowStart: 2048,
          rowEnd: 2303,
          col: 0,
          colEnd: 1,
          length: 256,
        }),
        workbook: {
          getSheet: () => undefined,
        },
        counters,
        aggregateCache,
        readCellValueByIndex: () => {
          throw new Error('unexpected cell scan')
        },
      }),
    ).toEqual(valueCell(10))
    expect(counters.directAggregatePageEvaluations).toBe(1)
    expect(counters.directAggregatePageFullHits).toBe(4)
    expect(counters.directAggregatePageEdgeCells).toBe(0)
  })

  it('normalizes direct criteria operands from literals, cells, and concat variants', () => {
    const readCellValueByIndex = (cellIndex: number | undefined): CellValue => {
      switch (cellIndex) {
        case undefined:
          return emptyCell
        case 1:
          return valueCell(45123)
        case 2:
          return errorCell(ErrorCode.Value)
        case 3:
          return stringCell('plain')
        case 4:
          return stringCell('not a date')
        default:
          return emptyCell
      }
    }

    expect(
      readRuntimeDirectCriteriaOperandValue({
        operand: { kind: 'literal', value: stringCell('literal') },
        readCellValueByIndex,
        stringifyCriteriaValue,
      }),
    ).toEqual(stringCell('literal'))
    expect(
      readRuntimeDirectCriteriaOperandValue({
        operand: { kind: 'cell', cellIndex: 3 },
        readCellValueByIndex,
        stringifyCriteriaValue,
      }),
    ).toEqual(stringCell('plain'))
    expect(
      readRuntimeDirectCriteriaOperandValue({
        operand: { kind: 'cell-string-concat', cellIndex: 3, prefix: '*', suffix: '?' },
        readCellValueByIndex,
        stringifyCriteriaValue,
      }),
    ).toEqual(stringCell('*plain?'))
    expect(
      readRuntimeDirectCriteriaOperandValue({
        operand: { kind: 'cell-string-concat', cellIndex: 2, prefix: '*', suffix: '?' },
        readCellValueByIndex,
        stringifyCriteriaValue,
      }),
    ).toEqual(errorCell(ErrorCode.Value))
    expect(
      readRuntimeDirectCriteriaOperandValue({
        operand: { kind: 'cell-month-boundary-string-concat', cellIndex: 1, prefix: '>=', suffix: '', offsetMonths: 1 },
        readCellValueByIndex,
        stringifyCriteriaValue,
      }),
    ).toMatchObject({ tag: ValueTag.String, value: expect.stringMatching(/^>=\d+$/) })
    expect(
      readRuntimeDirectCriteriaOperandValue({
        operand: { kind: 'cell-month-boundary-string-concat', cellIndex: 2, prefix: '>=', suffix: '', offsetMonths: 1 },
        readCellValueByIndex,
        stringifyCriteriaValue,
      }),
    ).toEqual(errorCell(ErrorCode.Value))
    expect(
      readRuntimeDirectCriteriaOperandValue({
        operand: { kind: 'cell-month-boundary-string-concat', cellIndex: 4, prefix: '>=', suffix: '', offsetMonths: 0 },
        readCellValueByIndex,
        stringifyCriteriaValue,
      }),
    ).toEqual(errorCell(ErrorCode.Value))
  })
})
