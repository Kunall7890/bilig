import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeDirectCriteriaDescriptor } from '../engine/runtime-state.js'
import {
  numericLikeValueInView,
  strictNumericAggregateCandidateInView,
} from '../engine/services/formula-evaluation-direct-criteria-transforms.js'
import { tryEvaluateDirectIndexOffset } from '../engine/services/formula-evaluation-direct-index.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from '../engine/services/runtime-column-store-service.js'

const emptyCell: CellValue = { tag: ValueTag.Empty }
const valueCell = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const booleanCell = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const stringCell = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const errorCell = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

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
  })
})
