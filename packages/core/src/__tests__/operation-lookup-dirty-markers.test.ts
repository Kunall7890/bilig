import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { RuntimeDirectLookupDescriptor } from '../engine/runtime-state.js'
import { createOperationLookupDirtyMarkerService } from '../engine/services/operation-lookup-dirty-markers.js'
import { lookupImpactCacheKey } from '../engine/services/direct-formula-recalc-helpers.js'

function exactUniform(
  overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>> = {},
): Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }> {
  return {
    kind: 'exact-uniform-numeric',
    operandCellIndex: 1,
    sheetName: 'Sheet1',
    sheetId: 7,
    rowStart: 0,
    rowEnd: 4,
    col: 0,
    length: 5,
    columnVersion: 2,
    structureVersion: 3,
    sheetColumnVersions: new Uint32Array([2]),
    start: 1,
    step: 1,
    searchMode: 1,
    ...overrides,
  }
}

function approximateUniform(
  overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>> = {},
): Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }> {
  return {
    kind: 'approximate-uniform-numeric',
    operandCellIndex: 1,
    sheetName: 'Sheet1',
    sheetId: 7,
    rowStart: 0,
    rowEnd: 4,
    col: 0,
    length: 5,
    columnVersion: 2,
    structureVersion: 3,
    sheetColumnVersions: new Uint32Array([2]),
    start: 1,
    step: 1,
    matchMode: 1,
    ...overrides,
  }
}

function createService(
  request: {
    formulas?: ReadonlyMap<number, RuntimeDirectLookupDescriptor>
    dependents?: Uint32Array
    singleDependent?: number
    skipApproximate?: boolean
    operandValue?: { value: { tag: ValueTag.Number; value: number } | { tag: ValueTag.String; value: string }; stringId?: number }
  } = {},
) {
  const markFormulaChanged = vi.fn((cellIndex: number, count: number) => count + cellIndex)
  const noteExactLookupLiteralWrite = vi.fn()
  const noteSortedLookupLiteralWrite = vi.fn()
  const service = createOperationLookupDirtyMarkerService({
    state: {
      workbook: {
        getSheet: (sheetName: string) => (sheetName === 'Sheet1' ? { id: 7 } : undefined),
      },
      formulas: {
        get: (cellIndex: number) => {
          const directLookup = request.formulas?.get(cellIndex)
          return directLookup === undefined ? undefined : { directLookup }
        },
      },
      strings: {
        get: (id: number) => (id === 1 ? 'needle' : 'other'),
      },
    },
    getEntityDependents: () => request.dependents ?? new Uint32Array(0),
    getSingleEntityDependent: () => request.singleDependent ?? -2,
    markFormulaChanged,
    readCellValueForLookup: () => request.operandValue ?? { value: { tag: ValueTag.String, value: 'ignored' }, stringId: 1 },
    canSkipApproximateLookupDirtyMark: () => request.skipApproximate ?? false,
    noteExactLookupLiteralWrite,
    noteSortedLookupLiteralWrite,
    lookupImpactCacheKey,
  })
  return { service, markFormulaChanged, noteExactLookupLiteralWrite, noteSortedLookupLiteralWrite }
}

describe('operation lookup dirty markers', () => {
  it('marks a single numeric exact lookup dependent only when the changed value can affect the operand', () => {
    const { service, markFormulaChanged } = createService({
      formulas: new Map([[10, exactUniform()]]),
      singleDependent: 10,
      operandValue: { value: { tag: ValueTag.Number, value: 1 } },
    })

    expect(
      service.markAffectedExactLookupDependents(
        {
          sheetName: 'Sheet1',
          row: 2,
          col: 0,
          oldValue: { tag: ValueTag.Number, value: 9 },
          newValue: { tag: ValueTag.Number, value: 10 },
        },
        5,
        new Map(),
      ),
    ).toBe(5)
    expect(markFormulaChanged).not.toHaveBeenCalled()

    expect(
      service.markAffectedExactLookupDependents(
        {
          sheetName: 'Sheet1',
          row: 2,
          col: 0,
          oldValue: { tag: ValueTag.Number, value: 1 },
          newValue: { tag: ValueTag.Number, value: 9 },
        },
        5,
        new Map(),
      ),
    ).toBe(15)
    expect(markFormulaChanged).toHaveBeenCalledWith(10, 5)
  })

  it('caches exact lookup operand keys and marks matching dependents', () => {
    const { service, markFormulaChanged } = createService({
      formulas: new Map([
        [10, exactUniform()],
        [11, exactUniform({ rowStart: 4, rowEnd: 8 })],
      ]),
      dependents: Uint32Array.of(10, 11),
    })
    const caches = new Map()

    expect(
      service.markAffectedExactLookupDependents(
        {
          sheetName: 'Sheet1',
          row: 2,
          col: 0,
          oldValue: { tag: ValueTag.String, value: 'needle' },
          newValue: { tag: ValueTag.String, value: 'other' },
        },
        0,
        caches,
      ),
    ).toBe(10)
    expect(markFormulaChanged).toHaveBeenCalledTimes(1)
    expect(caches.get('7:0')).toBeDefined()
  })

  it('marks approximate dependents unless the lookup planner says the write is safe to skip', () => {
    const request = {
      sheetName: 'Sheet1',
      row: 2,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 2 },
      newValue: { tag: ValueTag.Number, value: 3 },
    }
    const marked = createService({
      formulas: new Map([[10, approximateUniform()]]),
      dependents: Uint32Array.of(10),
    })

    expect(marked.service.markAffectedApproximateLookupDependents(request, 1)).toBe(11)
    expect(marked.markFormulaChanged).toHaveBeenCalledWith(10, 1)

    const skipped = createService({
      formulas: new Map([[10, approximateUniform()]]),
      dependents: Uint32Array.of(10),
      skipApproximate: true,
    })
    expect(skipped.service.markAffectedApproximateLookupDependents(request, 1)).toBe(1)
    expect(skipped.markFormulaChanged).not.toHaveBeenCalled()
  })

  it('records exact and sorted lookup writes after dirty marking', () => {
    const exact = createService()
    const request = {
      sheetName: 'Sheet1',
      row: 2,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 2 },
      newValue: { tag: ValueTag.Number, value: 3 },
    }

    expect(exact.service.noteExactLookupLiteralWriteWhenDirty(request, 1, new Map())).toBe(1)
    expect(exact.noteExactLookupLiteralWrite).toHaveBeenCalledWith(request)

    const sorted = createService()
    expect(sorted.service.noteSortedLookupLiteralWriteWhenDirty(request, 1)).toBe(1)
    expect(sorted.noteSortedLookupLiteralWrite).toHaveBeenCalledWith(request)
  })
})
