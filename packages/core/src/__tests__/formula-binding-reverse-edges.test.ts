import { describe, expect, it } from 'vitest'
import { EdgeArena } from '../edge-arena.js'
import { entityPayload, makeCellEntity, makeExactLookupColumnEntity, makeRangeEntity, makeSortedLookupColumnEntity } from '../entity-ids.js'
import {
  appendFormulaBindingReverseEdge,
  getFormulaBindingReverseEdgeSlice,
  removeFormulaBindingReverseEdge,
  setFormulaBindingReverseEdgeSlice,
  syncFormulaBindingRangeDependencyEdges,
  type FormulaBindingReverseEdgeState,
} from '../engine/services/formula-binding-reverse-edges.js'

function createReverseState(): FormulaBindingReverseEdgeState {
  return {
    reverseCellEdges: [],
    reverseRangeEdges: [],
    reverseDefinedNameEdges: new Map(),
    reverseTableEdges: new Map(),
    reverseSpillEdges: new Map(),
    reverseAggregateColumnEdges: new Map(),
    reverseExactLookupColumnEdges: new Map(),
    reverseSortedLookupColumnEdges: new Map(),
  }
}

describe('formula binding reverse edges', () => {
  it('stores edge slices in the registry matching the dependency entity kind', () => {
    const edgeArena = new EdgeArena()
    const state = createReverseState()
    const cellEntity = makeCellEntity(2)
    const rangeEntity = makeRangeEntity(3)
    const exactLookupEntity = makeExactLookupColumnEntity(1, 4)
    const sortedLookupEntity = makeSortedLookupColumnEntity(1, 5)

    appendFormulaBindingReverseEdge(state, edgeArena, cellEntity, 10)
    appendFormulaBindingReverseEdge(state, edgeArena, rangeEntity, 11)
    appendFormulaBindingReverseEdge(state, edgeArena, exactLookupEntity, 12)
    appendFormulaBindingReverseEdge(state, edgeArena, sortedLookupEntity, 13)

    expect(edgeArena.read(getFormulaBindingReverseEdgeSlice(state, cellEntity)!)).toEqual(Uint32Array.from([10]))
    expect(edgeArena.read(state.reverseCellEdges[2]!)).toEqual(Uint32Array.from([10]))
    expect(edgeArena.read(state.reverseRangeEdges[3]!)).toEqual(Uint32Array.from([11]))
    expect(edgeArena.read(state.reverseExactLookupColumnEdges.get(entityPayload(exactLookupEntity))!)).toEqual(Uint32Array.from([12]))
    expect(edgeArena.read(state.reverseSortedLookupColumnEdges.get(entityPayload(sortedLookupEntity))!)).toEqual(Uint32Array.from([13]))
  })

  it('deduplicates appended edges and removes empty slices from keyed maps', () => {
    const edgeArena = new EdgeArena()
    const state = createReverseState()
    const exactLookupEntity = makeExactLookupColumnEntity(2, 7)

    appendFormulaBindingReverseEdge(state, edgeArena, exactLookupEntity, 20)
    appendFormulaBindingReverseEdge(state, edgeArena, exactLookupEntity, 20)
    expect(edgeArena.read(getFormulaBindingReverseEdgeSlice(state, exactLookupEntity)!)).toEqual(Uint32Array.from([20]))

    removeFormulaBindingReverseEdge(state, edgeArena, exactLookupEntity, 20)
    expect(getFormulaBindingReverseEdgeSlice(state, exactLookupEntity)).toBeUndefined()

    const slice = edgeArena.replace(edgeArena.empty(), Uint32Array.from([30]))
    setFormulaBindingReverseEdgeSlice(state, makeSortedLookupColumnEntity(2, 8), slice)
    setFormulaBindingReverseEdgeSlice(state, makeSortedLookupColumnEntity(2, 8), edgeArena.empty())
    expect(getFormulaBindingReverseEdgeSlice(state, makeSortedLookupColumnEntity(2, 8))).toBeUndefined()
  })

  it('syncs range dependency edges by removing stale sources and appending new sources', () => {
    const edgeArena = new EdgeArena()
    const state = createReverseState()
    const rangeEntity = makeRangeEntity(9)
    const staleDependency = makeCellEntity(1)
    const retainedDependency = makeCellEntity(2)
    const newDependency = makeCellEntity(3)

    appendFormulaBindingReverseEdge(state, edgeArena, staleDependency, rangeEntity)
    appendFormulaBindingReverseEdge(state, edgeArena, retainedDependency, rangeEntity)

    syncFormulaBindingRangeDependencyEdges(state, edgeArena, rangeEntity, {
      oldDependencySources: Uint32Array.from([staleDependency, retainedDependency]),
      newDependencySources: Uint32Array.from([retainedDependency, newDependency]),
    })

    expect(getFormulaBindingReverseEdgeSlice(state, staleDependency)).toBeUndefined()
    expect(edgeArena.read(getFormulaBindingReverseEdgeSlice(state, retainedDependency)!)).toEqual(Uint32Array.from([rangeEntity]))
    expect(edgeArena.read(getFormulaBindingReverseEdgeSlice(state, newDependency)!)).toEqual(Uint32Array.from([rangeEntity]))
  })
})
