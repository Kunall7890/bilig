import type { EdgeArena, EdgeSlice } from '../../edge-arena.js'
import { entityPayload, isExactLookupColumnEntity, isRangeEntity, isSortedLookupColumnEntity } from '../../entity-ids.js'

export interface FormulaBindingReverseEdgeState {
  reverseCellEdges: Array<EdgeSlice | undefined>
  reverseRangeEdges: Array<EdgeSlice | undefined>
  reverseDefinedNameEdges: Map<string, Set<number>>
  reverseTableEdges: Map<string, Set<number>>
  reverseSpillEdges: Map<string, Set<number>>
  reverseAggregateColumnEdges: Map<number, Set<number>>
  reverseExactLookupColumnEdges: Map<number, EdgeSlice>
  reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
}

export function setFormulaBindingReverseEdgeSlice(state: FormulaBindingReverseEdgeState, entityId: number, slice: EdgeSlice): void {
  const empty = slice.ptr < 0 || slice.len === 0
  if (isRangeEntity(entityId)) {
    state.reverseRangeEdges[entityPayload(entityId)] = empty ? undefined : slice
    return
  }
  if (isExactLookupColumnEntity(entityId)) {
    if (empty) {
      state.reverseExactLookupColumnEdges.delete(entityPayload(entityId))
    } else {
      state.reverseExactLookupColumnEdges.set(entityPayload(entityId), slice)
    }
    return
  }
  if (isSortedLookupColumnEntity(entityId)) {
    if (empty) {
      state.reverseSortedLookupColumnEdges.delete(entityPayload(entityId))
    } else {
      state.reverseSortedLookupColumnEdges.set(entityPayload(entityId), slice)
    }
    return
  }
  state.reverseCellEdges[entityPayload(entityId)] = empty ? undefined : slice
}

export function getFormulaBindingReverseEdgeSlice(state: FormulaBindingReverseEdgeState, entityId: number): EdgeSlice | undefined {
  if (isRangeEntity(entityId)) {
    return state.reverseRangeEdges[entityPayload(entityId)]
  }
  if (isExactLookupColumnEntity(entityId)) {
    return state.reverseExactLookupColumnEdges.get(entityPayload(entityId))
  }
  if (isSortedLookupColumnEntity(entityId)) {
    return state.reverseSortedLookupColumnEdges.get(entityPayload(entityId))
  }
  return state.reverseCellEdges[entityPayload(entityId)]
}

export function appendFormulaBindingReverseEdge(
  state: FormulaBindingReverseEdgeState,
  edgeArena: EdgeArena,
  entityId: number,
  dependentEntityId: number,
): void {
  const slice = getFormulaBindingReverseEdgeSlice(state, entityId) ?? edgeArena.empty()
  setFormulaBindingReverseEdgeSlice(state, entityId, edgeArena.appendUnique(slice, dependentEntityId))
}

export function removeFormulaBindingReverseEdge(
  state: FormulaBindingReverseEdgeState,
  edgeArena: EdgeArena,
  entityId: number,
  dependentEntityId: number,
): void {
  const slice = getFormulaBindingReverseEdgeSlice(state, entityId)
  if (!slice) {
    return
  }
  setFormulaBindingReverseEdgeSlice(state, entityId, edgeArena.removeValue(slice, dependentEntityId))
}

export function syncFormulaBindingRangeDependencyEdges(
  state: FormulaBindingReverseEdgeState,
  edgeArena: EdgeArena,
  rangeEntity: number,
  deps: { oldDependencySources: Uint32Array; newDependencySources: Uint32Array },
): void {
  const nextSources = new Set<number>(deps.newDependencySources)
  deps.oldDependencySources.forEach((dependencyEntity) => {
    if (!nextSources.has(dependencyEntity)) {
      removeFormulaBindingReverseEdge(state, edgeArena, dependencyEntity, rangeEntity)
    }
  })
  const priorSources = new Set<number>(deps.oldDependencySources)
  deps.newDependencySources.forEach((dependencyEntity) => {
    if (!priorSources.has(dependencyEntity)) {
      appendFormulaBindingReverseEdge(state, edgeArena, dependencyEntity, rangeEntity)
    }
  })
}
