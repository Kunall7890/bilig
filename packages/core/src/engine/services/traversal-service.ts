import { Effect } from 'effect'
import type { EdgeArena, EdgeSlice } from '../../edge-arena.js'
import {
  makeExactLookupColumnEntity,
  makeSortedLookupColumnEntity,
  entityPayload,
  isExactLookupColumnEntity,
  isRangeEntity,
  isSortedLookupColumnEntity,
} from '../../entity-ids.js'
import { growUint32 } from '../../engine-buffer-utils.js'
import type { EngineRuntimeState, RuntimeDirectCriteriaDescriptor, RuntimeDirectCriteriaOperand, U32 } from '../runtime-state.js'
import { EngineTraversalError } from '../errors.js'
import type { RegionGraph } from '../../deps/region-graph.js'
import { CellFlags } from '../../cell-store.js'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

const INITIAL_TRAVERSAL_SCRATCH_CAPACITY = 16
const DIRECT_CRITERIA_AGGREGATE_FORMULA_DEPENDENCY_CACHE_LIMIT = 16_384

interface DirectCriteriaAggregateFormulaDependencyCacheEntry {
  readonly canFilter: boolean
  readonly dependencies: readonly number[]
}

export interface EngineTraversalService {
  readonly getEntityDependents: (entityId: number) => Effect.Effect<Uint32Array, EngineTraversalError>
  readonly getSingleEntityDependent: (entityId: number) => Effect.Effect<number, EngineTraversalError>
  readonly collectFormulaDependents: (entityId: number) => Effect.Effect<Uint32Array, EngineTraversalError>
  readonly forEachFormulaDependencyCell: (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ) => Effect.Effect<void, EngineTraversalError>
  readonly forEachSheetCell: (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ) => Effect.Effect<void, EngineTraversalError>
  readonly getEntityDependentsNow: (entityId: number) => Uint32Array
  readonly getSingleEntityDependentNow: (entityId: number) => number
  readonly getCellDependentsNow: (cellIndex: number) => Uint32Array
  readonly getSingleCellDependentNow: (cellIndex: number) => number
  readonly collectFormulaDependentsNow: (entityId: number) => Uint32Array
  readonly forEachFormulaDependencyCellNow: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void
  readonly forEachSheetCellNow: (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void) => void
}

function traversalErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export function createEngineTraversalService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas' | 'ranges'>
  readonly regionGraph: Pick<RegionGraph, 'getRegion' | 'collectFormulaDependentsForCell'>
  readonly edgeArena: EdgeArena
  readonly reverseState: {
    reverseCellEdges: Array<EdgeSlice | undefined>
    reverseRangeEdges: Array<EdgeSlice | undefined>
    reverseExactLookupColumnEdges: Map<number, EdgeSlice>
    reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
  }
}): EngineTraversalService {
  const NO_ENTITY_DEPENDENT = -1
  const MULTIPLE_ENTITY_DEPENDENTS = -2
  let topoFormulaBuffer: U32 = new Uint32Array(INITIAL_TRAVERSAL_SCRATCH_CAPACITY)
  let topoEntityQueue: U32 = new Uint32Array(INITIAL_TRAVERSAL_SCRATCH_CAPACITY)
  let topoFormulaSeenEpoch = 1
  let topoRangeSeenEpoch = 1
  let topoExactLookupSeenEpoch = 1
  let topoSortedLookupSeenEpoch = 1
  let topoFormulaSeen: U32 = new Uint32Array(INITIAL_TRAVERSAL_SCRATCH_CAPACITY)
  let topoRangeSeen: U32 = new Uint32Array(INITIAL_TRAVERSAL_SCRATCH_CAPACITY)
  const topoExactLookupSeen = new Map<number, number>()
  const topoSortedLookupSeen = new Map<number, number>()
  let directRegionFormulaCacheVersion = -1
  const directRegionFormulaMaxCellCache = new Map<number, number | undefined>()
  const directCriteriaAggregateFormulaDependencyCache = new Map<string, DirectCriteriaAggregateFormulaDependencyCacheEntry>()

  const ensureTraversalScratchCapacity = (cellSize: number, entitySize: number, rangeSize: number): void => {
    if (cellSize > topoFormulaBuffer.length) {
      topoFormulaBuffer = growUint32(topoFormulaBuffer, cellSize)
    }
    if (cellSize > topoFormulaSeen.length) {
      topoFormulaSeen = growUint32(topoFormulaSeen, cellSize)
    }
    if (entitySize > topoEntityQueue.length) {
      topoEntityQueue = growUint32(topoEntityQueue, entitySize)
    }
    if (rangeSize > topoRangeSeen.length) {
      topoRangeSeen = growUint32(topoRangeSeen, rangeSize)
    }
  }

  const ensureEntityQueueCapacity = (size: number): void => {
    if (size <= topoEntityQueue.length) {
      return
    }
    let capacity = topoEntityQueue.length
    while (capacity < size) {
      capacity *= 2
    }
    topoEntityQueue = growUint32(topoEntityQueue, capacity)
  }

  const ensureFormulaBufferCapacity = (size: number): void => {
    if (size <= topoFormulaBuffer.length) {
      return
    }
    let capacity = topoFormulaBuffer.length
    while (capacity < size) {
      capacity *= 2
    }
    topoFormulaBuffer = growUint32(topoFormulaBuffer, capacity)
  }

  const getReverseEdgeSlice = (entityId: number): EdgeSlice | undefined => {
    if (isRangeEntity(entityId)) {
      return args.reverseState.reverseRangeEdges[entityPayload(entityId)]
    }
    if (isExactLookupColumnEntity(entityId)) {
      return args.reverseState.reverseExactLookupColumnEdges.get(entityPayload(entityId))
    }
    if (isSortedLookupColumnEntity(entityId)) {
      return args.reverseState.reverseSortedLookupColumnEdges.get(entityPayload(entityId))
    }
    return args.reverseState.reverseCellEdges[entityPayload(entityId)]
  }

  const getEntityDependentsNow = (entityId: number): Uint32Array =>
    args.edgeArena.readView(getReverseEdgeSlice(entityId) ?? args.edgeArena.empty())

  const getCellDependentsNow = (cellIndex: number): Uint32Array =>
    args.edgeArena.readView(args.reverseState.reverseCellEdges[cellIndex] ?? args.edgeArena.empty())

  const getSingleCellDependentNow = (cellIndex: number): number => {
    const slice = args.reverseState.reverseCellEdges[cellIndex]
    if (!slice || slice.len === 0 || slice.ptr < 0) {
      return NO_ENTITY_DEPENDENT
    }
    if (slice.len !== 1) {
      return MULTIPLE_ENTITY_DEPENDENTS
    }
    return args.edgeArena.valueAt(slice, 0)
  }

  const getSingleEntityDependentNow = (entityId: number): number => {
    const slice = getReverseEdgeSlice(entityId)
    if (!slice || slice.len === 0 || slice.ptr < 0) {
      return NO_ENTITY_DEPENDENT
    }
    if (slice.len !== 1) {
      return MULTIPLE_ENTITY_DEPENDENTS
    }
    return args.edgeArena.valueAt(slice, 0)
  }

  const forEachFormulaDependencyCellNow = (cellIndex: number, fn: (dependencyCellIndex: number) => void): void => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return
    }
    const seen = new Set<number>()
    const push = (dependencyCellIndex: number): void => {
      if (seen.has(dependencyCellIndex)) {
        return
      }
      seen.add(dependencyCellIndex)
      fn(dependencyCellIndex)
    }
    const cellMatchesLiteral = (candidateCellIndex: number, expected: CellValue): boolean | undefined => {
      const store = args.state.workbook.cellStore
      const tag = (candidateCellIndex === -1 ? ValueTag.Empty : (store.tags[candidateCellIndex] ?? ValueTag.Empty)) as ValueTag
      if (tag !== expected.tag) {
        return false
      }
      switch (expected.tag) {
        case ValueTag.Empty:
          return true
        case ValueTag.Number:
          return Object.is(candidateCellIndex === -1 ? 0 : (store.numbers[candidateCellIndex] ?? 0), expected.value)
        case ValueTag.Boolean:
          return (candidateCellIndex !== -1 && (store.numbers[candidateCellIndex] ?? 0) !== 0) === expected.value
        case ValueTag.Error:
          return (
            ((candidateCellIndex === -1 ? ErrorCode.None : (store.errors[candidateCellIndex] ?? ErrorCode.None)) as ErrorCode) ===
            expected.code
          )
        case ValueTag.String: {
          const expectedStringId = expected.stringId ?? 0
          return expectedStringId !== 0
            ? candidateCellIndex !== -1 && (store.stringIds[candidateCellIndex] ?? 0) === expectedStringId
            : undefined
        }
      }
    }
    const cellMatchesCriterion = (candidateCellIndex: number, criterion: RuntimeDirectCriteriaOperand): boolean | undefined => {
      if (criterion.kind === 'literal') {
        return cellMatchesLiteral(candidateCellIndex, criterion.value)
      }
      if (criterion.kind !== 'cell') {
        return undefined
      }
      const store = args.state.workbook.cellStore
      if (((store.flags[criterion.cellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
        return undefined
      }
      const expected = store.getValue(criterion.cellIndex, () => '')
      return cellMatchesLiteral(candidateCellIndex, expected)
    }
    const directCriteriaRangeVersionKey = (range: { sheetName: string; rowStart: number; rowEnd: number; col: number }): string => {
      const sheet = args.state.workbook.getSheet(range.sheetName)
      return `${range.sheetName}:${range.rowStart}:${range.rowEnd}:${range.col}:${sheet?.columnVersions[range.col] ?? 0}:${
        sheet?.structureVersion ?? 0
      }`
    }
    const cellValueCacheKey = (sourceCellIndex: number): string | undefined => {
      const store = args.state.workbook.cellStore
      const tag = (sourceCellIndex === -1 ? ValueTag.Empty : (store.tags[sourceCellIndex] ?? ValueTag.Empty)) as ValueTag
      switch (tag) {
        case ValueTag.Empty:
          return 'e:'
        case ValueTag.Number:
          return `n:${
            Object.is(sourceCellIndex === -1 ? 0 : (store.numbers[sourceCellIndex] ?? 0), -0) ? 0 : (store.numbers[sourceCellIndex] ?? 0)
          }`
        case ValueTag.Boolean:
          return sourceCellIndex !== -1 && (store.numbers[sourceCellIndex] ?? 0) !== 0 ? 'b:1' : 'b:0'
        case ValueTag.Error:
          return `r:${(sourceCellIndex === -1 ? ErrorCode.None : (store.errors[sourceCellIndex] ?? ErrorCode.None)) as ErrorCode}`
        case ValueTag.String: {
          const stringId = sourceCellIndex === -1 ? 0 : (store.stringIds[sourceCellIndex] ?? 0)
          return stringId === 0 ? undefined : `s:${stringId}`
        }
      }
    }
    const directCriteriaOperandCacheKey = (criterion: RuntimeDirectCriteriaOperand): string | undefined => {
      if (criterion.kind === 'literal') {
        if (criterion.value.tag === ValueTag.String && (criterion.value.stringId ?? 0) === 0) {
          return undefined
        }
        switch (criterion.value.tag) {
          case ValueTag.Empty:
            return 'l:e:'
          case ValueTag.Number:
            return `l:n:${Object.is(criterion.value.value, -0) ? 0 : criterion.value.value}`
          case ValueTag.Boolean:
            return criterion.value.value ? 'l:b:1' : 'l:b:0'
          case ValueTag.Error:
            return `l:r:${criterion.value.code}`
          case ValueTag.String:
            return `l:s:${criterion.value.stringId}`
        }
      }
      if (criterion.kind !== 'cell') {
        return undefined
      }
      const store = args.state.workbook.cellStore
      if (((store.flags[criterion.cellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
        return undefined
      }
      const valueKey = cellValueCacheKey(criterion.cellIndex)
      return valueKey === undefined ? undefined : `c:${criterion.cellIndex}:${valueKey}`
    }
    const directCriteriaAggregateFormulaDependencyCacheKey = (directCriteria: RuntimeDirectCriteriaDescriptor): string | undefined => {
      const aggregateRange = directCriteria.aggregateRange
      if (!aggregateRange) {
        return undefined
      }
      const pairKeys: string[] = []
      for (let index = 0; index < directCriteria.criteriaPairs.length; index += 1) {
        const pair = directCriteria.criteriaPairs[index]!
        const criterionKey = directCriteriaOperandCacheKey(pair.criterion)
        if (criterionKey === undefined) {
          return undefined
        }
        pairKeys.push(`${directCriteriaRangeVersionKey(pair.range)}:${criterionKey}`)
      }
      return `${args.state.formulas.version}\u0000${directCriteria.aggregateKind}\u0000${directCriteriaRangeVersionKey(
        aggregateRange,
      )}\u0000${pairKeys.join('|')}`
    }
    const rememberDirectCriteriaAggregateFormulaDependencyEntry = (
      key: string,
      entry: DirectCriteriaAggregateFormulaDependencyCacheEntry,
    ): DirectCriteriaAggregateFormulaDependencyCacheEntry => {
      if (directCriteriaAggregateFormulaDependencyCache.size >= DIRECT_CRITERIA_AGGREGATE_FORMULA_DEPENDENCY_CACHE_LIMIT) {
        const firstKey = directCriteriaAggregateFormulaDependencyCache.keys().next().value
        if (firstKey !== undefined) {
          directCriteriaAggregateFormulaDependencyCache.delete(firstKey)
        }
      }
      directCriteriaAggregateFormulaDependencyCache.set(key, entry)
      return entry
    }
    const pushMatchedDirectCriteriaAggregateFormulas = (directCriteria: RuntimeDirectCriteriaDescriptor): boolean => {
      const aggregateRange = directCriteria.aggregateRange
      if (!aggregateRange) {
        return true
      }
      const cacheKey = directCriteriaAggregateFormulaDependencyCacheKey(directCriteria)
      const cached = cacheKey === undefined ? undefined : directCriteriaAggregateFormulaDependencyCache.get(cacheKey)
      if (cached) {
        for (let index = 0; index < cached.dependencies.length; index += 1) {
          push(cached.dependencies[index]!)
        }
        return cached.canFilter
      }
      const limitingPair = directCriteria.criteriaPairs.find((pair) => cellMatchesCriterion(-1, pair.criterion) === false)
      if (!limitingPair) {
        return false
      }
      const aggregateSheet = args.state.workbook.getSheet(aggregateRange.sheetName)
      const limitingSheet = args.state.workbook.getSheet(limitingPair.range.sheetName)
      if (!aggregateSheet || !limitingSheet) {
        return false
      }
      let canFilter = true
      const matchedDependencies: number[] = []
      const visitRowOffset = (rowOffset: number): void => {
        if (!canFilter || rowOffset < 0 || rowOffset >= aggregateRange.length) {
          return
        }
        for (let pairIndex = 0; pairIndex < directCriteria.criteriaPairs.length; pairIndex += 1) {
          const pair = directCriteria.criteriaPairs[pairIndex]!
          if (rowOffset >= pair.range.length) {
            canFilter = false
            return
          }
          const pairSheet = args.state.workbook.getSheet(pair.range.sheetName)
          if (!pairSheet) {
            canFilter = false
            return
          }
          const criteriaCellIndex = pairSheet.grid.get(pair.range.rowStart + rowOffset, pair.range.col)
          const matches = cellMatchesCriterion(criteriaCellIndex, pair.criterion)
          if (matches === undefined) {
            canFilter = false
            return
          }
          if (!matches) {
            return
          }
        }
        const aggregateCellIndex = aggregateSheet.grid.get(aggregateRange.rowStart + rowOffset, aggregateRange.col)
        if (aggregateCellIndex !== -1 && ((args.state.workbook.cellStore.flags[aggregateCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
          matchedDependencies.push(aggregateCellIndex)
        }
      }
      limitingSheet.grid.forEachInRange(
        limitingPair.range.rowStart,
        limitingPair.range.col,
        limitingPair.range.rowEnd,
        limitingPair.range.col,
        (criteriaCellIndex) => {
          visitRowOffset((args.state.workbook.cellStore.rows[criteriaCellIndex] ?? 0) - limitingPair.range.rowStart)
        },
      )
      for (let index = 0; index < matchedDependencies.length; index += 1) {
        push(matchedDependencies[index]!)
      }
      if (cacheKey !== undefined) {
        rememberDirectCriteriaAggregateFormulaDependencyEntry(cacheKey, { canFilter, dependencies: matchedDependencies })
      }
      return canFilter
    }
    for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
      push(formula.dependencyIndices[index]!)
    }
    if (formula.directCriteria === undefined) {
      for (let index = 0; index < formula.graphRangeDependencies.length; index += 1) {
        const members = args.state.ranges.expandToCells(formula.graphRangeDependencies[index]!)
        for (let memberIndex = 0; memberIndex < members.length; memberIndex += 1) {
          push(members[memberIndex]!)
        }
      }
    }
    const maxFormulaCellInDirectRegion = (regionId: number): number | undefined => {
      const formulaVersion = args.state.formulas.version
      if (directRegionFormulaCacheVersion !== formulaVersion) {
        directRegionFormulaCacheVersion = formulaVersion
        directRegionFormulaMaxCellCache.clear()
      }
      if (directRegionFormulaMaxCellCache.has(regionId)) {
        return directRegionFormulaMaxCellCache.get(regionId)
      }
      const region = args.regionGraph.getRegion(regionId)
      const sheet = region ? args.state.workbook.getSheet(region.sheetName) : undefined
      if (!region || !sheet) {
        directRegionFormulaMaxCellCache.set(regionId, undefined)
        return undefined
      }
      let bestCellIndex: number | undefined
      let bestRank = -1
      for (let row = region.rowStart; row <= region.rowEnd; row += 1) {
        const dependencyCellIndex = sheet.grid.get(row, region.col)
        if (dependencyCellIndex === -1 || ((args.state.workbook.cellStore.flags[dependencyCellIndex] ?? 0) & CellFlags.HasFormula) === 0) {
          continue
        }
        const rank = args.state.workbook.cellStore.topoRanks[dependencyCellIndex] ?? 0
        if (bestCellIndex === undefined || rank > bestRank || (rank === bestRank && dependencyCellIndex > bestCellIndex)) {
          bestCellIndex = dependencyCellIndex
          bestRank = rank
        }
      }
      directRegionFormulaMaxCellCache.set(regionId, bestCellIndex)
      return bestCellIndex
    }
    const pushDirectRegion = (regionId: number | undefined): void => {
      if (regionId === undefined) {
        return
      }
      if (formula.directCriteria !== undefined) {
        const dependencyCellIndex = maxFormulaCellInDirectRegion(regionId)
        if (dependencyCellIndex !== undefined) {
          push(dependencyCellIndex)
        }
        return
      }
      const region = args.regionGraph.getRegion(regionId)
      if (!region) {
        return
      }
      const sheet = args.state.workbook.getSheet(region.sheetName)
      if (!sheet) {
        return
      }
      for (let row = region.rowStart; row <= region.rowEnd; row += 1) {
        const dependencyCellIndex = sheet.grid.get(row, region.col)
        if (dependencyCellIndex !== -1) {
          push(dependencyCellIndex)
        }
      }
    }
    const pushDirectLookupRange = (range: { sheetName: string; rowStart: number; rowEnd: number; col: number } | undefined): void => {
      if (!range) {
        return
      }
      const sheet = args.state.workbook.getSheet(range.sheetName)
      if (!sheet) {
        return
      }
      for (let row = range.rowStart; row <= range.rowEnd; row += 1) {
        const dependencyCellIndex = sheet.grid.get(row, range.col)
        if (dependencyCellIndex !== -1) {
          push(dependencyCellIndex)
        }
      }
    }
    pushDirectRegion(formula.directAggregate?.regionId)
    if (formula.directCriteria) {
      if (!pushMatchedDirectCriteriaAggregateFormulas(formula.directCriteria)) {
        pushDirectRegion(formula.directCriteria.aggregateRange?.regionId)
      }
      for (let index = 0; index < formula.directCriteria.criteriaPairs.length; index += 1) {
        const pair = formula.directCriteria.criteriaPairs[index]!
        pushDirectRegion(pair.range.regionId)
        if (pair.criterion.kind !== 'literal') {
          push(pair.criterion.cellIndex)
        }
      }
    }
    const directLookup = formula.directLookup
    if (directLookup) {
      if (directLookup.kind === 'exact' || directLookup.kind === 'approximate') {
        pushDirectLookupRange({
          sheetName: directLookup.prepared.sheetName,
          rowStart: directLookup.prepared.rowStart,
          rowEnd: directLookup.prepared.rowEnd,
          col: directLookup.prepared.col,
        })
        push(directLookup.operandCellIndex)
      } else {
        pushDirectLookupRange({
          sheetName: directLookup.sheetName,
          rowStart: directLookup.rowStart,
          rowEnd: directLookup.rowEnd,
          col: directLookup.col,
        })
        push(directLookup.operandCellIndex)
      }
    }
  }

  const forEachSheetCellNow = (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void): void => {
    const sheet = args.state.workbook.getSheetById(sheetId)
    if (!sheet) {
      return
    }
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      fn(cellIndex, row, col)
    })
  }

  const collectFormulaDependentsNow = (entityId: number): Uint32Array => {
    ensureTraversalScratchCapacity(
      Math.max(args.state.workbook.cellStore.size + 1, 1),
      Math.max(args.state.workbook.cellStore.size + args.state.ranges.size + 1, 1),
      Math.max(args.state.ranges.size + 1, 1),
    )

    topoFormulaSeenEpoch += 1
    if (topoFormulaSeenEpoch === 0xffff_ffff) {
      topoFormulaSeenEpoch = 1
      topoFormulaSeen.fill(0)
    }
    topoRangeSeenEpoch += 1
    if (topoRangeSeenEpoch === 0xffff_ffff) {
      topoRangeSeenEpoch = 1
      topoRangeSeen.fill(0)
    }
    topoExactLookupSeenEpoch += 1
    if (topoExactLookupSeenEpoch === 0xffff_ffff) {
      topoExactLookupSeenEpoch = 1
      topoExactLookupSeen.clear()
    }
    topoSortedLookupSeenEpoch += 1
    if (topoSortedLookupSeenEpoch === 0xffff_ffff) {
      topoSortedLookupSeenEpoch = 1
      topoSortedLookupSeen.clear()
    }

    let entityQueueLength = 1
    let formulaCount = 0
    topoEntityQueue[0] = entityId

    for (let queueIndex = 0; queueIndex < entityQueueLength; queueIndex += 1) {
      const currentEntity = topoEntityQueue[queueIndex]!
      if (!isRangeEntity(currentEntity) && !isExactLookupColumnEntity(currentEntity) && !isSortedLookupColumnEntity(currentEntity)) {
        const cellIndex = entityPayload(currentEntity)
        const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
        const position = args.state.workbook.getCellPosition(cellIndex)
        if (sheetId !== undefined && position) {
          const regionDependents = args.regionGraph.collectFormulaDependentsForCell(sheetId, position.row, position.col)
          for (let index = 0; index < regionDependents.length; index += 1) {
            const formulaCellIndex = regionDependents[index]!
            if (topoFormulaSeen[formulaCellIndex] === topoFormulaSeenEpoch) {
              continue
            }
            topoFormulaSeen[formulaCellIndex] = topoFormulaSeenEpoch
            ensureFormulaBufferCapacity(formulaCount + 1)
            topoFormulaBuffer[formulaCount] = formulaCellIndex
            formulaCount += 1
          }
        }
        if (sheetId !== undefined && position) {
          const exactLookupEntity = makeExactLookupColumnEntity(sheetId, position.col)
          const sortedLookupEntity = makeSortedLookupColumnEntity(sheetId, position.col)
          ensureEntityQueueCapacity(entityQueueLength + 2)
          topoEntityQueue[entityQueueLength] = exactLookupEntity
          entityQueueLength += 1
          topoEntityQueue[entityQueueLength] = sortedLookupEntity
          entityQueueLength += 1
        }
      }
      const dependents = getEntityDependentsNow(currentEntity)
      for (let index = 0; index < dependents.length; index += 1) {
        const dependent = dependents[index]!
        if (!(isRangeEntity(dependent) || isExactLookupColumnEntity(dependent) || isSortedLookupColumnEntity(dependent))) {
          const formulaCellIndex = entityPayload(dependent)
          if (topoFormulaSeen[formulaCellIndex] === topoFormulaSeenEpoch) {
            continue
          }
          topoFormulaSeen[formulaCellIndex] = topoFormulaSeenEpoch
          ensureFormulaBufferCapacity(formulaCount + 1)
          topoFormulaBuffer[formulaCount] = formulaCellIndex
          formulaCount += 1
          continue
        }
        if (isRangeEntity(dependent)) {
          const rangeIndex = entityPayload(dependent)
          if (topoRangeSeen[rangeIndex] === topoRangeSeenEpoch) {
            continue
          }
          topoRangeSeen[rangeIndex] = topoRangeSeenEpoch
        } else if (isExactLookupColumnEntity(dependent)) {
          const lookupColumnPayload = entityPayload(dependent)
          if (topoExactLookupSeen.get(lookupColumnPayload) === topoExactLookupSeenEpoch) {
            continue
          }
          topoExactLookupSeen.set(lookupColumnPayload, topoExactLookupSeenEpoch)
        } else {
          const lookupColumnPayload = entityPayload(dependent)
          if (topoSortedLookupSeen.get(lookupColumnPayload) === topoSortedLookupSeenEpoch) {
            continue
          }
          topoSortedLookupSeen.set(lookupColumnPayload, topoSortedLookupSeenEpoch)
        }
        ensureEntityQueueCapacity(entityQueueLength + 1)
        topoEntityQueue[entityQueueLength] = dependent
        entityQueueLength += 1
      }
    }

    return topoFormulaBuffer.subarray(0, formulaCount)
  }

  return {
    getEntityDependents(entityId) {
      return Effect.try({
        try: () => Uint32Array.from(getEntityDependentsNow(entityId)),
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to read entity dependents', cause),
            cause,
          }),
      })
    },
    getSingleEntityDependent(entityId) {
      return Effect.try({
        try: () => getSingleEntityDependentNow(entityId),
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to read single entity dependent', cause),
            cause,
          }),
      })
    },
    collectFormulaDependents(entityId) {
      return Effect.try({
        try: () => Uint32Array.from(collectFormulaDependentsNow(entityId)),
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to collect formula dependents', cause),
            cause,
          }),
      })
    },
    forEachFormulaDependencyCell(cellIndex, fn) {
      return Effect.try({
        try: () => {
          forEachFormulaDependencyCellNow(cellIndex, fn)
        },
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to iterate formula dependencies', cause),
            cause,
          }),
      })
    },
    forEachSheetCell(sheetId, fn) {
      return Effect.try({
        try: () => {
          forEachSheetCellNow(sheetId, fn)
        },
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage('Failed to iterate sheet cells', cause),
            cause,
          }),
      })
    },
    getEntityDependentsNow,
    getSingleEntityDependentNow,
    getCellDependentsNow,
    getSingleCellDependentNow,
    collectFormulaDependentsNow,
    forEachFormulaDependencyCellNow,
    forEachSheetCellNow,
  }
}
