import type { CompiledFormula } from '@bilig/formula'
import { FormulaMode } from '@bilig/protocol'
import type { EdgeSlice } from '../../edge-arena.js'
import {
  entityPayload,
  isRangeEntity,
  makeCellEntity,
  makeExactLookupColumnEntity,
  makeRangeEntity,
  makeSortedLookupColumnEntity,
} from '../../entity-ids.js'
import { spillDependencyKeyFromRef, tableDependencyKey } from '../../engine-metadata-utils.js'
import type { RuntimeFormula } from '../runtime-state.js'
import {
  aggregateColumnDependencyKey,
  directCriteriaAggregateColumn,
  directLookupColumnInfo,
  removeDirectAggregateColumnReverseEdges,
  removeDirectCriteriaAggregateColumnReverseEdge,
  removeTrackedReverseEdge,
} from './formula-binding-dependency-helpers.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { FormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'
import { clearFormulaRuntimeFlags } from './formula-binding-cell-flags.js'

export function clearFormulaBindingNow(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly formulaMemberCounts: FormulaBindingMemberCounts
  readonly untrackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string | undefined,
    compiled: Pick<CompiledFormula, 'deps'> | undefined,
  ) => void
  readonly removeReverseEdge: (entityId: number, dependentEntityId: number) => void
  readonly setReverseEdgeSlice: (entityId: number, slice: EdgeSlice) => void
  readonly pruneTrackedDependencyCell: (cellIndex: number, ownerCellIndex: number) => void
  readonly updateVolatileFormulaIndex: (cellIndex: number, formula: RuntimeFormula | undefined) => void
  readonly cellIndex: number
}): boolean {
  const serviceArgs = args.serviceArgs
  const existing = serviceArgs.state.formulas.get(args.cellIndex)
  const needsWasmProgramSync = existing !== undefined && existing.compiled.mode === FormulaMode.WasmFastPath && existing.programLength > 0
  if (existing) {
    serviceArgs.regionGraph.clearFormulaSubscriptions(args.cellIndex)
    const ownerSheetName = serviceArgs.state.workbook.getSheetNameById(serviceArgs.state.workbook.cellStore.sheetIds[args.cellIndex]!)
    args.untrackFormulaSheetIndexes(args.cellIndex, ownerSheetName, existing.compiled)
    const sheetId = serviceArgs.state.workbook.cellStore.sheetIds[args.cellIndex]
    const col = serviceArgs.state.workbook.getCellPosition(args.cellIndex)?.col
    if (sheetId !== undefined) {
      args.formulaMemberCounts.decrement(sheetId, col)
    }
    const dependencyEntities = serviceArgs.edgeArena.readView(existing.dependencyEntities)
    const formulaEntity = makeCellEntity(args.cellIndex)
    for (let index = 0; index < dependencyEntities.length; index += 1) {
      const dependencyEntity = dependencyEntities[index]!
      args.removeReverseEdge(dependencyEntity, formulaEntity)
      if (!isRangeEntity(dependencyEntity)) {
        args.pruneTrackedDependencyCell(entityPayload(dependencyEntity), args.cellIndex)
      }
    }
    existing.compiled.symbolicNames.forEach((name) => {
      removeTrackedReverseEdge(serviceArgs.reverseState.reverseDefinedNameEdges, normalizeDefinedName(name), args.cellIndex)
    })
    existing.compiled.symbolicTables.forEach((name) => {
      removeTrackedReverseEdge(serviceArgs.reverseState.reverseTableEdges, tableDependencyKey(name), args.cellIndex)
    })
    existing.compiled.symbolicSpills.forEach((key) => {
      removeTrackedReverseEdge(serviceArgs.reverseState.reverseSpillEdges, spillDependencyKeyFromRef(key, ownerSheetName), args.cellIndex)
    })
    const existingDirectLookup = existing.directLookup
    if (existingDirectLookup) {
      const lookupInfo = directLookupColumnInfo(existingDirectLookup)
      const lookupSheet = serviceArgs.state.workbook.getSheet(lookupInfo.sheetName)
      if (lookupSheet) {
        const lookupEntity = lookupInfo.isExact
          ? makeExactLookupColumnEntity(lookupSheet.id, lookupInfo.col)
          : makeSortedLookupColumnEntity(lookupSheet.id, lookupInfo.col)
        args.removeReverseEdge(lookupEntity, formulaEntity)
        if (existingDirectLookup.kind === 'approximate' || existingDirectLookup.kind === 'approximate-uniform-numeric') {
          const rowStart =
            existingDirectLookup.kind === 'approximate' ? existingDirectLookup.prepared.rowStart : existingDirectLookup.rowStart
          const rowEnd = existingDirectLookup.kind === 'approximate' ? existingDirectLookup.prepared.rowEnd : existingDirectLookup.rowEnd
          for (let row = rowStart; row <= rowEnd; row += 1) {
            const memberCellIndex = serviceArgs.ensureCellTrackedByCoords(lookupSheet.id, row, lookupInfo.col)
            args.removeReverseEdge(makeCellEntity(memberCellIndex), lookupEntity)
          }
        }
      }
    }
    const directCriteriaAggregate = directCriteriaAggregateColumn(existing.directCriteria)
    if (directCriteriaAggregate) {
      const aggregateSheet = serviceArgs.state.workbook.getSheet(directCriteriaAggregate.sheetName)
      if (aggregateSheet) {
        removeDirectCriteriaAggregateColumnReverseEdge(
          serviceArgs.reverseState.reverseAggregateColumnEdges,
          aggregateColumnDependencyKey(aggregateSheet.id, directCriteriaAggregate.col),
          args.cellIndex,
        )
      }
    }
    removeDirectAggregateColumnReverseEdges(
      serviceArgs.reverseState.reverseAggregateColumnEdges,
      serviceArgs.state.workbook,
      existing.directAggregate,
      args.cellIndex,
    )
    for (let index = 0; index < existing.rangeDependencies.length; index += 1) {
      const rangeIndex = existing.rangeDependencies[index]!
      const dependencySources = serviceArgs.state.ranges.getDependencySourceEntities(rangeIndex)
      const released = serviceArgs.state.ranges.release(rangeIndex)
      if (!released.removed) {
        continue
      }
      const rangeEntity = makeRangeEntity(rangeIndex)
      for (let sourceIndex = 0; sourceIndex < dependencySources.length; sourceIndex += 1) {
        const dependencyEntity = dependencySources[sourceIndex]!
        args.removeReverseEdge(dependencyEntity, rangeEntity)
        if (!isRangeEntity(dependencyEntity)) {
          args.pruneTrackedDependencyCell(entityPayload(dependencyEntity), args.cellIndex)
        }
      }
      args.setReverseEdgeSlice(rangeEntity, serviceArgs.edgeArena.empty())
    }
    serviceArgs.edgeArena.free(existing.dependencyEntities)
    serviceArgs.compiledPlans.release(existing.planId)
  }
  serviceArgs.formulaInstances.delete(args.cellIndex)
  serviceArgs.formulaFamilies.unregisterFormula(args.cellIndex)
  args.updateVolatileFormulaIndex(args.cellIndex, undefined)
  serviceArgs.state.formulas.delete(args.cellIndex)
  clearFormulaRuntimeFlags(serviceArgs.state.workbook.cellStore, args.cellIndex)
  if (needsWasmProgramSync) {
    serviceArgs.scheduleWasmProgramSync()
  }
  return existing !== undefined
}
