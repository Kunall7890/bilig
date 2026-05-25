import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'
import { errorValue } from '../../engine-value-utils.js'
import { mapStructuralAxisIndex, mapStructuralAxisInterval } from '../../engine-structural-utils.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { dependencyTouchesSheet, rangeDependencyAxisAffected, runtimeDirectRangeAxisAffected } from './structure-formula-rewrite-guards.js'
import { canDeferSimpleStructuralFormulaSource, classifySimpleDeleteStructuralFormulaSource } from './structure-formula-source-deferral.js'
import { isCellIndexMapped, structuralAxisIndexAffected } from './structure-runtime-cleanup.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'

export interface CollectStructuralFormulaImpactsOptions {
  readonly targetSheetId: number | undefined
  readonly transform: StructuralAxisTransform
  readonly sheetName: string
  readonly changedDefinedNames: ReadonlySet<string>
  readonly changedTableNames: ReadonlySet<string>
  readonly markDeferredStructuralFormulaSources: () => void
}

export interface StructuralFormulaImpacts {
  readonly formulaCellIndices: number[]
  readonly rebindCellIndices: number[]
  readonly preservedCellIndices: number[]
  readonly precomputedChangedInputCellIndices: number[]
  readonly ownerPositions: Map<number, { sheetName: string; row: number; col: number }>
  readonly precomputedDirectAggregateValueCellIndices: number[]
  readonly directAggregateRetargetCellIndices: number[]
}

const EMPTY_OWNER_POSITIONS = new Map<number, { sheetName: string; row: number; col: number }>()
const EMPTY_STRUCTURAL_FORMULA_IMPACTS: StructuralFormulaImpacts = {
  formulaCellIndices: [],
  rebindCellIndices: [],
  preservedCellIndices: [],
  precomputedChangedInputCellIndices: [],
  ownerPositions: EMPTY_OWNER_POSITIONS,
  precomputedDirectAggregateValueCellIndices: [],
  directAggregateRetargetCellIndices: [],
}

function tryCollectOwnedFamilyOnlyStructuralImpacts(
  args: CreateEngineStructureServiceArgs,
  argsForImpact: CollectStructuralFormulaImpactsOptions,
): StructuralFormulaImpacts | undefined {
  if (
    argsForImpact.targetSheetId === undefined ||
    argsForImpact.changedDefinedNames.size > 0 ||
    argsForImpact.changedTableNames.size > 0 ||
    argsForImpact.transform.kind === 'delete' ||
    argsForImpact.transform.axis !== 'column' ||
    !args.canUseFormulaFamilyIndex()
  ) {
    return undefined
  }
  const ownedFormulaCount = args.countFormulaSheetMembers(argsForImpact.targetSheetId)
  if (ownedFormulaCount === 0 || ownedFormulaCount !== args.state.formulas.size) {
    return undefined
  }
  const deferredMemberCount = args.tryDeferFormulaFamilyStructuralSourceTransforms(
    argsForImpact.targetSheetId,
    {
      ownerSheetName: argsForImpact.sheetName,
      targetSheetName: argsForImpact.sheetName,
      transform: argsForImpact.transform,
      preservesValue: true,
    },
    (representativeCellIndex) => {
      const representative = args.state.formulas.get(representativeCellIndex)
      return representative !== undefined && canDeferSimpleStructuralFormulaSource(args, representative, argsForImpact.transform)
    },
  )
  if (deferredMemberCount !== ownedFormulaCount) {
    return undefined
  }
  argsForImpact.markDeferredStructuralFormulaSources()
  return EMPTY_STRUCTURAL_FORMULA_IMPACTS
}

function tryCollectOwnedDirectAggregateRowInsertImpacts(
  args: CreateEngineStructureServiceArgs,
  argsForImpact: CollectStructuralFormulaImpactsOptions,
): StructuralFormulaImpacts | undefined {
  if (
    argsForImpact.targetSheetId === undefined ||
    argsForImpact.changedDefinedNames.size > 0 ||
    argsForImpact.changedTableNames.size > 0 ||
    argsForImpact.transform.kind !== 'insert' ||
    argsForImpact.transform.axis !== 'row'
  ) {
    return undefined
  }
  const ownedFormulaCount = args.countFormulaSheetMembers(argsForImpact.targetSheetId)
  if (ownedFormulaCount === 0 || ownedFormulaCount !== args.state.formulas.size) {
    return undefined
  }

  const directAggregateRetargetCellIndices: number[] = []
  let ownedFormulaVisitCount = 0
  let canUseDirectAggregateFastPath = true
  const targetSheetStructureVersion = args.state.workbook.getSheetById(argsForImpact.targetSheetId)?.structureVersion
  args.forEachFormulaCellOwnedBySheet(argsForImpact.sheetName, (cellIndex) => {
    if (!canUseDirectAggregateFastPath) {
      return
    }
    ownedFormulaVisitCount += 1
    const formula = args.state.formulas.get(cellIndex)
    const ownerRow =
      targetSheetStructureVersion === 1
        ? args.state.workbook.cellStore.rows[cellIndex]
        : args.state.workbook.getCellPosition(cellIndex)?.row
    const directAggregate = formula?.directAggregate
    if (
      ownerRow === undefined ||
      !formula ||
      !directAggregate ||
      directAggregate.sheetName !== argsForImpact.sheetName ||
      formula.compiled.symbolicSpills.length > 0
    ) {
      canUseDirectAggregateFastPath = false
      return
    }
    const ownerPositionAffected = structuralAxisIndexAffected(ownerRow, argsForImpact.transform)
    const directRangeAffected = runtimeDirectRangeAxisAffected(
      argsForImpact.targetSheetId,
      argsForImpact.sheetName,
      argsForImpact.transform,
      directAggregate,
    )
    if (ownerPositionAffected || directRangeAffected) {
      directAggregateRetargetCellIndices.push(cellIndex)
    }
  })

  if (!canUseDirectAggregateFastPath || ownedFormulaVisitCount !== ownedFormulaCount) {
    return undefined
  }
  return {
    formulaCellIndices: [],
    rebindCellIndices: [],
    preservedCellIndices: [],
    precomputedChangedInputCellIndices: [],
    ownerPositions: EMPTY_OWNER_POSITIONS,
    precomputedDirectAggregateValueCellIndices: [],
    directAggregateRetargetCellIndices,
  }
}

export function collectStructuralFormulaImpacts(
  args: CreateEngineStructureServiceArgs,
  argsForImpact: CollectStructuralFormulaImpactsOptions,
): StructuralFormulaImpacts {
  const ownedFamilyOnlyImpacts = tryCollectOwnedFamilyOnlyStructuralImpacts(args, argsForImpact)
  if (ownedFamilyOnlyImpacts) {
    return ownedFamilyOnlyImpacts
  }
  const ownedDirectAggregateRowInsertImpacts = tryCollectOwnedDirectAggregateRowInsertImpacts(args, argsForImpact)
  if (ownedDirectAggregateRowInsertImpacts) {
    return ownedDirectAggregateRowInsertImpacts
  }

  const formulaCellIndices = new Set<number>()
  const rebindCellIndices = new Set<number>()
  const preservedCellIndices = new Set<number>()
  const precomputedChangedInputCellIndices = new Set<number>()
  const candidateCellIndices = new Set<number>()
  const ownerPositions = new Map<number, { sheetName: string; row: number; col: number }>()
  const precomputedDirectAggregateValueCellIndices = new Set<number>()
  const directAggregateRetargetCellIndices = new Set<number>()
  let sharedOwnedPreservingSourceTransform: RuntimeFormula['structuralSourceTransform']
  let deferredOwnedFormulaFamilyMemberCount = 0
  const targetSheetStructureVersion =
    argsForImpact.targetSheetId === undefined ? undefined : args.state.workbook.getSheetById(argsForImpact.targetSheetId)?.structureVersion
  const sheetStructureVersionById = new Map<number, number | undefined>()
  const getSheetStructureVersion = (sheetId: number): number | undefined => {
    if (sheetId === argsForImpact.targetSheetId) {
      return targetSheetStructureVersion
    }
    if (sheetStructureVersionById.has(sheetId)) {
      return sheetStructureVersionById.get(sheetId)
    }
    const structureVersion = args.state.workbook.getSheetById(sheetId)?.structureVersion
    sheetStructureVersionById.set(sheetId, structureVersion)
    return structureVersion
  }
  const readCellPosition = (cellIndex: number): { sheetId: number; row: number; col: number } | undefined => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    if (sheetId === undefined || sheetId === 0) {
      return undefined
    }
    if (getSheetStructureVersion(sheetId) === 1) {
      const row = args.state.workbook.cellStore.rows[cellIndex]
      const col = args.state.workbook.cellStore.cols[cellIndex]
      return row === undefined || col === undefined ? undefined : { sheetId, row, col }
    }
    return args.state.workbook.getCellPosition(cellIndex)
  }
  const ownedPreservingSourceTransform = (): NonNullable<RuntimeFormula['structuralSourceTransform']> =>
    (sharedOwnedPreservingSourceTransform ??= {
      ownerSheetName: argsForImpact.sheetName,
      targetSheetName: argsForImpact.sheetName,
      transform: argsForImpact.transform,
      preservesValue: true,
    })
  const tryDeferOwnedFormulaFamilies = (): boolean => {
    if (
      argsForImpact.targetSheetId === undefined ||
      argsForImpact.changedDefinedNames.size > 0 ||
      argsForImpact.changedTableNames.size > 0 ||
      argsForImpact.transform.kind === 'delete' ||
      argsForImpact.transform.axis !== 'column' ||
      !args.canUseFormulaFamilyIndex()
    ) {
      return false
    }
    const ownedFormulaCount = args.countFormulaSheetMembers(argsForImpact.targetSheetId)
    if (ownedFormulaCount === 0) {
      return false
    }
    const deferredMemberCount = args.tryDeferFormulaFamilyStructuralSourceTransforms(
      argsForImpact.targetSheetId,
      ownedPreservingSourceTransform(),
      (representativeCellIndex) => {
        const representative = args.state.formulas.get(representativeCellIndex)
        return representative !== undefined && canDeferSimpleStructuralFormulaSource(args, representative, argsForImpact.transform)
      },
    )
    if (deferredMemberCount !== undefined) {
      argsForImpact.markDeferredStructuralFormulaSources()
      deferredOwnedFormulaFamilyMemberCount = deferredMemberCount
      return true
    }
    const familyIds: number[] = []
    let familyMemberCount = 0
    let canDeferFamilies = true
    args.forEachFormulaFamily((family) => {
      if (!canDeferFamilies || family.sheetId !== argsForImpact.targetSheetId) {
        return
      }
      const representativeCellIndex = family.runs.find((run) => run.cellIndices.length > 0)?.cellIndices[0]
      const representative = representativeCellIndex === undefined ? undefined : args.state.formulas.get(representativeCellIndex)
      if (!representative || !canDeferSimpleStructuralFormulaSource(args, representative, argsForImpact.transform)) {
        canDeferFamilies = false
        return
      }
      familyIds.push(family.id)
      family.runs.forEach((run) => {
        familyMemberCount += run.cellIndices.length
      })
    })
    if (!canDeferFamilies || familyIds.length === 0 || familyMemberCount !== ownedFormulaCount) {
      return false
    }
    const transform = ownedPreservingSourceTransform()
    familyIds.forEach((familyId) => {
      args.setFormulaFamilyStructuralSourceTransform(familyId, transform)
    })
    argsForImpact.markDeferredStructuralFormulaSources()
    deferredOwnedFormulaFamilyMemberCount = ownedFormulaCount
    return true
  }
  const canSkipOwnedDirectAggregateCandidate = (
    cellIndex: number,
    ownerPosition: { readonly row: number; readonly col: number } | undefined,
  ): boolean => {
    if (argsForImpact.changedDefinedNames.size > 0 || argsForImpact.changedTableNames.size > 0) {
      return false
    }
    if (argsForImpact.targetSheetId === undefined) {
      return false
    }
    const formula = args.state.formulas.get(cellIndex)
    if (!formula?.directAggregate) {
      return false
    }
    if (!ownerPosition) {
      return false
    }
    const ownerAxisIndex = argsForImpact.transform.axis === 'row' ? ownerPosition.row : ownerPosition.col
    if (structuralAxisIndexAffected(ownerAxisIndex, argsForImpact.transform)) {
      return false
    }
    return !runtimeDirectRangeAxisAffected(
      argsForImpact.targetSheetId,
      argsForImpact.sheetName,
      argsForImpact.transform,
      formula.directAggregate,
    )
  }
  const tryPrecomputeDeletedDirectAggregateValue = (
    cellIndex: number,
    formula: RuntimeFormula,
    ownerPosition: { row: number; col: number },
  ): boolean => {
    if (
      argsForImpact.changedDefinedNames.size > 0 ||
      argsForImpact.changedTableNames.size > 0 ||
      argsForImpact.transform.kind !== 'delete' ||
      argsForImpact.transform.axis !== 'row' ||
      argsForImpact.targetSheetId === undefined
    ) {
      return false
    }
    const directAggregate = formula.directAggregate
    if (!directAggregate || directAggregate.sheetName !== argsForImpact.sheetName || directAggregate.aggregateKind !== 'sum') {
      return false
    }
    if (mapStructuralAxisIndex(ownerPosition.row, argsForImpact.transform) === undefined) {
      return false
    }
    const overlapStart = Math.max(directAggregate.rowStart, argsForImpact.transform.start)
    const overlapEnd = Math.min(directAggregate.rowEnd, argsForImpact.transform.start + argsForImpact.transform.count - 1)
    if (overlapStart > overlapEnd) {
      return false
    }
    const aggregateSheet = args.state.workbook.getSheet(argsForImpact.sheetName)
    if (!aggregateSheet) {
      return false
    }
    const currentValue = args.state.workbook.cellStore.getValue(cellIndex, () => '')
    if (currentValue.tag !== ValueTag.Number) {
      return false
    }
    let deletedContribution = 0
    for (let row = overlapStart; row <= overlapEnd; row += 1) {
      const memberCellIndex =
        aggregateSheet.structureVersion === 1
          ? aggregateSheet.grid.getPhysical(row, directAggregate.col)
          : aggregateSheet.grid.get(row, directAggregate.col)
      if (memberCellIndex === -1) {
        continue
      }
      const memberValue = args.state.workbook.cellStore.getValue(memberCellIndex, () => '')
      switch (memberValue.tag) {
        case ValueTag.Number:
          deletedContribution += memberValue.value
          break
        case ValueTag.Boolean:
          deletedContribution += memberValue.value ? 1 : 0
          break
        case ValueTag.Empty:
        case ValueTag.String:
          break
        case ValueTag.Error:
          return false
      }
    }
    args.state.workbook.cellStore.setValue(cellIndex, {
      tag: ValueTag.Number,
      value: currentValue.value - deletedContribution,
    })
    precomputedChangedInputCellIndices.add(cellIndex)
    precomputedDirectAggregateValueCellIndices.add(cellIndex)
    return true
  }
  const canRetargetDirectAggregateWithoutFormulaRewrite = (
    formula: RuntimeFormula,
    ownerPosition: { row: number; col: number },
  ): boolean => {
    if (
      argsForImpact.changedDefinedNames.size > 0 ||
      argsForImpact.changedTableNames.size > 0 ||
      argsForImpact.targetSheetId === undefined ||
      argsForImpact.transform.axis !== 'row'
    ) {
      return false
    }
    const directAggregate = formula.directAggregate
    if (!directAggregate || directAggregate.sheetName !== argsForImpact.sheetName) {
      return false
    }
    if (formula.compiled.symbolicSpills.length > 0) {
      return false
    }
    if (mapStructuralAxisIndex(ownerPosition.row, argsForImpact.transform) === undefined) {
      return false
    }
    return mapStructuralAxisInterval(directAggregate.rowStart, directAggregate.rowEnd, argsForImpact.transform) !== undefined
  }
  const tryQueueDirectAggregateStructuralRetarget = (
    cellIndex: number,
    formula: RuntimeFormula,
    ownerPosition: { row: number; col: number },
  ): boolean => {
    if (!canRetargetDirectAggregateWithoutFormulaRewrite(formula, ownerPosition)) {
      return false
    }
    if (argsForImpact.transform.kind === 'delete' && !tryPrecomputeDeletedDirectAggregateValue(cellIndex, formula, ownerPosition)) {
      return false
    }
    directAggregateRetargetCellIndices.add(cellIndex)
    return true
  }
  const tryDeferOwnedSimpleFormula = (cellIndex: number): boolean => {
    if (argsForImpact.changedDefinedNames.size > 0 || argsForImpact.changedTableNames.size > 0) {
      return false
    }
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return false
    }
    if (canDeferSimpleStructuralFormulaSource(args, formula, argsForImpact.transform)) {
      formula.structuralSourceTransform = ownedPreservingSourceTransform()
      argsForImpact.markDeferredStructuralFormulaSources()
      preservedCellIndices.add(cellIndex)
      return true
    }
    const ownerPosition = readCellPosition(cellIndex)
    const ownerAxisIndex =
      ownerPosition === undefined ? undefined : argsForImpact.transform.axis === 'row' ? ownerPosition.row : ownerPosition.col
    if (ownerAxisIndex === undefined || mapStructuralAxisIndex(ownerAxisIndex, argsForImpact.transform) === undefined) {
      return false
    }
    const deleteClassification = classifySimpleDeleteStructuralFormulaSource(
      args,
      formula,
      argsForImpact.targetSheetId,
      argsForImpact.transform,
      targetSheetStructureVersion,
    )
    const dependsOnPrecomputedRefError = formula.dependencyIndices.some((dependencyCellIndex) =>
      precomputedChangedInputCellIndices.has(dependencyCellIndex),
    )
    const preservesBinding = deleteClassification === 'preserves-binding' && !dependsOnPrecomputedRefError
    const preservesValue = preservesBinding
    const becomesRefError = deleteClassification === 'ref-error' || dependsOnPrecomputedRefError
    if (!preservesBinding && !becomesRefError && !dependsOnPrecomputedRefError) {
      return false
    }
    formula.structuralSourceTransform = {
      ownerSheetName: argsForImpact.sheetName,
      targetSheetName: argsForImpact.sheetName,
      transform: argsForImpact.transform,
      preservesValue,
    }
    argsForImpact.markDeferredStructuralFormulaSources()
    if (preservesValue) {
      preservedCellIndices.add(cellIndex)
    } else if (becomesRefError || dependsOnPrecomputedRefError) {
      args.state.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Ref))
      precomputedChangedInputCellIndices.add(cellIndex)
    } else {
      formulaCellIndices.add(cellIndex)
    }
    return true
  }
  const deferredOwnedFormulaFamilies = tryDeferOwnedFormulaFamilies()
  if (!deferredOwnedFormulaFamilies) {
    args.forEachFormulaCellOwnedBySheet(argsForImpact.sheetName, (cellIndex) => {
      if (tryDeferOwnedSimpleFormula(cellIndex)) {
        return
      }
      const formula = args.state.formulas.get(cellIndex)
      const ownerPosition = readCellPosition(cellIndex)
      if (
        ownerPosition &&
        mapStructuralAxisIndex(argsForImpact.transform.axis === 'row' ? ownerPosition.row : ownerPosition.col, argsForImpact.transform) ===
          undefined
      ) {
        return
      }
      if (canSkipOwnedDirectAggregateCandidate(cellIndex, ownerPosition)) {
        return
      }
      if (formula && ownerPosition && tryQueueDirectAggregateStructuralRetarget(cellIndex, formula, ownerPosition)) {
        return
      }
      candidateCellIndices.add(cellIndex)
    })
  }
  const ownedFamilyDeferralCoversEveryFormula =
    deferredOwnedFormulaFamilies && deferredOwnedFormulaFamilyMemberCount === args.state.formulas.size
  if (!ownedFamilyDeferralCoversEveryFormula) {
    args.collectFormulaCellsReferencingSheet(argsForImpact.sheetName).forEach((cellIndex) => {
      if (directAggregateRetargetCellIndices.has(cellIndex)) {
        return
      }
      const formula = args.state.formulas.get(cellIndex)
      if (formula?.structuralSourceTransform !== undefined) {
        return
      }
      candidateCellIndices.add(cellIndex)
    })
  }
  if (argsForImpact.changedDefinedNames.size > 0) {
    args.collectFormulaCellsForDefinedNames([...argsForImpact.changedDefinedNames]).forEach((cellIndex) => {
      candidateCellIndices.add(cellIndex)
    })
  }
  if (argsForImpact.changedTableNames.size > 0) {
    args.collectFormulaCellsForTables([...argsForImpact.changedTableNames]).forEach((cellIndex) => {
      candidateCellIndices.add(cellIndex)
    })
  }
  if (args.state.counters && candidateCellIndices.size > 0) {
    addEngineCounter(args.state.counters, 'structuralFormulaImpactCandidates', candidateCellIndices.size)
  }
  candidateCellIndices.forEach((cellIndex) => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return
    }
    if (formula.structuralSourceTransform !== undefined) {
      return
    }
    if (!isCellIndexMapped(args, cellIndex)) {
      return
    }
    const ownerPosition = readCellPosition(cellIndex)
    if (!ownerPosition) {
      return
    }
    const ownerSheetName = args.state.workbook.getSheetNameById(ownerPosition.sheetId)
    if (!ownerSheetName) {
      return
    }
    if (tryQueueDirectAggregateStructuralRetarget(cellIndex, formula, ownerPosition)) {
      return
    }
    ownerPositions.set(cellIndex, { sheetName: ownerSheetName, row: ownerPosition.row, col: ownerPosition.col })
    const formulaValuePrecomputed = tryPrecomputeDeletedDirectAggregateValue(cellIndex, formula, ownerPosition)
    const axisIndex = argsForImpact.transform.axis === 'row' ? ownerPosition?.row : ownerPosition?.col
    const ownerPositionAffected =
      ownerSheetName === argsForImpact.sheetName &&
      axisIndex !== undefined &&
      structuralAxisIndexAffected(axisIndex, argsForImpact.transform)
    const touchesChangedName =
      argsForImpact.changedDefinedNames.size > 0 &&
      formula.compiled.symbolicNames.some((name) => argsForImpact.changedDefinedNames.has(normalizeDefinedName(name)))
    const touchesChangedTable =
      argsForImpact.changedTableNames.size > 0 && formula.compiled.symbolicTables.some((name) => argsForImpact.changedTableNames.has(name))
    if (!touchesChangedName && !touchesChangedTable && canDeferSimpleStructuralFormulaSource(args, formula, argsForImpact.transform)) {
      formula.structuralSourceTransform =
        ownerSheetName === argsForImpact.sheetName
          ? ownedPreservingSourceTransform()
          : {
              ownerSheetName,
              targetSheetName: argsForImpact.sheetName,
              transform: argsForImpact.transform,
              preservesValue: true,
            }
      argsForImpact.markDeferredStructuralFormulaSources()
      preservedCellIndices.add(cellIndex)
      return
    }
    const dependencyPositionAffected =
      !ownerPositionAffected &&
      argsForImpact.targetSheetId !== undefined &&
      (formula.dependencyIndices.some((dependencyCellIndex) => {
        if (args.state.workbook.cellStore.sheetIds[dependencyCellIndex] !== argsForImpact.targetSheetId) {
          return false
        }
        const dependencyAxisIndex = args.state.workbook.getCellAxisIndex(dependencyCellIndex, argsForImpact.transform.axis)
        return dependencyAxisIndex !== undefined && structuralAxisIndexAffected(dependencyAxisIndex, argsForImpact.transform)
      }) ||
        formula.rangeDependencies.some((rangeIndex) =>
          rangeDependencyAxisAffected(args.state.ranges.getDescriptor(rangeIndex), argsForImpact.targetSheetId!, argsForImpact.transform),
        ) ||
        runtimeDirectRangeAxisAffected(
          argsForImpact.targetSheetId,
          argsForImpact.sheetName,
          argsForImpact.transform,
          formula.directAggregate,
        ) ||
        runtimeDirectRangeAxisAffected(
          argsForImpact.targetSheetId,
          argsForImpact.sheetName,
          argsForImpact.transform,
          formula.directCriteria?.aggregateRange,
        ) ||
        formula.directCriteria?.criteriaPairs.some((pair) =>
          runtimeDirectRangeAxisAffected(argsForImpact.targetSheetId, argsForImpact.sheetName, argsForImpact.transform, pair.range),
        ) ||
        runtimeDirectRangeAxisAffected(
          argsForImpact.targetSheetId,
          argsForImpact.sheetName,
          argsForImpact.transform,
          formula.directLookup?.kind === 'exact' || formula.directLookup?.kind === 'approximate'
            ? {
                sheetName: formula.directLookup.prepared.sheetName,
                rowStart: formula.directLookup.prepared.rowStart,
                rowEnd: formula.directLookup.prepared.rowEnd,
                col: formula.directLookup.prepared.col,
              }
            : formula.directLookup?.kind === 'exact-uniform-numeric' || formula.directLookup?.kind === 'approximate-uniform-numeric'
              ? {
                  sheetName: formula.directLookup.sheetName,
                  rowStart: formula.directLookup.rowStart,
                  rowEnd: formula.directLookup.rowEnd,
                  col: formula.directLookup.col,
                }
              : undefined,
        ))
    const touchesSheetDependency =
      !ownerPositionAffected &&
      !dependencyPositionAffected &&
      formula.compiled.deps.some((dependency) => dependencyTouchesSheet(dependency, argsForImpact.sheetName))
    if (!ownerPositionAffected && !dependencyPositionAffected && !touchesSheetDependency && !touchesChangedName && !touchesChangedTable) {
      return
    }
    formulaCellIndices.add(cellIndex)
    if (ownerPositionAffected || dependencyPositionAffected || touchesSheetDependency || touchesChangedName || touchesChangedTable) {
      rebindCellIndices.add(cellIndex)
    }
    if (formulaValuePrecomputed) {
      rebindCellIndices.add(cellIndex)
    }
  })
  return {
    formulaCellIndices: [...formulaCellIndices],
    rebindCellIndices: [...rebindCellIndices],
    preservedCellIndices: [...preservedCellIndices],
    precomputedChangedInputCellIndices: [...precomputedChangedInputCellIndices],
    ownerPositions,
    precomputedDirectAggregateValueCellIndices: [...precomputedDirectAggregateValueCellIndices],
    directAggregateRetargetCellIndices: [...directAggregateRetargetCellIndices],
  }
}
