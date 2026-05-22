import { Effect } from 'effect'
import type { StructuralAxisTransform } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import { composeFormulaFamilyStructuralSourceTransform } from '../../formula/formula-family-store.js'
import { mapStructuralAxisIndex, structuralTransformForOp } from '../../engine-structural-utils.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import { EngineStructureError } from '../errors.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { captureAxisRangeCellState, captureSheetCellState, shouldCaptureStoredCell } from './structure-cell-state.js'
import {
  dependencyTouchesSheet,
  structuralDirectAggregateRewritePreservesValue,
  structuralRewritePreservesBinding,
  structuralRewritePreservesValue,
} from './structure-formula-rewrite-guards.js'
import {
  rewriteFormulaFromTemplate,
  rewriteFormulaSourceFallback,
  rewriteStructuralFormulaCompiled,
  structuralRewritePreservesDirectCellDependencies,
  type StructuralFormulaRewriteCache,
} from './structure-formula-rewrite.js'
import { rewriteDefinedNamesForStructuralTransform, rewriteWorkbookMetadataForStructuralTransform } from './structure-metadata-rewrite.js'
import { rewriteFormulaSourceForDeletedStructuredReferences } from './structure-structured-ref-rewrite.js'
import {
  clearSpillArtifactsForSheet,
  clearPivotOutputsForSheet,
  clearRemovedCellRuntimeState,
  collectStructuralRangeDependencies,
  isCellIndexMapped,
} from './structure-runtime-cleanup.js'
import { canDeferSimpleStructuralFormulaSource } from './structure-formula-source-deferral.js'
import { materializeDeferredStructuralFormulaSources as materializeDeferredStructuralFormulaSourcesNow } from './structure-deferred-formula-sources.js'
import { collectStructuralFormulaImpacts } from './structure-formula-impacts.js'
import type {
  CreateEngineStructureServiceArgs,
  EngineStructureService,
  StructuralAxisOpResult,
  StructuralFormulaRebindInput,
} from './structure-service-types.js'

export type {
  CreateEngineStructureServiceArgs,
  EngineStructureService,
  EngineStructureState,
  StructuralAxisOp,
  StructuralAxisOpResult,
  StructuralFormulaRebindInput,
} from './structure-service-types.js'

const EMPTY_STRING_SET = new Set<string>()
const EMPTY_DELETED_TABLE_COLUMNS: [] = []

export function createEngineStructureService(args: CreateEngineStructureServiceArgs): EngineStructureService {
  let hasDeferredStructuralFormulaSources = false

  const resolveStructuralFormulaRebindInputs = (argsForResolve: {
    readonly formulaCellIndices: readonly number[]
    readonly sheetName: string
    readonly transform: StructuralAxisTransform
    readonly transaction: StructuralTransaction
    readonly changedDefinedNames: ReadonlySet<string>
    readonly changedTableNames: ReadonlySet<string>
    readonly deletedTableColumns: readonly {
      readonly tableName: string
      readonly columnName: string
    }[]
    readonly ownerPositions: ReadonlyMap<number, { sheetName: string; row: number; col: number }>
    readonly precomputedDirectAggregateValueCellIndices: readonly number[]
  }) => {
    const inputs: StructuralFormulaRebindInput[] = []
    const preservedCellIndices: number[] = []
    const templateRewriteCache: StructuralFormulaRewriteCache = new Map()
    const remappedCellsByIndex = new Map(argsForResolve.transaction.remappedCells.map((entry) => [entry.cellIndex, entry] as const))
    const precomputedDirectAggregateValueCellIndices = new Set(argsForResolve.precomputedDirectAggregateValueCellIndices)
    argsForResolve.formulaCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      const previousOwnerPosition = argsForResolve.ownerPositions.get(cellIndex)
      if (!previousOwnerPosition) {
        return
      }
      const ownerSheetName = previousOwnerPosition.sheetName
      const touchesChangedName = formula.compiled.symbolicNames.some((name) =>
        argsForResolve.changedDefinedNames.has(normalizeDefinedName(name)),
      )
      const touchesChangedTable = formula.compiled.symbolicTables.some((name) => argsForResolve.changedTableNames.has(name))
      const touchesTargetSheetDependency = formula.compiled.deps.some((dependency) =>
        dependencyTouchesSheet(dependency, argsForResolve.sheetName),
      )
      const shouldBypassTemplateStructuralRewrite = ownerSheetName !== argsForResolve.sheetName && touchesTargetSheetDependency
      const representative = remappedCellsByIndex.get(cellIndex)
      const previousOwnerRow = representative?.fromRow ?? previousOwnerPosition.row
      const previousOwnerCol = representative?.fromCol ?? previousOwnerPosition.col
      const ownerRow =
        representative?.toRow ??
        (ownerSheetName === argsForResolve.sheetName && argsForResolve.transform.axis === 'row'
          ? mapStructuralAxisIndex(previousOwnerRow, argsForResolve.transform)
          : previousOwnerRow)
      const ownerCol =
        representative?.toCol ??
        (ownerSheetName === argsForResolve.sheetName && argsForResolve.transform.axis === 'column'
          ? mapStructuralAxisIndex(previousOwnerCol, argsForResolve.transform)
          : previousOwnerCol)
      if (ownerRow === undefined || ownerCol === undefined) {
        return
      }
      if (!touchesChangedName && !touchesChangedTable && canDeferSimpleStructuralFormulaSource(args, formula, argsForResolve.transform)) {
        formula.structuralSourceTransform = {
          ownerSheetName,
          targetSheetName: argsForResolve.sheetName,
          transform: argsForResolve.transform,
          preservesValue: true,
        }
        hasDeferredStructuralFormulaSources = true
        preservedCellIndices.push(cellIndex)
        return
      }
      const templateRewrite =
        !touchesChangedName &&
        !touchesChangedTable &&
        !shouldBypassTemplateStructuralRewrite &&
        formula.templateId !== undefined &&
        previousOwnerRow !== undefined &&
        previousOwnerCol !== undefined
          ? rewriteFormulaFromTemplate(
              templateRewriteCache,
              formula,
              {
                templateId: formula.templateId,
                ownerSheetName,
                targetSheetName: argsForResolve.sheetName,
                representativeRow: previousOwnerRow,
                representativeCol: previousOwnerCol,
                ownerRow,
                ownerCol,
              },
              argsForResolve.sheetName,
              argsForResolve.transform,
            )
          : undefined
      const compiledRewrite =
        templateRewrite === undefined
          ? rewriteStructuralFormulaCompiled(formula, ownerSheetName, argsForResolve.sheetName, argsForResolve.transform)
          : undefined
      const rewritten = !touchesChangedName && !touchesChangedTable ? (compiledRewrite ?? templateRewrite) : compiledRewrite
      const changedMetadataFormulaSource = (): string => {
        const structuralSource = rewriteFormulaSourceFallback(
          formula.source,
          ownerSheetName,
          argsForResolve.sheetName,
          argsForResolve.transform,
        )
        if (!touchesChangedTable) {
          return structuralSource
        }
        return rewriteFormulaSourceForDeletedStructuredReferences(structuralSource, argsForResolve.deletedTableColumns) ?? structuralSource
      }
      if (!rewritten) {
        if (!touchesChangedName && !touchesChangedTable && formula.directAggregate !== undefined) {
          return
        }
        const canReuseCompiled =
          formula.compiled.symbolicNames.length === 0 &&
          formula.compiled.symbolicTables.length === 0 &&
          formula.compiled.symbolicSpills.length === 0
        inputs.push(
          canReuseCompiled
            ? {
                cellIndex,
                ownerSheetName,
                ownerRow,
                ownerCol,
                source: formula.source,
                compiled: formula.compiled,
                ...(formula.templateId === undefined ? {} : { templateId: formula.templateId }),
              }
            : {
                cellIndex,
                ownerSheetName,
                ownerRow,
                ownerCol,
                source: touchesChangedName || touchesChangedTable ? changedMetadataFormulaSource() : formula.source,
              },
        )
        return
      }
      if (touchesChangedName || touchesChangedTable) {
        inputs.push({
          cellIndex,
          ownerSheetName,
          ownerRow,
          ownerCol,
          source: changedMetadataFormulaSource(),
        })
        return
      }
      const preservesDirectCellDependencies = structuralRewritePreservesDirectCellDependencies(args, formula, rewritten, ownerSheetName)
      const preservesBinding =
        structuralRewritePreservesBinding(
          formula,
          rewritten,
          formula.rangeDependencies.every((rangeIndex) => args.state.ranges.getFormulaMembersView(rangeIndex).length === 0),
        ) || preservesDirectCellDependencies
      const preservesValue =
        precomputedDirectAggregateValueCellIndices.has(cellIndex) ||
        structuralRewritePreservesValue(formula, rewritten, argsForResolve.transform) ||
        structuralDirectAggregateRewritePreservesValue(formula, rewritten, argsForResolve.transform)
      const hasOnlyPlaceholderDirectDependencies =
        formula.dependencyIndices.length > 0 &&
        !formula.dependencyIndices.every((dependencyCellIndex) => shouldCaptureStoredCell(args, dependencyCellIndex))
      const rewrittenDirectDependenciesChanged =
        formula.compiled.deps.length !== rewritten.compiled.deps.length ||
        formula.compiled.deps.some((dependency, index) => dependency !== rewritten.compiled.deps[index])
      const rewrittenPlaceholderDependencyNeedsRebind =
        preservesBinding && rewrittenDirectDependenciesChanged && hasOnlyPlaceholderDirectDependencies
      inputs.push({
        cellIndex,
        ownerSheetName,
        ownerRow,
        ownerCol,
        source: rewritten.source,
        compiled: rewritten.compiled,
        ...(formula.templateId === undefined || rewritten.source !== formula.source ? {} : { templateId: formula.templateId }),
        preservesBinding: preservesBinding && !rewrittenPlaceholderDependencyNeedsRebind,
        preservesValue,
      })
    })
    return { inputs, preservedCellIndices }
  }

  const materializeDeferredStructuralFormulaSources = (): void => {
    hasDeferredStructuralFormulaSources = materializeDeferredStructuralFormulaSourcesNow(args, hasDeferredStructuralFormulaSources)
  }
  const hasFormulaLocalStructuralSourceTransforms = (): boolean => {
    let found = false
    args.state.formulas.forEach((formula) => {
      if (formula.structuralSourceTransform !== undefined) {
        found = true
      }
    })
    return found
  }
  const canPreserveDeferredStructuralFormulaSourcesForTransform = (sheetName: string, transform: StructuralAxisTransform): boolean => {
    if (
      !hasDeferredStructuralFormulaSources ||
      transform.kind !== 'insert' ||
      transform.axis !== 'column' ||
      hasFormulaLocalStructuralSourceTransforms()
    ) {
      return false
    }
    const nextTransform = {
      ownerSheetName: sheetName,
      targetSheetName: sheetName,
      transform,
      preservesValue: true,
    }
    const pendingFamilyTransforms = args.peekFormulaFamilyStructuralSourceTransforms()
    return (
      pendingFamilyTransforms.length > 0 &&
      pendingFamilyTransforms.every((entry) => composeFormulaFamilyStructuralSourceTransform(entry.transform, nextTransform) !== undefined)
    )
  }

  const service: EngineStructureService = {
    captureSheetCellState(sheetName) {
      return Effect.try({
        try: () => captureSheetCellState(args, sheetName),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture sheet cell state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureRowRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(args, sheetName, 'row', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture row state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureColumnRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(args, sheetName, 'column', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture column state for ${sheetName}`,
            cause,
          }),
      })
    },
    materializeDeferredStructuralFormulaSourcesNow: materializeDeferredStructuralFormulaSources,
    materializeDeferredStructuralFormulaSources() {
      return Effect.try({
        try: () => service.materializeDeferredStructuralFormulaSourcesNow(),
        catch: (cause) =>
          new EngineStructureError({
            message: 'Failed to materialize deferred structural formula sources',
            cause,
          }),
      })
    },
    applyStructuralAxisOpNow(op): StructuralAxisOpResult {
      const transform = structuralTransformForOp(op)
      const sheetName = op.sheetName
      if (!canPreserveDeferredStructuralFormulaSourcesForTransform(sheetName, transform)) {
        materializeDeferredStructuralFormulaSources()
      }
      const targetSheetId = args.state.workbook.getSheet(sheetName)?.id
      const hasStructuralMetadata = args.state.workbook.hasStructuralMetadataForSheet(sheetName)

      const hasPivots = hasStructuralMetadata && args.state.workbook.hasPivots()
      if (hasPivots) {
        clearPivotOutputsForSheet(args, sheetName)
      }
      const { changedTableNames, tableHeaderCellWrites, deletedTableColumns } = hasStructuralMetadata
        ? rewriteWorkbookMetadataForStructuralTransform(args, sheetName, transform)
        : { changedTableNames: EMPTY_STRING_SET, tableHeaderCellWrites: [], deletedTableColumns: EMPTY_DELETED_TABLE_COLUMNS }
      const changedDefinedNames = hasStructuralMetadata
        ? rewriteDefinedNamesForStructuralTransform(args, sheetName, transform, deletedTableColumns, changedTableNames)
        : EMPTY_STRING_SET
      const impactedFormulas = collectStructuralFormulaImpacts(args, {
        targetSheetId,
        transform,
        sheetName,
        changedDefinedNames,
        changedTableNames,
        markDeferredStructuralFormulaSources: () => {
          hasDeferredStructuralFormulaSources = true
        },
      })

      const transaction =
        args.state.workbook.planStructuralAxisTransform(sheetName, transform) ??
        (() => {
          throw new Error(`Missing sheet for structural op: ${sheetName}`)
        })()
      const hadSheetSpillMetadata = hasStructuralMetadata && args.state.workbook.listSpills().some((spill) => spill.sheetName === sheetName)
      const preStructuralSpillArtifacts = hadSheetSpillMetadata
        ? clearSpillArtifactsForSheet(args, sheetName)
        : { changedCellIndices: [], ownerCellIndices: [] }

      switch (op.kind) {
        case 'insertRows':
          args.state.workbook.insertRows(sheetName, op.start, op.count, op.entries)
          break
        case 'deleteRows':
          args.state.workbook.deleteRows(sheetName, op.start, op.count)
          break
        case 'moveRows':
          args.state.workbook.moveRows(sheetName, op.start, op.count, op.target)
          break
        case 'insertColumns':
          args.state.workbook.insertColumns(sheetName, op.start, op.count, op.entries)
          break
        case 'deleteColumns':
          args.state.workbook.deleteColumns(sheetName, op.start, op.count)
          break
        case 'moveColumns':
          args.state.workbook.moveColumns(sheetName, op.start, op.count, op.target)
          break
      }

      args.state.workbook.applyPlannedStructuralTransaction(transaction)

      const tableHeaderCellChangedIndices = tableHeaderCellWrites.flatMap((write) => {
        const cellIndex = args.writeTableHeaderCell(write.sheetName, write.row, write.col, write.value)
        return cellIndex === undefined ? [] : [cellIndex]
      })

      const hasNoFormulaStructuralWork =
        impactedFormulas.formulaCellIndices.length === 0 &&
        impactedFormulas.rebindCellIndices.length === 0 &&
        impactedFormulas.precomputedChangedInputCellIndices.length === 0 &&
        impactedFormulas.precomputedDirectAggregateValueCellIndices.length === 0 &&
        impactedFormulas.directAggregateRetargetCellIndices.length === 0
      if (
        hasNoFormulaStructuralWork &&
        transaction.removedCellIndices.length === 0 &&
        changedDefinedNames.size === 0 &&
        changedTableNames.size === 0 &&
        tableHeaderCellChangedIndices.length === 0 &&
        !hasPivots &&
        !hadSheetSpillMetadata
      ) {
        return {
          transaction,
          changedCellIndices: [],
          precomputedChangedInputCellIndices: [],
          formulaCellIndices: [],
          topologyChanged: false,
          graphRefreshRequired: false,
        }
      }

      const structuralRangeDependencies = collectStructuralRangeDependencies(args, impactedFormulas.formulaCellIndices)

      let hadCycleFormulas: boolean | undefined
      const hasCycleFormulas = (): boolean => {
        if (hadCycleFormulas !== undefined) {
          return hadCycleFormulas
        }
        if (args.state.counters) {
          addEngineCounter(args.state.counters, 'cycleFormulaScans')
        }
        let found = false
        args.state.formulas.forEach((_formula, cellIndex) => {
          if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
            found = true
          }
        })
        hadCycleFormulas = found
        return found
      }
      const removedFormulaCellIndices = transaction.removedCellIndices.filter((cellIndex) => args.state.formulas.has(cellIndex))
      const removedFormulaCellIndexSet = new Set<number>(removedFormulaCellIndices)
      const removedCycleFormulaCount = removedFormulaCellIndices.filter(
        (cellIndex) => ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0,
      ).length
      transaction.removedCellIndices.forEach((cellIndex) => {
        clearRemovedCellRuntimeState(args, cellIndex)
      })

      args.retargetRangeDependencies(transaction, structuralRangeDependencies)
      const directRetargetedFormulaCellIndices: number[] = []
      const directRetargetedPreservedFormulaCellIndices: number[] = []
      const precomputedDirectAggregateValueCellIndices = new Set(impactedFormulas.precomputedDirectAggregateValueCellIndices)
      const directAggregateRetargetInputs = impactedFormulas.directAggregateRetargetCellIndices
        .filter((cellIndex) => isCellIndexMapped(args, cellIndex))
        .map((cellIndex) => ({
          cellIndex,
          ownerSheetName: sheetName,
          preservesValue: true,
        }))
      const directAggregateRetargetedCellIndices = args.retargetDirectAggregateFormulasForStructuralTransform(
        directAggregateRetargetInputs,
        sheetName,
        transform,
      )
      if (directAggregateRetargetedCellIndices.length > 0) {
        directRetargetedFormulaCellIndices.push(...directAggregateRetargetedCellIndices)
        directRetargetedPreservedFormulaCellIndices.push(...directAggregateRetargetedCellIndices)
        hasDeferredStructuralFormulaSources = true
      }
      const rebindResolution = resolveStructuralFormulaRebindInputs({
        formulaCellIndices: impactedFormulas.rebindCellIndices.filter((cellIndex) => isCellIndexMapped(args, cellIndex)),
        sheetName,
        transform,
        transaction,
        changedDefinedNames,
        changedTableNames,
        deletedTableColumns,
        ownerPositions: impactedFormulas.ownerPositions,
        precomputedDirectAggregateValueCellIndices: [...precomputedDirectAggregateValueCellIndices],
      })
      const rebindInputs = rebindResolution.inputs
      const remainingRebindInputs: StructuralFormulaRebindInput[] = []
      rebindInputs.forEach((input) => {
        const formula = args.state.formulas.get(input.cellIndex)
        const directAggregateRetargeted =
          input.preservesBinding === true &&
          formula?.directAggregate !== undefined &&
          args.retargetDirectAggregateFormulaForStructuralTransform(input, sheetName, transform)
        if (directAggregateRetargeted) {
          hasDeferredStructuralFormulaSources = true
        }
        if (
          directAggregateRetargeted ||
          (input.preservesBinding === true &&
            formula?.directAggregate !== undefined &&
            input.compiled !== undefined &&
            args.rewriteFormulaCompiledPreservingBinding(input))
        ) {
          directRetargetedFormulaCellIndices.push(input.cellIndex)
          if (input.preservesValue) {
            directRetargetedPreservedFormulaCellIndices.push(input.cellIndex)
          }
          return
        }
        remainingRebindInputs.push(input)
      })
      if (args.state.counters && remainingRebindInputs.length > 0) {
        addEngineCounter(args.state.counters, 'structuralFormulaRebindInputs', remainingRebindInputs.length)
      }
      const formulaCellIndices = impactedFormulas.formulaCellIndices.filter((cellIndex) => isCellIndexMapped(args, cellIndex))
      const structuralSpillFormulaCellIndices = preStructuralSpillArtifacts.ownerCellIndices.filter(
        (cellIndex) => isCellIndexMapped(args, cellIndex) && args.state.formulas.has(cellIndex),
      )
      const onlyDirectAggregateFormulaCells =
        formulaCellIndices.length > 0 &&
        formulaCellIndices.every((cellIndex) => args.state.formulas.get(cellIndex)?.directAggregate !== undefined)
      args.rebindFormulaCells(remainingRebindInputs)
      const reboundFormulaCellIndices = new Set([
        ...directRetargetedFormulaCellIndices,
        ...remainingRebindInputs.map((input) => input.cellIndex),
      ])
      const preservedFormulaCellIndices = new Set([
        ...impactedFormulas.preservedCellIndices,
        ...rebindResolution.preservedCellIndices,
        ...directRetargetedPreservedFormulaCellIndices,
        ...remainingRebindInputs.filter((input) => input.preservesValue).map((input) => input.cellIndex),
      ])
      const lostSurvivingFormulaCells = impactedFormulas.formulaCellIndices.some(
        (cellIndex) =>
          !reboundFormulaCellIndices.has(cellIndex) && !isCellIndexMapped(args, cellIndex) && !removedFormulaCellIndexSet.has(cellIndex),
      )
      const hasNonPreservedRebind = remainingRebindInputs.some((input) => input.preservesBinding !== true)
      const needsDeleteAcyclicRebindCheck =
        transform.kind === 'delete' &&
        changedDefinedNames.size === 0 &&
        changedTableNames.size === 0 &&
        (hasNonPreservedRebind || lostSurvivingFormulaCells)
      const deleteOnlyAcyclicRebind = needsDeleteAcyclicRebindCheck && !hasCycleFormulas()
      const topologyChanged = removedFormulaCellIndices.length > 0 || hasNonPreservedRebind || lostSurvivingFormulaCells
      const graphRefreshRequired =
        ((hasNonPreservedRebind || lostSurvivingFormulaCells) && !onlyDirectAggregateFormulaCells && !deleteOnlyAcyclicRebind) ||
        removedCycleFormulaCount > 0
      const recalculatedFormulaCellIndices = [
        ...formulaCellIndices.filter((cellIndex) => !preservedFormulaCellIndices.has(cellIndex)),
        ...structuralSpillFormulaCellIndices,
      ]
      return {
        transaction,
        changedCellIndices: [
          ...transaction.removedCellIndices,
          ...preStructuralSpillArtifacts.changedCellIndices,
          ...tableHeaderCellChangedIndices,
        ],
        precomputedChangedInputCellIndices: impactedFormulas.precomputedChangedInputCellIndices,
        formulaCellIndices: [...new Set(recalculatedFormulaCellIndices)],
        topologyChanged,
        graphRefreshRequired,
      }
    },
    applyStructuralAxisOp(op) {
      return Effect.try({
        try: () => service.applyStructuralAxisOpNow(op),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to apply structural operation ${op.kind}`,
            cause,
          }),
      })
    },
  }

  return service
}
