import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import {
  createDeferredInitialFormulaFamilyRunMap,
  flushDeferredInitialFormulaFamilyRuns,
  noteDeferredFormulaFamilyRunMember,
  type DeferredInitialFormulaFamilyRun,
} from './formula-initialization-family-runs.js'
import {
  flushAlignedFreshFormulaFamilyRuns,
  readAlignedFreshFormulaFamilyRunsFromRefs,
} from './formula-initialization-restored-family-runs.js'
import { createInitialFormulaValueWriter } from './formula-initialization-value-writer.js'
import {
  materializeDeferredFormulaInstances,
  readAlignedFreshFormulaInstancesFromRefs,
  writeDeferredFormulaInstance,
} from './formula-initialization-fresh-instances.js'
import { initialFormulaEntryRefAt, type InitialFormulaEntryRefSource } from './formula-initialization-refs.js'
import { createInitialFormulaCellIndexPlan } from './formula-initialization-cell-index-plan.js'
import { hasPendingFormulaDependency } from './formula-initialization-predicates.js'
import { tryBindHydratedFreshDirectFormula } from './formula-initialization-hydrated-direct-scalar.js'
import type {
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'

const MAX_EAGER_FRESH_FORMULA_INSTANCE_RECORDS = 16_384

export interface HydratedPreparedFormulaInitializationUncheckedArgs {
  readonly serviceArgs: EngineFormulaInitializationServiceArgs
  readonly refs: InitialFormulaEntryRefSource<HydratedPreparedFormulaInitializationRef>
  readonly potentialNewCells: number | undefined
  readonly hasCycleMembersNow: () => boolean
  readonly resolveSheetName: (sheetId: number) => string
  readonly registerDeferredFormulaFamilyRun: (run: DeferredInitialFormulaFamilyRun) => void
}

export function initializeHydratedPreparedCellFormulasAtNowUnchecked({
  serviceArgs: args,
  refs,
  potentialNewCells,
  hasCycleMembersNow,
  resolveSheetName,
  registerDeferredFormulaFamilyRun,
}: HydratedPreparedFormulaInitializationUncheckedArgs): void {
  if (refs.length === 0) {
    return
  }

  args.beginMutationCollection()
  args.checkEvaluationBudget()
  let hadCycleMembersBefore: boolean | undefined
  const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
  let topologyChanged = false
  let compileMs = 0
  let workbookDateSystem: string | undefined
  const resolveWorkbookDateSystem = (): string | undefined =>
    (workbookDateSystem ??= args.state.workbook.getCalculationSettings().dateSystem)
  const reservedNewCells = potentialNewCells ?? refs.length
  const hadExistingFormulas = args.state.formulas.size > 0
  args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
  args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
  args.resetMaterializedCellScratch(reservedNewCells)
  const { targetCellIndices, pendingInitialFormulaCellIndices, pendingFormulaCells } = createInitialFormulaCellIndexPlan({
    refs,
    hadExistingFormulas,
    resolveCellIndex: (ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col),
    checkEvaluationBudget: args.checkEvaluationBudget,
  })
  let canAssignTopoInBatch = !hadExistingFormulas
  let needsFreshTopoRebuild = false
  let nextTopoRank = 0
  const canEagerHydrateFreshFormulaInstances =
    !hadExistingFormulas && args.hydrateFreshFormulaInstances !== undefined && refs.length <= MAX_EAGER_FRESH_FORMULA_INSTANCE_RECORDS
  const shouldDeferFormulaInstanceTable =
    !hadExistingFormulas && (canEagerHydrateFreshFormulaInstances || args.deferFormulaInstanceTableRebuild !== undefined)
  const alignedFreshFormulaInstances = canEagerHydrateFreshFormulaInstances ? readAlignedFreshFormulaInstancesFromRefs(refs) : undefined
  const alignedFreshFormulaFamilyRuns = readAlignedFreshFormulaFamilyRunsFromRefs({
    refs,
    hadExistingFormulas,
    counters: args.state.counters,
  })
  const shouldDeferFormulaFamilyIndex =
    args.deferFormulaFamilyIndexRebuild !== undefined && (!hadExistingFormulas || alignedFreshFormulaFamilyRuns !== undefined)
  const deferredFormulaInstances =
    canEagerHydrateFreshFormulaInstances && alignedFreshFormulaInstances === undefined
      ? Array<FormulaInstanceSnapshot>(refs.length)
      : undefined
  let deferredFormulaInstanceCount = 0
  const canCaptureDeferredFormulaFamilyRuns = !shouldDeferFormulaFamilyIndex || args.deferFormulaFamilyIndexRuns !== undefined
  const deferredFormulaFamilyRuns =
    hadExistingFormulas || alignedFreshFormulaFamilyRuns !== undefined || !canCaptureDeferredFormulaFamilyRuns
      ? undefined
      : createDeferredInitialFormulaFamilyRunMap()

  args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
  try {
    args.clearTemplateFormulaCache()
    const compileStarted = performance.now()
    const valueWriter = createInitialFormulaValueWriter(args)
    const bindFormulaEntries = (): void => {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
          args.checkEvaluationBudget()
          const ref = initialFormulaEntryRefAt(refs, refIndex)
          const cellIndex = hadExistingFormulas ? pendingInitialFormulaCellIndices[refIndex]! : targetCellIndices[refIndex]!
          const ownerSheetName = resolveSheetName(ref.sheetId)
          const usedHydratedDirectFastBinding = tryBindHydratedFreshDirectFormula(args, hadExistingFormulas, cellIndex, ownerSheetName, ref)
          if (usedHydratedDirectFastBinding) {
            topologyChanged = true
          } else {
            topologyChanged =
              args.bindPreparedFormula(cellIndex, ownerSheetName, ref.source, ref.compiled, ref.templateId, {
                deferFamilyRegistration:
                  shouldDeferFormulaFamilyIndex || deferredFormulaFamilyRuns !== undefined || alignedFreshFormulaFamilyRuns !== undefined,
                preserveCachedValueOnFullRecalc: ref.preserveCachedValueOnFullRecalc === true,
                deferFormulaInstanceRegistration: shouldDeferFormulaInstanceTable,
                assumeFreshFormula: !hadExistingFormulas,
                ownerPosition: {
                  sheetName: ownerSheetName,
                  row: ref.row,
                  col: ref.col,
                },
                resolveWorkbookDateSystem,
              }) || topologyChanged
          }
          const runtimeFormula = args.state.formulas.get(cellIndex)
          deferredFormulaInstanceCount = writeDeferredFormulaInstance(
            deferredFormulaInstances,
            deferredFormulaInstanceCount,
            { cellIndex, row: ref.row, col: ref.col, ownerSheetName },
            runtimeFormula,
          )
          if (alignedFreshFormulaFamilyRuns === undefined) {
            noteDeferredFormulaFamilyRunMember({
              runs: deferredFormulaFamilyRuns,
              prepared: {
                cellIndex,
                sheetId: ref.sheetId,
                row: ref.row,
                col: ref.col,
                ...(ref.templateId !== undefined ? { templateId: ref.templateId } : {}),
              },
              runtimeFormula,
            })
          }
          valueWriter.writeValueAt(cellIndex, ref.sheetId, ref.col, ref.value)
          if (canAssignTopoInBatch && pendingFormulaCells) {
            const hasPendingDependency =
              runtimeFormula !== undefined &&
              hasPendingFormulaDependency(runtimeFormula, pendingFormulaCells, (rangeIndex) => args.state.ranges.getMembersView(rangeIndex))
            if (!runtimeFormula || hasPendingDependency) {
              needsFreshTopoRebuild ||= hasPendingDependency
              canAssignTopoInBatch = false
            } else {
              args.state.workbook.cellStore.topoRanks[cellIndex] = nextTopoRank
              nextTopoRank += 1
            }
          }
          if (pendingFormulaCells) {
            pendingFormulaCells.delete(cellIndex)
          }
        }
        if (
          !flushAlignedFreshFormulaFamilyRuns({
            runs: alignedFreshFormulaFamilyRuns,
            shouldDeferFormulaFamilyIndex,
            deferFormulaFamilyIndexRuns: args.deferFormulaFamilyIndexRuns,
            deferFormulaFamilyIndexRebuild: args.deferFormulaFamilyIndexRebuild,
            registerFormulaFamilyRun: registerDeferredFormulaFamilyRun,
            checkEvaluationBudget: args.checkEvaluationBudget,
          })
        ) {
          flushDeferredInitialFormulaFamilyRuns({
            runs: deferredFormulaFamilyRuns,
            shouldDeferFormulaFamilyIndex,
            deferFormulaFamilyIndexRuns: args.deferFormulaFamilyIndexRuns,
            deferFormulaFamilyIndexRebuild: args.deferFormulaFamilyIndexRebuild,
            registerFormulaFamilyRun: registerDeferredFormulaFamilyRun,
            checkEvaluationBudget: args.checkEvaluationBudget,
          })
        }
        if (shouldDeferFormulaInstanceTable) {
          if (alignedFreshFormulaInstances !== undefined) {
            args.hydrateFreshFormulaInstances?.(alignedFreshFormulaInstances)
          } else if (deferredFormulaInstances) {
            args.hydrateFreshFormulaInstances?.(materializeDeferredFormulaInstances(deferredFormulaInstances, deferredFormulaInstanceCount))
          } else {
            args.deferFormulaInstanceTableRebuild?.()
          }
        }
        args.checkEvaluationBudget()
        valueWriter.flush()
      })
    }
    args.checkEvaluationBudget()
    args.withInitialFormulaCells(pendingInitialFormulaCellIndices, bindFormulaEntries)
    compileMs += performance.now() - compileStarted
  } finally {
    args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
  }

  if ((topologyChanged || needsFreshTopoRebuild) && !(canAssignTopoInBatch && !hadExistingFormulas)) {
    args.checkEvaluationBudget()
    const repaired =
      !needsFreshTopoRebuild &&
      !hadCycleMembersBeforeNow() &&
      refs.length > 0 &&
      args.repairTopoRanks(targetCellIndices.length > 0 ? targetCellIndices : pendingInitialFormulaCellIndices)
    if (!repaired) {
      args.checkEvaluationBudget()
      args.rebuildTopoRanks()
      args.checkEvaluationBudget()
      args.detectCycles()
    }
  }
  args.checkEvaluationBudget()
  const lastMetrics = args.state.getLastMetrics()
  args.state.setLastMetrics({
    ...lastMetrics,
    batchId: lastMetrics.batchId + 1,
    changedInputCount: 0,
    compileMs,
    recalcMs: 0,
  })
}
