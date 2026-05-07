import {
  type CompiledFormula,
  renameCompiledFormulaSheetReferences,
  renameCompiledFormulaSheetReferenceMetadata,
  renameCompiledFormulaSheetReferenceMetadataInPlace,
} from '@bilig/formula'
import type { CompiledPlanRecord, RuntimeDirectScalarDescriptor } from '../runtime-state.js'
import {
  canRewriteCompiledPreservingBindings,
  canRewriteCompiledPreservingDirectAggregate,
  canRewriteCompiledPreservingDirectScalar,
} from './formula-binding-shape-helpers.js'
import { buildDirectScalarDescriptor } from './formula-binding-direct-scalar.js'
import { renameDirectAggregateDescriptorSheet, type ParsedCompiledFormula } from './formula-binding-direct-descriptors.js'
import { appendSheetRenameSourceTransform, directRegionIdsForFormula } from './formula-binding-dependency-helpers.js'
import type { FormulaBindingSheetIndex } from './formula-binding-sheet-index.js'
import type {
  BindPreparedFormulaOptions,
  CreateEngineFormulaBindingServiceArgs,
  FormulaOwnerPosition,
} from './formula-binding-service-types.js'

export function createFormulaBindingSheetRenameHandler(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly formulaSheetIndex: FormulaBindingSheetIndex
  readonly rangeDependenciesHaveNoFormulaMembers: (rangeDependencies: Uint32Array) => boolean
  readonly untrackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string | undefined,
    compiled: Pick<CompiledFormula, 'deps'> | undefined,
  ) => void
  readonly trackFormulaSheetIndexes: (cellIndex: number, ownerSheetName: string, compiled: Pick<CompiledFormula, 'deps'>) => void
  readonly canRetainUnmanagedCompiledPlan: (
    existingPlanId: number,
    compiled: CompiledFormula,
    directScalar: RuntimeDirectScalarDescriptor | undefined,
  ) => boolean
  readonly makeUnmanagedCompiledPlan: (source: string, compiled: CompiledFormula, templateId: number | undefined) => CompiledPlanRecord
  readonly rewriteFormulaMetadataPreservingRuntimeNow: (
    cellIndex: number,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    ownerPosition?: FormulaOwnerPosition,
  ) => boolean
  readonly bindPreparedFormulaNow: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    options?: BindPreparedFormulaOptions,
  ) => boolean
}): {
  readonly deferCellFormulasForSheetRenameNow: (oldSheetName: string, newSheetName: string) => number
  readonly rewriteCellFormulasForSheetRenameNow: (oldSheetName: string, newSheetName: string, formulaChangedCount: number) => number
} {
  const deferCellFormulasForSheetRenameNow = (oldSheetName: string, newSheetName: string): number => {
    let touchedCount = 0
    const movedSheetCells = args.formulaSheetIndex.moveSheetName(oldSheetName, newSheetName)
    if (movedSheetCells.references.size > 0) {
      movedSheetCells.references.forEach((cellIndex) => {
        const formula = args.serviceArgs.state.formulas.get(cellIndex)
        if (!formula) {
          return
        }
        appendSheetRenameSourceTransform(formula, oldSheetName, newSheetName)
        if (formula.directAggregate !== undefined) {
          formula.directAggregate = renameDirectAggregateDescriptorSheet({
            descriptor: formula.directAggregate,
            oldSheetName,
            newSheetName,
            regionGraph: args.serviceArgs.regionGraph,
          })
          args.serviceArgs.regionGraph.replaceFormulaSubscriptions(cellIndex, directRegionIdsForFormula(formula))
        }
        touchedCount += 1
      })
    }
    return touchedCount
  }

  const rewriteCellFormulasForSheetRenameNow = (oldSheetName: string, newSheetName: string, formulaChangedCount: number): number => {
    const referencedCandidates = args.formulaSheetIndex.getReferencingSheetSet(oldSheetName)
    const ownerCandidates = args.formulaSheetIndex.getOwnedBySheetSet(oldSheetName)
    const seen: number[] = []
    let seenSet: Set<number> | undefined
    const rewriteCandidate = (cellIndex: number): void => {
      if (seenSet) {
        if (seenSet.has(cellIndex)) {
          return
        }
        seenSet.add(cellIndex)
      } else {
        for (let index = 0; index < seen.length; index += 1) {
          if (seen[index] === cellIndex) {
            return
          }
        }
        seen.push(cellIndex)
        if (seen.length > 8) {
          seenSet = new Set(seen)
        }
      }
      const formula = args.serviceArgs.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      const ownerSheetName = args.serviceArgs.state.workbook.getSheetNameById(
        args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]!,
      )
      if (!ownerSheetName) {
        return
      }
      const previousOwnerSheetName = ownerSheetName === newSheetName ? oldSheetName : ownerSheetName
      if (
        args.serviceArgs.compiledPlans.isSoleOwner(formula.planId, formula.compiled) &&
        formula.compiled.symbolicNames.length === 0 &&
        formula.compiled.symbolicTables.length === 0 &&
        formula.compiled.symbolicSpills.length === 0 &&
        formula.directLookup === undefined &&
        formula.directCriteria === undefined &&
        (formula.directAggregate !== undefined || args.rangeDependenciesHaveNoFormulaMembers(formula.rangeDependencies))
      ) {
        if (previousOwnerSheetName !== ownerSheetName) {
          args.formulaSheetIndex.removeOwner(previousOwnerSheetName, cellIndex)
        }
        const sourceChanged = renameCompiledFormulaSheetReferenceMetadataInPlace(formula.compiled, oldSheetName, newSheetName)
        if (sourceChanged) {
          args.formulaSheetIndex.removeReference(oldSheetName, cellIndex)
          appendSheetRenameSourceTransform(formula, oldSheetName, newSheetName)
        }
        if (formula.directAggregate !== undefined) {
          formula.directAggregate = renameDirectAggregateDescriptorSheet({
            descriptor: formula.directAggregate,
            oldSheetName,
            newSheetName,
            regionGraph: args.serviceArgs.regionGraph,
          })
          args.serviceArgs.regionGraph.replaceFormulaSubscriptions(cellIndex, directRegionIdsForFormula(formula))
        }
        if (previousOwnerSheetName !== ownerSheetName) {
          args.formulaSheetIndex.appendOwner(ownerSheetName, cellIndex)
        }
        if (sourceChanged) {
          args.formulaSheetIndex.appendReference(newSheetName, cellIndex)
        }
        return
      }
      const renamedMetadata = renameCompiledFormulaSheetReferenceMetadata(formula.compiled, oldSheetName, newSheetName)
      if (!renamedMetadata.sourceChanged) {
        if (ownerSheetName === newSheetName) {
          args.untrackFormulaSheetIndexes(cellIndex, oldSheetName, formula.compiled)
          args.trackFormulaSheetIndexes(cellIndex, newSheetName, formula.compiled)
        }
        return
      }
      const position = args.serviceArgs.state.workbook.getCellPosition(cellIndex)
      const ownerPosition = position
        ? {
            sheetName: ownerSheetName,
            row: position.row,
            col: position.col,
          }
        : undefined
      const canPreserveRuntime =
        formula.directLookup === undefined &&
        formula.directCriteria === undefined &&
        (canRewriteCompiledPreservingBindings(formula, renamedMetadata.compiled) ||
          canRewriteCompiledPreservingDirectAggregate(formula, renamedMetadata.compiled) ||
          canRewriteCompiledPreservingDirectScalar(formula, renamedMetadata.compiled))
      if (canPreserveRuntime) {
        const nextDirectAggregate = formula.directAggregate
          ? renameDirectAggregateDescriptorSheet({
              descriptor: formula.directAggregate,
              oldSheetName,
              newSheetName,
              regionGraph: args.serviceArgs.regionGraph,
            })
          : undefined
        const nextDirectScalar =
          formula.directScalar === undefined
            ? undefined
            : buildDirectScalarDescriptor({
                compiled: renamedMetadata.compiled as ParsedCompiledFormula,
                ownerSheetName,
                ownerSheetId: args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex],
                workbook: args.serviceArgs.state.workbook,
                ensureCellTracked: args.serviceArgs.ensureCellTracked,
                ensureCellTrackedByCoords: args.serviceArgs.ensureCellTrackedByCoords,
              })
        if (formula.directScalar === undefined || nextDirectScalar !== undefined) {
          const plan = args.canRetainUnmanagedCompiledPlan(formula.planId, renamedMetadata.compiled, nextDirectScalar)
            ? args.makeUnmanagedCompiledPlan(formula.source, renamedMetadata.compiled, formula.templateId)
            : args.serviceArgs.compiledPlans.replace(formula.planId, formula.source, renamedMetadata.compiled, formula.templateId)
          args.untrackFormulaSheetIndexes(cellIndex, previousOwnerSheetName, formula.compiled)
          formula.compiled = plan.compiled
          formula.plan = plan
          formula.planId = plan.id
          formula.constants = plan.compiled.constants
          formula.constNumberLength = plan.compiled.constants.length
          formula.directAggregate = nextDirectAggregate
          formula.directScalar = nextDirectScalar
          appendSheetRenameSourceTransform(formula, oldSheetName, newSheetName)
          args.trackFormulaSheetIndexes(cellIndex, ownerSheetName, formula.compiled)
          if (formula.directAggregate !== undefined) {
            args.serviceArgs.regionGraph.replaceFormulaSubscriptions(cellIndex, directRegionIdsForFormula(formula))
          }
          return
        }
      }
      if (previousOwnerSheetName !== ownerSheetName) {
        args.untrackFormulaSheetIndexes(cellIndex, previousOwnerSheetName, formula.compiled)
      }
      const renamed = renameCompiledFormulaSheetReferences(formula.compiled, oldSheetName, newSheetName)
      const preserved =
        ownerPosition &&
        args.rewriteFormulaMetadataPreservingRuntimeNow(cellIndex, renamed.source, renamed.compiled, formula.templateId, ownerPosition)
      if (!preserved) {
        args.bindPreparedFormulaNow(cellIndex, ownerSheetName, renamed.source, renamed.compiled, formula.templateId)
      }
    }
    referencedCandidates?.forEach(rewriteCandidate)
    ownerCandidates?.forEach(rewriteCandidate)
    return formulaChangedCount
  }

  return {
    deferCellFormulasForSheetRenameNow,
    rewriteCellFormulasForSheetRenameNow,
  }
}
