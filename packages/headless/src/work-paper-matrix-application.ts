import type { EngineCellMutationRef } from '@bilig/core'
import { translateFormulaReferences } from '@bilig/formula'
import { buildMatrixMutationPlan, type MatrixMutationDimensionImpact } from './matrix-mutation-plan.js'
import { stripLeadingEquals } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperSheet } from './work-paper-types.js'

export interface WorkPaperCellMutationApplyOptions {
  captureUndo?: boolean
  potentialNewCells?: number
  source?: 'local' | 'restore'
  returnUndoOps?: boolean
  reuseRefs?: boolean
  skipDimensionUpdate?: boolean
}

export interface WorkPaperMatrixApplyOptions {
  captureUndo?: boolean
  deferLiteralAddresses?: ReadonlySet<string>
  skipNulls?: boolean
}

type MatrixMutationPlanInput = Parameters<typeof buildMatrixMutationPlan>[0]

export function applyWorkPaperSerializedMatrix(input: {
  readonly applyCellMutationRefs: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly applyRawContent: (address: WorkPaperCellAddress, content: RawCellContent) => void
  readonly flushPendingBatchOps: () => void
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
  readonly serialized: RawCellContent[][]
  readonly sourceAnchor: WorkPaperCellAddress
  readonly targetLeftCorner: WorkPaperCellAddress
}): void {
  const { serialized, sourceAnchor, targetLeftCorner } = input
  input.flushPendingBatchOps()
  const { refs, potentialNewCells } = buildMatrixMutationPlan({
    target: targetLeftCorner,
    content: serialized,
    rewriteFormula: (formula, destination, rowOffset, columnOffset) =>
      input.rewriteFormulaForStorage(
        translateFormulaReferences(
          stripLeadingEquals(formula),
          destination.row - (sourceAnchor.row + rowOffset),
          destination.col - (sourceAnchor.col + columnOffset),
        ),
        destination.sheet,
      ),
  })
  if (refs.length === 0) {
    return
  }
  input.applyCellMutationRefs(refs, {
    captureUndo: true,
    potentialNewCells,
    source: 'local',
    returnUndoOps: false,
    reuseRefs: true,
  })
}

export function applyWorkPaperMatrixContents(input: {
  readonly address: WorkPaperCellAddress
  readonly applyCellMutationRefs: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly content: WorkPaperSheet
  readonly flushPendingBatchOps: () => void
  readonly isEvaluationSuspended?: () => boolean
  readonly options?: WorkPaperMatrixApplyOptions
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
  readonly updateSheetDimensionsAfterCellMutationRefs?: (refs: readonly EngineCellMutationRef[]) => void
  readonly updateSheetDimensionsAfterMatrixMutationImpact?: (impact: MatrixMutationDimensionImpact) => void
}): void {
  const options = input.options ?? {}
  input.flushPendingBatchOps()
  const planInput: MatrixMutationPlanInput = {
    target: input.address,
    content: input.content,
    includeCombinedRefs: false,
    rewriteFormula: (formula, destination) => input.rewriteFormulaForStorage(stripLeadingEquals(formula), destination.sheet),
  }
  if (options.deferLiteralAddresses !== undefined) {
    planInput.deferLiteralAddresses = options.deferLiteralAddresses
  }
  if (options.skipNulls !== undefined) {
    planInput.skipNulls = options.skipNulls
  }
  const {
    leadingRefs,
    leadingFreshNumericRefCount,
    leadingPotentialNewCells,
    formulaRefs,
    formulaPotentialNewCells,
    refCount,
    dimensionImpact,
    potentialNewCells,
    trailingLiteralRefs,
    trailingLiteralPotentialNewCells,
  } = buildMatrixMutationPlan(planInput)
  if (refCount === 0) {
    return
  }
  const applyPlannedRefs = (phaseRefs: readonly EngineCellMutationRef[], applyOptions: WorkPaperCellMutationApplyOptions): void => {
    if (phaseRefs.length === 0) {
      return
    }
    input.applyCellMutationRefs(phaseRefs, applyOptions)
  }
  const phaseSource = options.captureUndo === false ? 'restore' : 'local'
  const createApplyOptions = (phasePotentialNewCells = potentialNewCells): WorkPaperCellMutationApplyOptions => {
    const applyOptions: WorkPaperCellMutationApplyOptions = {
      potentialNewCells: phasePotentialNewCells,
      source: phaseSource,
      returnUndoOps: false,
      reuseRefs: true,
    }
    if (options.captureUndo !== undefined) {
      applyOptions.captureUndo = options.captureUndo
    }
    return applyOptions
  }
  const updateSheetDimensionsAfterCellMutationRefs = input.updateSheetDimensionsAfterCellMutationRefs
  const canUpdateDimensionsOnce =
    updateSheetDimensionsAfterCellMutationRefs !== undefined && (phaseSource !== 'local' || input.isEvaluationSuspended?.() !== true)
  const updateDimensionsOnce = (refs: readonly EngineCellMutationRef[]): void => {
    if (input.updateSheetDimensionsAfterMatrixMutationImpact) {
      input.updateSheetDimensionsAfterMatrixMutationImpact(dimensionImpact)
    } else {
      updateSheetDimensionsAfterCellMutationRefs?.(refs)
    }
  }

  if (formulaRefs.length === 0) {
    applyPlannedRefs(leadingRefs, createApplyOptions())
    return
  }

  const canApplyFormulaMatrixInOnePass =
    trailingLiteralRefs.length === 0 &&
    !dimensionImpact.hasDynamicFormula &&
    (!canApplyLeadingRefsThroughFreshNumericFastPath(leadingRefs.length, leadingFreshNumericRefCount, leadingPotentialNewCells) ||
      canApplyFreshNumericAggregateMatrixInOnePass({
        formulaPotentialNewCells,
        formulaRefs,
        leadingFreshNumericRefCount,
        leadingPotentialNewCells,
        leadingRefs,
      }))
  if (canApplyFormulaMatrixInOnePass) {
    const mergedRefs = mergeMatrixMutationRefPhases(leadingRefs, formulaRefs, trailingLiteralRefs)
    const applyOptions = createApplyOptions()
    if (canUpdateDimensionsOnce) {
      applyOptions.skipDimensionUpdate = true
    }
    applyPlannedRefs(mergedRefs, applyOptions)
    if (canUpdateDimensionsOnce) {
      updateDimensionsOnce(mergedRefs)
    }
    return
  }

  const createPhasedApplyOptions = (phasePotentialNewCells: number): WorkPaperCellMutationApplyOptions => {
    const phasedOptions = createApplyOptions(phasePotentialNewCells)
    if (canUpdateDimensionsOnce) {
      phasedOptions.skipDimensionUpdate = true
    }
    return phasedOptions
  }

  applyPlannedRefs(leadingRefs, createPhasedApplyOptions(leadingPotentialNewCells))
  applyPlannedRefs(formulaRefs, createPhasedApplyOptions(formulaPotentialNewCells))
  applyPlannedRefs(trailingLiteralRefs, createPhasedApplyOptions(trailingLiteralPotentialNewCells))
  if (canUpdateDimensionsOnce) {
    updateDimensionsOnce(mergeMatrixMutationRefPhases(leadingRefs, formulaRefs, trailingLiteralRefs))
  }
}

function mergeMatrixMutationRefPhases(
  leadingRefs: readonly EngineCellMutationRef[],
  formulaRefs: readonly EngineCellMutationRef[],
  trailingLiteralRefs: readonly EngineCellMutationRef[],
): readonly EngineCellMutationRef[] {
  if (leadingRefs.length === 0 && trailingLiteralRefs.length === 0) {
    return formulaRefs
  }
  if (formulaRefs.length === 0 && trailingLiteralRefs.length === 0) {
    return leadingRefs
  }
  if (leadingRefs.length === 0 && formulaRefs.length === 0) {
    return trailingLiteralRefs
  }
  return [...leadingRefs, ...formulaRefs, ...trailingLiteralRefs]
}

function canApplyLeadingRefsThroughFreshNumericFastPath(
  leadingRefCount: number,
  leadingFreshNumericRefCount: number,
  leadingPotentialNewCells: number,
): boolean {
  return leadingRefCount >= 32 && leadingPotentialNewCells === leadingRefCount && leadingFreshNumericRefCount === leadingRefCount
}

function canApplyFreshNumericAggregateMatrixInOnePass(input: {
  readonly formulaPotentialNewCells: number
  readonly formulaRefs: readonly EngineCellMutationRef[]
  readonly leadingFreshNumericRefCount: number
  readonly leadingPotentialNewCells: number
  readonly leadingRefs: readonly EngineCellMutationRef[]
}): boolean {
  if (
    input.leadingRefs.length < 32 ||
    input.formulaRefs.length === 0 ||
    input.leadingPotentialNewCells !== input.leadingRefs.length ||
    input.leadingFreshNumericRefCount !== input.leadingRefs.length ||
    input.formulaPotentialNewCells !== input.formulaRefs.length
  ) {
    return false
  }
  const firstLeading = input.leadingRefs[0]
  const firstMutation = firstLeading?.mutation
  if (firstLeading === undefined || firstMutation?.kind !== 'setCellValue' || typeof firstMutation.value !== 'number') {
    return false
  }
  const firstFormula = input.formulaRefs[0]?.mutation
  if (firstFormula?.kind !== 'setCellFormula') {
    return false
  }
  let currentRow = firstMutation.row
  let currentWidth = 0
  let colCount = 0
  let rowCount = 1
  for (let index = 0; index < input.leadingRefs.length; index += 1) {
    const ref = input.leadingRefs[index]!
    const mutation = ref.mutation
    if (
      ref.sheetId !== firstLeading.sheetId ||
      mutation.kind !== 'setCellValue' ||
      typeof mutation.value !== 'number' ||
      Object.is(mutation.value, -0)
    ) {
      return false
    }
    if (mutation.row === currentRow) {
      if (mutation.col !== firstMutation.col + currentWidth) {
        return false
      }
      currentWidth += 1
    } else {
      if (mutation.row !== currentRow + 1 || mutation.col !== firstMutation.col || currentWidth === 0) {
        return false
      }
      if (colCount === 0) {
        colCount = currentWidth
      } else if (currentWidth !== colCount) {
        return false
      }
      currentRow = mutation.row
      currentWidth = 1
      rowCount += 1
    }
  }
  if (colCount === 0) {
    colCount = currentWidth
  } else if (currentWidth !== colCount) {
    return false
  }
  if (rowCount !== input.formulaRefs.length || colCount < 2 || input.leadingRefs.length !== rowCount * colCount) {
    return false
  }
  const formulaCol = firstMutation.col + colCount
  for (let index = 0; index < input.formulaRefs.length; index += 1) {
    const ref = input.formulaRefs[index]!
    const mutation = ref.mutation
    if (
      ref.sheetId !== firstLeading.sheetId ||
      ref.cellIndex !== undefined ||
      mutation.kind !== 'setCellFormula' ||
      mutation.row !== firstMutation.row + index ||
      mutation.col !== formulaCol
    ) {
      return false
    }
  }
  return true
}
