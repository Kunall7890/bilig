import type { EngineCellMutationRef, EngineFreshDirectAggregateMatrixPlan } from '@bilig/core/headless-runtime'
import { columnToIndex, translateFormulaReferences } from '@bilig/formula'
import { buildMatrixMutationPlan, type MatrixMutationDimensionImpact } from './matrix-mutation-plan.js'
import { workPaperFormulaMayResizeDynamically } from './work-paper-sheet-inspection.js'
import { stripLeadingEquals } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperSheet } from './work-paper-types.js'

export interface WorkPaperCellMutationApplyOptions {
  captureUndo?: boolean
  potentialNewCells?: number
  source?: 'local' | 'restore'
  returnUndoOps?: boolean
  reuseRefs?: boolean
  skipDimensionUpdate?: boolean
  freshDirectAggregateMatrixPlan?: EngineFreshDirectAggregateMatrixPlan
}

export interface WorkPaperMatrixApplyOptions {
  captureUndo?: boolean
  deferLiteralAddresses?: ReadonlySet<string>
  skipNulls?: boolean
  trustedFreshCells?: boolean
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
  const phaseSource = options.captureUndo === false ? 'restore' : 'local'
  const createApplyOptions = (phasePotentialNewCells: number): WorkPaperCellMutationApplyOptions => {
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
  const updateDimensionsOnce = (refs: readonly EngineCellMutationRef[], impact: MatrixMutationDimensionImpact): void => {
    if (input.updateSheetDimensionsAfterMatrixMutationImpact) {
      input.updateSheetDimensionsAfterMatrixMutationImpact(impact)
    } else {
      updateSheetDimensionsAfterCellMutationRefs?.(refs)
    }
  }
  const freshNumericFormulaPlan = tryBuildFreshNumericFormulaColumnMatrixPlan(planInput, options.trustedFreshCells === true)
  if (freshNumericFormulaPlan !== undefined) {
    const applyOptions = createApplyOptions(freshNumericFormulaPlan.potentialNewCells)
    applyOptions.freshDirectAggregateMatrixPlan = freshNumericFormulaPlan.freshDirectAggregateMatrixPlan
    if (canUpdateDimensionsOnce) {
      applyOptions.skipDimensionUpdate = true
    }
    input.applyCellMutationRefs(freshNumericFormulaPlan.refs, applyOptions)
    if (canUpdateDimensionsOnce) {
      updateDimensionsOnce(freshNumericFormulaPlan.refs, freshNumericFormulaPlan.dimensionImpact)
    }
    return
  }
  const {
    leadingRefs,
    leadingFreshNumericRefCount,
    leadingPotentialNewCells,
    canApplyFreshNumericAggregateMatrixInOnePass,
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

  if (formulaRefs.length === 0) {
    applyPlannedRefs(leadingRefs, createApplyOptions(potentialNewCells))
    return
  }

  const canApplyFormulaMatrixInOnePass =
    trailingLiteralRefs.length === 0 &&
    !dimensionImpact.hasDynamicFormula &&
    (!canApplyLeadingRefsThroughFreshNumericFastPath(leadingRefs.length, leadingFreshNumericRefCount, leadingPotentialNewCells) ||
      canApplyFreshNumericAggregateMatrixInOnePass)
  if (canApplyFormulaMatrixInOnePass) {
    const mergedRefs = mergeMatrixMutationRefPhases(leadingRefs, formulaRefs, trailingLiteralRefs)
    const applyOptions = createApplyOptions(potentialNewCells)
    if (canUpdateDimensionsOnce) {
      applyOptions.skipDimensionUpdate = true
    }
    applyPlannedRefs(mergedRefs, applyOptions)
    if (canUpdateDimensionsOnce) {
      updateDimensionsOnce(mergedRefs, dimensionImpact)
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
    updateDimensionsOnce(mergeMatrixMutationRefPhases(leadingRefs, formulaRefs, trailingLiteralRefs), dimensionImpact)
  }
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === 'string' && content.trim().startsWith('=')
}

function isPotentialDirectAggregateFormulaContent(content: RawCellContent): content is string {
  if (!isFormulaContent(content)) {
    return false
  }
  const normalized = stripLeadingEquals(content).trimStart().toUpperCase()
  return (
    normalized.startsWith('SUM(') ||
    normalized.startsWith('AVERAGE(') ||
    normalized.startsWith('COUNT(') ||
    normalized.startsWith('COUNTA(') ||
    normalized.startsWith('MIN(') ||
    normalized.startsWith('MAX(')
  )
}

interface FreshMatrixPrecomputedResultShape {
  readonly aggregateKind: NonNullable<EngineFreshDirectAggregateMatrixPlan['precomputedFormulaResults']>['aggregateKind']
  readonly aggregateColStart: number
  readonly aggregateColEnd: number
  readonly resultOffset: number | undefined
}

const FRESH_MATRIX_DIRECT_AGGREGATE_SOURCE_RE =
  /^=?(SUM|AVERAGE|AVG|COUNT|MIN|MAX)\s*\(\s*([A-Za-z]+)([1-9]\d*):([A-Za-z]+)([1-9]\d*)\s*\)(?:\s*\+\s*([+-]?(?:\d+|\d*\.\d+)))?\s*$/i
const PRECOMPUTED_FRESH_DIRECT_AGGREGATE_MATRIX_MIN_ROWS = 128

function tryBuildFreshNumericFormulaColumnMatrixPlan(
  args: MatrixMutationPlanInput,
  trustedFreshCells: boolean,
):
  | {
      readonly refs: readonly EngineCellMutationRef[]
      readonly freshDirectAggregateMatrixPlan: EngineFreshDirectAggregateMatrixPlan
      readonly potentialNewCells: number
      readonly dimensionImpact: MatrixMutationDimensionImpact
    }
  | undefined {
  if (args.deferLiteralAddresses !== undefined || args.skipNulls === true || args.content.length === 0) {
    return undefined
  }
  const firstWidth = args.content[0]?.length ?? 0
  if (firstWidth < 3) {
    return undefined
  }
  const inputColCount = firstWidth - 1
  const rowCount = args.content.length
  const valueCount = rowCount * inputColCount
  const values = new Float64Array(valueCount)
  const canPrecomputeNativeSizedMatrix = trustedFreshCells && rowCount >= PRECOMPUTED_FRESH_DIRECT_AGGREGATE_MATRIX_MIN_ROWS
  const precomputedResults = canPrecomputeNativeSizedMatrix ? new Float64Array(rowCount) : undefined
  const refs: EngineCellMutationRef[] = []
  refs.length = valueCount + rowCount
  let valueCursor = 0
  let formulaCursor = valueCount
  let precomputedShape: FreshMatrixPrecomputedResultShape | undefined
  let canUsePrecomputedResults = canPrecomputeNativeSizedMatrix
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    const row = args.content[rowOffset]!
    if (row.length !== firstWidth) {
      return undefined
    }
    const destinationRow = args.target.row + rowOffset
    for (let columnOffset = 0; columnOffset < inputColCount; columnOffset += 1) {
      const raw = row[columnOffset]!
      if (typeof raw !== 'number' || Object.is(raw, -0)) {
        return undefined
      }
      refs[valueCursor] = {
        sheetId: args.target.sheet,
        mutation: {
          kind: 'setCellValue',
          row: destinationRow,
          col: args.target.col + columnOffset,
          value: raw,
        },
      }
      values[valueCursor] = raw
      valueCursor += 1
    }
    const rawFormula = row[inputColCount]!
    if (!isPotentialDirectAggregateFormulaContent(rawFormula)) {
      return undefined
    }
    const destination: WorkPaperCellAddress = {
      sheet: args.target.sheet,
      row: destinationRow,
      col: args.target.col + inputColCount,
    }
    const rewrittenFormula = args.rewriteFormula(rawFormula, destination, rowOffset, inputColCount)
    if (workPaperFormulaMayResizeDynamically(rewrittenFormula)) {
      return undefined
    }
    if (canUsePrecomputedResults && precomputedResults !== undefined) {
      const precomputed = tryEvaluateFreshMatrixDirectAggregateFormula({
        formula: rewrittenFormula,
        formulaCol: destination.col,
        inputColCount,
        matrixColStart: args.target.col,
        row: destination.row,
        rowOffset,
        shape: precomputedShape,
        values,
      })
      if (precomputed === undefined) {
        canUsePrecomputedResults = false
      } else {
        precomputedShape = precomputed.shape
        precomputedResults[rowOffset] = precomputed.result
      }
    }
    refs[formulaCursor] = {
      sheetId: args.target.sheet,
      mutation: {
        kind: 'setCellFormula',
        row: destinationRow,
        col: destination.col,
        formula: rewrittenFormula,
      },
    }
    formulaCursor += 1
  }
  return {
    refs,
    freshDirectAggregateMatrixPlan: {
      sheetId: args.target.sheet,
      rowStart: args.target.row,
      rowCount,
      colStart: args.target.col,
      inputColCount,
      ...(canUsePrecomputedResults && precomputedShape !== undefined && precomputedResults !== undefined
        ? {
            precomputedFormulaResults: {
              aggregateKind: precomputedShape.aggregateKind,
              aggregateColStart: precomputedShape.aggregateColStart,
              aggregateColEnd: precomputedShape.aggregateColEnd,
              ...(precomputedShape.resultOffset === undefined ? {} : { resultOffset: precomputedShape.resultOffset }),
              results: precomputedResults,
            },
          }
        : {}),
      ...(trustedFreshCells ? { trustedFreshCells: true } : {}),
      values,
    },
    potentialNewCells: refs.length,
    dimensionImpact: {
      hasDynamicFormula: false,
      maxClearCol: -1,
      maxClearRow: -1,
      maxSetCol: args.target.col + inputColCount,
      maxSetRow: args.target.row + rowCount - 1,
      sheetId: args.target.sheet,
    },
  }
}

function tryEvaluateFreshMatrixDirectAggregateFormula(input: {
  readonly formula: string
  readonly formulaCol: number
  readonly inputColCount: number
  readonly matrixColStart: number
  readonly row: number
  readonly rowOffset: number
  readonly shape: FreshMatrixPrecomputedResultShape | undefined
  readonly values: Float64Array
}): { readonly result: number; readonly shape: FreshMatrixPrecomputedResultShape } | undefined {
  const match = FRESH_MATRIX_DIRECT_AGGREGATE_SOURCE_RE.exec(input.formula.trim())
  if (!match) {
    return undefined
  }
  const aggregateKind = directAggregateKindFromSourceCallee(match[1]!)
  if (aggregateKind === undefined) {
    return undefined
  }
  const aggregateColStart = columnToIndex(match[2]!.toUpperCase())
  const aggregateRowStart = parseFreshMatrixA1RowIndex(match[3]!)
  const aggregateColEnd = columnToIndex(match[4]!.toUpperCase())
  const aggregateRowEnd = parseFreshMatrixA1RowIndex(match[5]!)
  const resultOffset = match[6] === undefined ? undefined : normalizeFreshMatrixResultOffset(Number(match[6]))
  if (
    aggregateRowStart !== input.row ||
    aggregateRowEnd !== input.row ||
    aggregateColStart < input.matrixColStart ||
    aggregateColEnd >= input.formulaCol ||
    aggregateColStart > aggregateColEnd ||
    aggregateColEnd - input.matrixColStart >= input.inputColCount
  ) {
    return undefined
  }
  const shape = {
    aggregateKind,
    aggregateColStart,
    aggregateColEnd,
    resultOffset,
  }
  if (
    input.shape !== undefined &&
    (input.shape.aggregateKind !== shape.aggregateKind ||
      input.shape.aggregateColStart !== shape.aggregateColStart ||
      input.shape.aggregateColEnd !== shape.aggregateColEnd ||
      (input.shape.resultOffset ?? 0) !== (shape.resultOffset ?? 0))
  ) {
    return undefined
  }
  const rowBase = input.rowOffset * input.inputColCount
  let sum = 0
  let count = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  for (let col = aggregateColStart; col <= aggregateColEnd; col += 1) {
    const value = input.values[rowBase + col - input.matrixColStart]!
    sum += value
    count += 1
    minimum = Math.min(minimum, value)
    maximum = Math.max(maximum, value)
  }
  const result =
    aggregateKind === 'sum'
      ? sum
      : aggregateKind === 'count'
        ? count
        : aggregateKind === 'average'
          ? count === 0
            ? undefined
            : sum / count
          : aggregateKind === 'min'
            ? minimum === Number.POSITIVE_INFINITY
              ? 0
              : minimum
            : maximum === Number.NEGATIVE_INFINITY
              ? 0
              : maximum
  return result === undefined ? undefined : { result: result + (resultOffset ?? 0), shape }
}

function directAggregateKindFromSourceCallee(
  callee: string,
): NonNullable<EngineFreshDirectAggregateMatrixPlan['precomputedFormulaResults']>['aggregateKind'] | undefined {
  switch (callee.toUpperCase()) {
    case 'SUM':
      return 'sum'
    case 'AVERAGE':
    case 'AVG':
      return 'average'
    case 'COUNT':
      return 'count'
    case 'MIN':
      return 'min'
    case 'MAX':
      return 'max'
    default:
      return undefined
  }
}

function normalizeFreshMatrixResultOffset(offset: number): number | undefined {
  return Object.is(offset, 0) ? undefined : offset
}

function parseFreshMatrixA1RowIndex(source: string): number {
  if (!/^[1-9]\d*$/.test(source)) {
    return -1
  }
  const rowNumber = Number(source)
  return Number.isSafeInteger(rowNumber) ? rowNumber - 1 : -1
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
