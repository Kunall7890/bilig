import type { EngineCellMutationRef } from '@bilig/core'
import { isBlankRawCellContent } from './work-paper-runtime-helpers.js'
import { workPaperFormulaMayResizeDynamically } from './work-paper-sheet-inspection.js'
import type { WorkPaperCellAddress, WorkPaperSheet, RawCellContent } from './work-paper-types.js'

export type MatrixMutationRef = EngineCellMutationRef

export interface MatrixMutationPlan {
  dimensionImpact: MatrixMutationDimensionImpact
  leadingRefs: MatrixMutationRef[]
  leadingPotentialNewCells: number
  formulaRefs: MatrixMutationRef[]
  formulaPotentialNewCells: number
  refCount: number
  refs: MatrixMutationRef[]
  potentialNewCells: number
  trailingLiteralRefs: MatrixMutationRef[]
  trailingLiteralPotentialNewCells: number
}

export interface MatrixMutationDimensionImpact {
  hasDynamicFormula: boolean
  maxClearCol: number
  maxClearRow: number
  maxSetCol: number
  maxSetRow: number
  sheetId: number
}

interface BuildMatrixMutationPlanArgs {
  target: WorkPaperCellAddress
  content: WorkPaperSheet
  rewriteFormula: (formula: string, destination: WorkPaperCellAddress, rowOffset: number, columnOffset: number) => string
  deferLiteralAddresses?: ReadonlySet<string>
  includeCombinedRefs?: boolean
  skipNulls?: boolean
}

function isFormulaContent(content: RawCellContent): content is string {
  return typeof content === 'string' && content.trim().startsWith('=')
}

export function buildMatrixMutationPlan(args: BuildMatrixMutationPlanArgs): MatrixMutationPlan {
  const leadingRefs: MatrixMutationRef[] = []
  const formulaRefs: MatrixMutationRef[] = []
  const trailingLiteralRefs: MatrixMutationRef[] = []
  let leadingPotentialNewCells = 0
  let formulaPotentialNewCells = 0
  let potentialNewCells = 0
  let trailingLiteralPotentialNewCells = 0
  let hasDynamicFormula = false
  let maxClearCol = -1
  let maxClearRow = -1
  let maxSetCol = -1
  let maxSetRow = -1
  const earliestFormulaRowByColumn: number[] = []

  const shouldDeferLiteral = (row: number, col: number): boolean => {
    const earliestFormulaRow = earliestFormulaRowByColumn[col]
    return earliestFormulaRow !== undefined && row > earliestFormulaRow
  }

  const shouldDeferLiteralAddress = (row: number, col: number): boolean => {
    const explicitAddresses = args.deferLiteralAddresses
    return explicitAddresses !== undefined && explicitAddresses.has(formatMatrixPlanAddress(row, col))
  }

  args.content.forEach((row, rowOffset) => {
    row.forEach((raw, columnOffset) => {
      const destination: WorkPaperCellAddress = {
        sheet: args.target.sheet,
        row: args.target.row + rowOffset,
        col: args.target.col + columnOffset,
      }

      if (isBlankRawCellContent(raw)) {
        if (!args.skipNulls) {
          maxClearRow = Math.max(maxClearRow, destination.row)
          maxClearCol = Math.max(maxClearCol, destination.col)
          const ref = {
            sheetId: args.target.sheet,
            mutation: { kind: 'clearCell', row: destination.row, col: destination.col },
          } satisfies MatrixMutationRef
          if (shouldDeferLiteral(destination.row, destination.col) || shouldDeferLiteralAddress(destination.row, destination.col)) {
            trailingLiteralRefs.push(ref)
          } else {
            leadingRefs.push(ref)
          }
        }
        return
      }

      potentialNewCells += 1
      maxSetRow = Math.max(maxSetRow, destination.row)
      maxSetCol = Math.max(maxSetCol, destination.col)

      if (isFormulaContent(raw)) {
        formulaPotentialNewCells += 1
        const rewrittenFormula = args.rewriteFormula(raw, destination, rowOffset, columnOffset)
        hasDynamicFormula ||= workPaperFormulaMayResizeDynamically(rewrittenFormula)
        const earliestFormulaRow = earliestFormulaRowByColumn[destination.col]
        if (earliestFormulaRow === undefined || destination.row < earliestFormulaRow) {
          earliestFormulaRowByColumn[destination.col] = destination.row
        }
        formulaRefs.push({
          sheetId: args.target.sheet,
          mutation: {
            kind: 'setCellFormula',
            row: destination.row,
            col: destination.col,
            formula: rewrittenFormula,
          },
        })
        return
      }

      const ref = {
        sheetId: args.target.sheet,
        mutation: {
          kind: 'setCellValue',
          row: destination.row,
          col: destination.col,
          value: raw,
        },
      } satisfies MatrixMutationRef
      if (shouldDeferLiteral(destination.row, destination.col) || shouldDeferLiteralAddress(destination.row, destination.col)) {
        trailingLiteralPotentialNewCells += 1
        trailingLiteralRefs.push(ref)
      } else {
        leadingPotentialNewCells += 1
        leadingRefs.push(ref)
      }
    })
  })

  return {
    dimensionImpact: {
      hasDynamicFormula,
      maxClearCol,
      maxClearRow,
      maxSetCol,
      maxSetRow,
      sheetId: args.target.sheet,
    },
    leadingRefs,
    leadingPotentialNewCells,
    formulaRefs,
    formulaPotentialNewCells,
    refCount: leadingRefs.length + formulaRefs.length + trailingLiteralRefs.length,
    refs: args.includeCombinedRefs === false ? [] : [...leadingRefs, ...formulaRefs, ...trailingLiteralRefs],
    potentialNewCells,
    trailingLiteralRefs,
    trailingLiteralPotentialNewCells,
  }
}

function formatMatrixPlanAddress(row: number, col: number): string {
  let index = col
  let label = ''
  do {
    label = String.fromCharCode(65 + (index % 26)) + label
    index = Math.floor(index / 26) - 1
  } while (index >= 0)
  return `${label}${row + 1}`
}
