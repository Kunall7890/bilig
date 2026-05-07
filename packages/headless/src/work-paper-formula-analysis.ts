import { parseFormula, serializeFormula } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import { WorkPaperNotAFormulaError, WorkPaperParseError } from './work-paper-errors.js'
import { collectFormulaNameRefs, stripLeadingEquals } from './work-paper-runtime-helpers.js'
import { compareSheetNames } from './work-paper-sheet-inspection.js'
import { calculateWorkPaperFormulaInScratchWorkbook, type WorkPaperScratchWorkbook } from './work-paper-scratch-evaluator.js'
import type { SerializedWorkPaperNamedExpression, WorkPaperConfig, WorkPaperSheets } from './work-paper-types.js'

export interface WorkPaperFormulaAnalysisHooks {
  readonly messageOf: (error: unknown, fallback: string) => string
}

export function normalizeWorkPaperFormula(formula: string, hooks: WorkPaperFormulaAnalysisHooks): string {
  if (!formula.trim().startsWith('=')) {
    throw new WorkPaperNotAFormulaError()
  }
  try {
    return `=${serializeFormula(parseFormula(stripLeadingEquals(formula)))}`
  } catch (error) {
    throw new WorkPaperParseError(hooks.messageOf(error, 'Unable to normalize formula'))
  }
}

export function calculateWorkPaperFormula(args: {
  readonly createWorkbook: (config: WorkPaperConfig) => WorkPaperScratchWorkbook
  readonly config: WorkPaperConfig
  readonly serializedSheets: WorkPaperSheets
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly formula: string
  readonly scope?: number
  readonly messageOf: (error: unknown, fallback: string) => string
}): CellValue | CellValue[][] {
  if (!args.formula.trim().startsWith('=')) {
    throw new WorkPaperNotAFormulaError()
  }
  try {
    return calculateWorkPaperFormulaInScratchWorkbook(args)
  } catch (error) {
    throw new WorkPaperParseError(args.messageOf(error, 'Unable to calculate formula'))
  }
}

export function getWorkPaperNamedExpressionsFromFormula(formula: string, hooks: WorkPaperFormulaAnalysisHooks): string[] {
  if (!formula.trim().startsWith('=')) {
    throw new WorkPaperNotAFormulaError()
  }
  try {
    const parsed = parseFormula(stripLeadingEquals(formula))
    const names = new Set<string>()
    collectFormulaNameRefs(parsed, names)
    return [...names].toSorted(compareSheetNames)
  } catch (error) {
    throw new WorkPaperParseError(hooks.messageOf(error, 'Unable to inspect formula'))
  }
}

export function validateWorkPaperFormula(formula: string): boolean {
  if (!formula.trim().startsWith('=')) {
    return false
  }
  try {
    parseFormula(stripLeadingEquals(formula))
    return true
  } catch {
    return false
  }
}
