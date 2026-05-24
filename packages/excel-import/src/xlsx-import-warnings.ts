import type * as XLSX from 'xlsx'

import { formulaShouldUseCachedUnsupportedFunctionValue } from '@bilig/core/headless-runtime'
import { FORMULA_VOLATILE_FUNCTION_NAMES } from '@bilig/formula'
import type { WorkbookFormulaAuditSnapshot } from '@bilig/protocol'

export const externalWorkbookReferencesWarning = 'External workbook links were preserved but not recalculated during XLSX import.'
export const externalPivotCachesWarning = 'External pivot caches were detected but not semantically imported during XLSX import.'
export const unsupportedCellStylesWarning = 'Some cell styles were ignored during XLSX import.'
export const macroExecutionDeclinedWarning = 'Macros were preserved but not executed during XLSX import.'
export const dataTableFormulasWarning = 'Excel data table formulas were preserved but not recalculated during XLSX import.'
export const unsupportedFormulaCachesWarning =
  'Unsupported formulas were preserved from cached XLSX values; recalculation may return Excel error values.'
export const definedNameFormulaCachesWarning =
  'Defined-name formulas contain volatile, external, or unsupported Excel semantics; dependent cached formula values may change during recalculation.'
export const volatileFormulasWarning =
  'Volatile formulas were preserved during XLSX import; cached formula values may depend on workbook calculation time.'

export function addWorkbookWarnings(workbook: XLSX.WorkBook, warnings: string[], ignoredDefinedNameCount: number): void {
  if (workbook.vbaraw) {
    warnings.push(macroExecutionDeclinedWarning)
  }
  if (ignoredDefinedNameCount > 0) {
    warnings.push('Some defined names were ignored during XLSX import.')
  }
}

function formulaWithoutDoubleQuotedStrings(formula: string): string {
  let stripped = ''
  let index = 0
  while (index < formula.length) {
    if (formula[index] !== '"') {
      stripped += formula[index]
      index += 1
      continue
    }
    stripped += ' '
    index += 1
    while (index < formula.length) {
      if (formula[index] === '"' && formula[index + 1] === '"') {
        stripped += '  '
        index += 2
        continue
      }
      stripped += ' '
      if (formula[index] === '"') {
        index += 1
        break
      }
      index += 1
    }
  }
  return stripped
}

export function formulaReferencesExternalWorkbook(formula: string): boolean {
  return /(?:^|[^A-Za-z0-9_])(?:'?\[[^\]\r\n]+\][^'!\r\n]*'?)!/u.test(formulaWithoutDoubleQuotedStrings(formula))
}

export function formulaReferencesVolatileFunction(formula: string): boolean {
  return volatileFunctionPattern.test(formulaWithoutDoubleQuotedStrings(formula))
}

export function readImportedFormulaAuditWarnings(formulaAudit: WorkbookFormulaAuditSnapshot | undefined): string[] {
  if (!formulaAudit) {
    return []
  }
  const warnings: string[] = []
  if (formulaAudit.diagnostics?.some((diagnostic) => diagnostic.code === 'unsupported-formula-cache')) {
    warnings.push(unsupportedFormulaCachesWarning)
  }
  if (formulaAuditHasRiskyDefinedNameFormula(formulaAudit)) {
    warnings.push(definedNameFormulaCachesWarning)
  }
  return warnings
}

export function formulaAuditHasRiskyDefinedNameFormula(formulaAudit: WorkbookFormulaAuditSnapshot | undefined): boolean {
  if (!formulaAudit) {
    return false
  }
  const definedNameEntries = formulaAudit.formulas.filter((entry) => entry.context === 'defined-name')
  if (definedNameEntries.length === 0) {
    return false
  }
  const definedFormulaNames = new Set(
    definedNameEntries.flatMap((entry) => {
      const normalizedName = entry.name?.trim().toUpperCase()
      return normalizedName ? [normalizedName] : []
    }),
  )
  return definedNameEntries.some(
    (entry) =>
      formulaReferencesExternalWorkbook(entry.formula) ||
      formulaReferencesVolatileFunction(entry.formula) ||
      formulaShouldUseCachedUnsupportedFunctionValue(entry.formula, definedFormulaNames),
  )
}

const volatileFunctionPattern = new RegExp(
  `(?:^|[^A-Z0-9_.])(?:_xlfn\\.)?(?:${FORMULA_VOLATILE_FUNCTION_NAMES.map(escapeRegExp).join('|')})\\s*\\(`,
  'iu',
)

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

export function workbookDefinedNamesReferenceExternalWorkbook(workbook: XLSX.WorkBook): boolean {
  return (workbook.Workbook?.Names ?? []).some((entry) => {
    const ref = typeof entry.Ref === 'string' ? entry.Ref.trim() : ''
    return ref.length > 0 && formulaReferencesExternalWorkbook(ref)
  })
}
