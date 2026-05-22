import type * as XLSX from 'xlsx'

export const externalWorkbookReferencesWarning = 'External workbook links were preserved but not recalculated during XLSX import.'
export const externalPivotCachesWarning = 'External pivot caches were detected but not semantically imported during XLSX import.'
export const unsupportedCellStylesWarning = 'Some cell styles were ignored during XLSX import.'
export const macroExecutionDeclinedWarning = 'Macros were preserved but not executed during XLSX import.'
export const dataTableFormulasWarning = 'Excel data table formulas were preserved but not recalculated during XLSX import.'
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
  return /(?:^|[^A-Z0-9_.])(?:NOW|RAND|RANDBETWEEN|TODAY)\s*\(/iu.test(formulaWithoutDoubleQuotedStrings(formula))
}

export function workbookDefinedNamesReferenceExternalWorkbook(workbook: XLSX.WorkBook): boolean {
  return (workbook.Workbook?.Names ?? []).some((entry) => {
    const ref = typeof entry.Ref === 'string' ? entry.Ref.trim() : ''
    return ref.length > 0 && formulaReferencesExternalWorkbook(ref)
  })
}
