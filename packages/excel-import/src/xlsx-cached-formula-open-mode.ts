import { formulaShouldUseCachedUnsupportedFunctionValue } from '@bilig/core/headless-runtime'
import type { WorkbookCalculationSettingsSnapshot, WorkbookFormulaAuditSnapshot } from '@bilig/protocol'
import { formulaReferencesExternalWorkbook, formulaReferencesVolatileFunction } from './xlsx-import-warnings.js'

export const largeTrustedCachedFormulaOpenModeThreshold = 50_000

export function shouldUseCachedFormulaOpenMode(args: {
  readonly cachedFormulaValueCount: number
  readonly formulaCellCount: number
  readonly calculationSettings: WorkbookCalculationSettingsSnapshot | undefined
  readonly formulaAudit: WorkbookFormulaAuditSnapshot | undefined
}): boolean {
  if (args.cachedFormulaValueCount === 0 || args.formulaCellCount === 0 || args.calculationSettings === undefined) {
    return false
  }
  if (args.calculationSettings.fullCalcOnLoad === false || args.calculationSettings.mode === 'manual') {
    return true
  }
  if (
    args.calculationSettings.fullCalcOnLoad === true ||
    args.calculationSettings.forceFullCalc === true ||
    args.calculationSettings.fullPrecision === false
  ) {
    return false
  }
  if (args.formulaCellCount < largeTrustedCachedFormulaOpenModeThreshold) {
    return false
  }
  if (args.cachedFormulaValueCount !== args.formulaCellCount) {
    return false
  }
  const calcChainCellCount = args.formulaAudit?.calcChain?.cells.length ?? 0
  if (calcChainCellCount === 0) {
    return false
  }
  const formulaAudit = args.formulaAudit
  const worksheetFormulaEntries = formulaAudit?.formulas.filter((entry) => entry.context === 'worksheet-cell') ?? []
  if (formulaAuditHasRiskyDefinedNameFormula(formulaAudit)) {
    return false
  }
  return (
    worksheetFormulaEntries.length === args.formulaCellCount &&
    worksheetFormulaEntries.every((entry) => isSafeTrustedCachedWorksheetFormula(entry)) &&
    calcChainExactlyCoversWorksheetFormulas(formulaAudit, worksheetFormulaEntries)
  )
}

function isSafeTrustedCachedWorksheetFormula(entry: WorkbookFormulaAuditSnapshot['formulas'][number]): boolean {
  if (entry.cacheStatus !== 'trustedCached') {
    return false
  }
  if (entry.formulaType === 'array' || entry.formulaType === 'dataTable') {
    return false
  }
  return !formulaReferencesExternalWorkbook(entry.formula) && !formulaReferencesVolatileFunction(entry.formula)
}

function formulaAuditHasRiskyDefinedNameFormula(formulaAudit: WorkbookFormulaAuditSnapshot | undefined): boolean {
  const definedNameEntries = formulaAudit?.formulas.filter((entry) => entry.context === 'defined-name') ?? []
  if (definedNameEntries.length === 0) {
    return false
  }
  const definedFormulaNames = new Set(
    definedNameEntries.flatMap((entry) => {
      const normalizedName = entry.name?.trim().toUpperCase()
      return normalizedName ? [normalizedName] : []
    }),
  )
  return definedNameEntries.some((entry) => formulaHasExcelOnlyCachedOpenSemantics(entry.formula, definedFormulaNames))
}

function formulaHasExcelOnlyCachedOpenSemantics(formula: string, definedFormulaNames: ReadonlySet<string>): boolean {
  return (
    formulaReferencesExternalWorkbook(formula) ||
    formulaReferencesVolatileFunction(formula) ||
    formulaShouldUseCachedUnsupportedFunctionValue(formula, definedFormulaNames)
  )
}

function calcChainExactlyCoversWorksheetFormulas(
  formulaAudit: WorkbookFormulaAuditSnapshot | undefined,
  worksheetFormulaEntries: WorkbookFormulaAuditSnapshot['formulas'],
): boolean {
  const calcChainCells = formulaAudit?.calcChain?.cells ?? []
  if (calcChainCells.length !== worksheetFormulaEntries.length) {
    return false
  }

  const formulaKeys = new Set<string>()
  for (const entry of worksheetFormulaEntries) {
    const key = formulaAddressKey(entry.sheetName, entry.address)
    if (key === undefined || formulaKeys.has(key)) {
      return false
    }
    formulaKeys.add(key)
  }

  const calcChainKeys = new Set<string>()
  for (const cell of calcChainCells) {
    const key = formulaAddressKey(cell.sheetName, cell.address)
    if (key === undefined || calcChainKeys.has(key) || !formulaKeys.has(key)) {
      return false
    }
    calcChainKeys.add(key)
  }
  return calcChainKeys.size === formulaKeys.size
}

function formulaAddressKey(sheetName: string | undefined, address: string | undefined): string | undefined {
  if (!sheetName || !address) {
    return undefined
  }
  return `${sheetName}\u0000${address.toUpperCase()}`
}
