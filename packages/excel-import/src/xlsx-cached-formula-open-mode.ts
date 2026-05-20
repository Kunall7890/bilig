import type { WorkbookCalculationSettingsSnapshot, WorkbookFormulaAuditSnapshot } from '@bilig/protocol'

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
  const worksheetFormulaEntries = args.formulaAudit?.formulas.filter((entry) => entry.context === 'worksheet-cell') ?? []
  return (
    worksheetFormulaEntries.length === args.formulaCellCount &&
    worksheetFormulaEntries.every((entry) => entry.cacheStatus === 'trustedCached')
  )
}
