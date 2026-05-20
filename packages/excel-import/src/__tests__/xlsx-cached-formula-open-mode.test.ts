import { describe, expect, it } from 'vitest'

import { largeTrustedCachedFormulaOpenModeThreshold, shouldUseCachedFormulaOpenMode } from '../xlsx-cached-formula-open-mode.js'

function trustedAudit(count: number) {
  return {
    formulas: Array.from({ length: count }, (_, index) => ({
      context: 'worksheet-cell' as const,
      clause: '18.3.1.40',
      sheetName: 'Sheet1',
      address: `A${index + 1}`,
      formula: 'B1',
      cachedValue: 1,
      cacheStatus: 'trustedCached' as const,
    })),
    calcChain: {
      packagePath: 'xl/calcChain.xml',
      cells: [{ sheetIndex: 1, sheetName: 'Sheet1', address: 'A1' }],
    },
  }
}

describe('cached formula open mode policy', () => {
  it('uses no-recalc open mode for explicitly manual or no-full-calc imports', () => {
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: 1,
        formulaCellCount: 1,
        calculationSettings: { mode: 'manual', compatibilityMode: 'excel-modern' },
        formulaAudit: undefined,
      }),
    ).toBe(true)
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: 1,
        formulaCellCount: 1,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', fullCalcOnLoad: false },
        formulaAudit: undefined,
      }),
    ).toBe(true)
  })

  it('uses trusted cached values for large automatic workbooks with complete caches and a calc chain', () => {
    const formulaCount = largeTrustedCachedFormulaOpenModeThreshold
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: formulaCount,
        formulaCellCount: formulaCount,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', calcId: 191029 },
        formulaAudit: trustedAudit(formulaCount),
      }),
    ).toBe(true)
  })

  it('does not use arbitrary cached values for small or stale-risk automatic workbooks', () => {
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: 1,
        formulaCellCount: 1,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern' },
        formulaAudit: trustedAudit(1),
      }),
    ).toBe(false)
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: largeTrustedCachedFormulaOpenModeThreshold,
        formulaCellCount: largeTrustedCachedFormulaOpenModeThreshold,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', forceFullCalc: true },
        formulaAudit: trustedAudit(largeTrustedCachedFormulaOpenModeThreshold),
      }),
    ).toBe(false)
  })
})
