import { describe, expect, it } from 'vitest'

import { largeTrustedCachedFormulaOpenModeThreshold, shouldUseCachedFormulaOpenMode } from '../xlsx-cached-formula-open-mode.js'

function trustedAudit(
  count: number,
  options: {
    readonly calcChainCount?: number
    readonly cacheStatus?: 'trustedCached' | 'staleRisk' | 'externalSubstitution' | 'engineRecomputed' | 'missing'
    readonly omitFormulaAddressAt?: number
    readonly mismatchedCalcChainAddressAt?: number
  } = {},
) {
  const calcChainCount = options.calcChainCount ?? count
  return {
    formulas: Array.from({ length: count }, (_, index) => ({
      context: 'worksheet-cell' as const,
      clause: '18.3.1.40',
      sheetName: 'Sheet1',
      ...(options.omitFormulaAddressAt === index ? {} : { address: `A${index + 1}` }),
      formula: 'B1',
      cachedValue: 1,
      cacheStatus: options.cacheStatus ?? ('trustedCached' as const),
    })),
    calcChain: {
      packagePath: 'xl/calcChain.xml',
      cells: Array.from({ length: calcChainCount }, (_, index) => ({
        sheetIndex: 1,
        sheetName: 'Sheet1',
        address: options.mismatchedCalcChainAddressAt === index ? `Z${index + 1}` : `A${index + 1}`,
      })),
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
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: largeTrustedCachedFormulaOpenModeThreshold,
        formulaCellCount: largeTrustedCachedFormulaOpenModeThreshold,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', calcId: 191029 },
        formulaAudit: trustedAudit(largeTrustedCachedFormulaOpenModeThreshold, { cacheStatus: 'staleRisk' }),
      }),
    ).toBe(false)
  })

  it('requires the calc chain to cover every trusted worksheet formula', () => {
    const formulaCount = largeTrustedCachedFormulaOpenModeThreshold
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: formulaCount,
        formulaCellCount: formulaCount,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', calcId: 191029 },
        formulaAudit: trustedAudit(formulaCount, { calcChainCount: 1 }),
      }),
    ).toBe(false)
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: formulaCount,
        formulaCellCount: formulaCount,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', calcId: 191029 },
        formulaAudit: trustedAudit(formulaCount, { mismatchedCalcChainAddressAt: 10 }),
      }),
    ).toBe(false)
    expect(
      shouldUseCachedFormulaOpenMode({
        cachedFormulaValueCount: formulaCount,
        formulaCellCount: formulaCount,
        calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern', calcId: 191029 },
        formulaAudit: trustedAudit(formulaCount, { omitFormulaAddressAt: 10 }),
      }),
    ).toBe(false)
  })
})
