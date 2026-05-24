import { describe, expect, it } from 'vitest'

import { memoryGateRssBudgets } from '../public-workbook-corpus-memory-gate.ts'

const mib = 1024 * 1024

describe('public workbook corpus memory gate budgets', () => {
  it('keeps public gates strict and synthetic gates explicit about runtime RSS headroom', () => {
    expect(memoryGateRssBudgets.publicWorkbookMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.synthetic750kMaxRssBytes).toBe(118 * mib)
    expect(memoryGateRssBudgets.syntheticRepeatedRowMetadataMaxRssBytes).toBe(118 * mib)
    expect(memoryGateRssBudgets.syntheticRepeatedStringMaxRssBytes).toBe(118 * mib)
    expect(memoryGateRssBudgets.syntheticDuplicateSharedStringMaxRssBytes).toBe(118 * mib)
    expect(memoryGateRssBudgets.syntheticMixedRichSharedStringMaxRssBytes).toBe(118 * mib)
    expect(memoryGateRssBudgets.syntheticFormulaHeavyMaxRssBytes).toBe(118 * mib)
    expect(memoryGateRssBudgets.syntheticCachedExternalFormulaMaxRssBytes).toBe(118 * mib)
  })
})
