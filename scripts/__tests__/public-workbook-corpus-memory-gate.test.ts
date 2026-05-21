import { describe, expect, it } from 'vitest'

import { memoryGateRssBudgets } from '../public-workbook-corpus-memory-gate.ts'

const mib = 1024 * 1024

describe('public workbook corpus memory gate budgets', () => {
  it('keeps every large-workbook gate at the V4 hard target', () => {
    expect(memoryGateRssBudgets.publicWorkbookMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.synthetic750kMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.syntheticRepeatedRowMetadataMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.syntheticRepeatedStringMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.syntheticDuplicateSharedStringMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.syntheticMixedRichSharedStringMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.syntheticFormulaHeavyMaxRssBytes).toBe(112 * mib)
    expect(memoryGateRssBudgets.syntheticCachedExternalFormulaMaxRssBytes).toBe(112 * mib)
  })
})
