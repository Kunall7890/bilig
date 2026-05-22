import { describe, expect, it } from 'vitest'
import { summarizeWorkbookAgentVerificationStatus } from './workbook-agent-verification-status.js'

describe('summarizeWorkbookAgentVerificationStatus', () => {
  it('does not treat skipped formula and invariant audits as complete verification', () => {
    const status = summarizeWorkbookAgentVerificationStatus({
      renderedReadback: [
        {
          requested: true,
          matched: true,
        },
      ],
      recalculationStatus: {
        upToDate: true,
      },
      formulaIssues: null,
      invariants: null,
      requireTargetRange: true,
      targetRangeCount: 1,
    })

    expect(status.verificationComplete).toBe(false)
    expect(status.renderedComplete).toBe(true)
    expect(status.recalculationComplete).toBe(true)
    expect(status.formulaComplete).toBe(false)
    expect(status.invariantsComplete).toBe(false)
    expect(status.missingChecks).toEqual(['formulaIssues', 'invariants'])
  })

  it('does not treat stale recalculation as complete verification', () => {
    const status = summarizeWorkbookAgentVerificationStatus({
      renderedReadback: [
        {
          requested: true,
          matched: true,
        },
      ],
      recalculationStatus: {
        upToDate: false,
      },
      formulaIssues: {
        summary: {
          actionableIssueCount: 0,
        },
      },
      invariants: {
        summary: {
          ok: true,
        },
      },
      requireTargetRange: true,
      targetRangeCount: 1,
    })

    expect(status.verificationComplete).toBe(false)
    expect(status.recalculationComplete).toBe(false)
    expect(status.missingChecks).toEqual(['recalculationStale'])
  })
})
