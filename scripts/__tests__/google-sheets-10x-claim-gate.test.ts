import { describe, expect, it } from 'vitest'

import { buildBiligDominanceScorecard } from '../gen-bilig-dominance-scorecard.ts'
import { buildGoogleSheetsTenXClaimGateReport, scorecardPermitsGoogleSheetsTenXClaim } from '../google-sheets-10x-claim-gate.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from '../ui-responsiveness-same-corpus-workloads.ts'
import { buildFixtureInput } from './bilig-dominance-scorecard.fixture.ts'

describe('Google Sheets 10x claim gate', () => {
  it('fails the current active-not-achieved dominance fixture even when diagnostic artifacts are current', () => {
    const input = buildFixtureInput()
    const report = buildGoogleSheetsTenXClaimGateReport({
      scorecard: buildBiligDominanceScorecard(input),
      uiScorecard: input.uiResponsivenessLiveBrowserScorecard,
      browserGatesSkipped: true,
    })

    expect(report.passed).toBe(false)
    expect(report.requiredWorkloadCount).toBe(requiredUiResponsivenessSameCorpusWorkloads.length)
    expect(report.tenXMeanAndP95CaseCount).toBe(0)
    expect(report.failures).toEqual(
      expect.arrayContaining([
        'Google Sheets 10x claim gate cannot run with BILIG_CI_SKIP_BROWSER=1; live browser proof must be enabled',
        'dominance scorecard goalStatus is active-not-achieved, expected achieved',
        'dominance scorecard claimPolicy.blanketTenXClaimAllowed is not true',
        'dominance scorecard completionAudit.allCriteriaPassed is not true',
        'dominance scorecard overallGoogleSheets10xStatus is not passed',
        'same-corpus UI proof is not captured',
        `same-corpus UI proof has 0/${String(requiredUiResponsivenessSameCorpusWorkloads.length)} required cases`,
        `same-corpus UI proof has 0/${String(requiredUiResponsivenessSameCorpusWorkloads.length)} 10x mean+p95 cases`,
      ]),
    )
  })

  it('passes only when scorecard, same-corpus proof, and browser gate mode are all green', () => {
    const report = buildGoogleSheetsTenXClaimGateReport({
      scorecard: passingScorecard(),
      uiScorecard: passingUiScorecard(),
      browserGatesSkipped: false,
    })

    expect(report).toEqual({
      passed: true,
      requiredWorkloadCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      tenXMeanAndP95CaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      failures: [],
    })
  })

  it('does not let a forged blanket claim boolean bypass the scorecard-level gate', () => {
    expect(
      scorecardPermitsGoogleSheetsTenXClaim({
        goalStatus: 'active-not-achieved',
        claimPolicy: {
          blanketTenXClaimAllowed: true,
          unmetRequirements: [],
        },
      }),
    ).toBe(false)
  })
})

function passingScorecard(): unknown {
  return {
    goalStatus: 'achieved',
    claimPolicy: {
      blanketTenXClaimAllowed: true,
      unmetRequirements: [],
    },
    completionAudit: {
      allCriteriaPassed: true,
      unmetRequirements: [],
      criteria: [],
    },
    overallGoogleSheets10xStatus: {
      passed: true,
      status: 'passed',
      unmetRequirements: [],
      categories: [
        { id: 'recalculation-speed', passed: true, gaps: [] },
        { id: 'structural-edit-performance', passed: true, gaps: [] },
        { id: 'large-workbook-scale', passed: true, gaps: [] },
        { id: 'ui-responsiveness', passed: true, gaps: [] },
      ],
    },
  }
}

function passingUiScorecard(): unknown {
  const requiredWorkloadCount = requiredUiResponsivenessSameCorpusWorkloads.length
  return {
    sameCorpusProof: {
      captured: true,
      requiredCaseCount: requiredWorkloadCount,
      tenXMeanAndP95CaseCount: requiredWorkloadCount,
      cases: requiredUiResponsivenessSameCorpusWorkloads.map((workload) => ({
        id: `same-corpus-${workload}`,
        workload,
        passed: true,
        tenXMeanAndP95AgainstGoogleSheets: true,
      })),
      runManifest: {
        currentContractEvidenceComplete: true,
        googleSheetsTenXRequirementSatisfied: true,
        invalidReasons: [],
        scenarioSummaryFieldCaseCount: requiredWorkloadCount,
        strictRenderedGridProofCaseCount: requiredWorkloadCount,
        visibleOperationResponseProofCaseCount: requiredWorkloadCount,
        biligAuthoritativeRenderProofCaseCount: requiredWorkloadCount,
        semanticUiProofCaseCount: requiredWorkloadCount,
        legacyInsufficientRenderedGridProofCaseCount: 0,
        tenXMeanAndP95CaseCount: requiredWorkloadCount,
      },
    },
  }
}
