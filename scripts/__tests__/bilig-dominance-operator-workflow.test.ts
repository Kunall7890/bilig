import { describe, expect, it } from 'vitest'

import { rootDir } from '../bilig-dominance-scorecard-input.ts'
import { loadOperatorWorkflowEvidence, operatorWorkflowGaps } from '../bilig-dominance-operator-workflow.ts'

describe('bilig dominance operator workflow evidence', () => {
  it('loads repo wiring for dominance checks and blanket-claim policy', () => {
    const evidence = loadOperatorWorkflowEvidence(rootDir)

    expect(operatorWorkflowGaps(evidence)).toEqual([])
    expect(evidence).toMatchObject({
      dominanceGenerateScriptPresent: true,
      dominanceCheckScriptPresent: true,
      dominanceAuditCheckScriptPresent: true,
      publicClaimsCheckScriptPresent: true,
      runCiDominanceCheckPresent: true,
      runCiDominanceAuditCheckPresent: true,
      runCiPublicClaimsCheckPresent: true,
      generatedSourceChecksSerialized: true,
      blanketClaimPolicyCoupledToCompletionAudit: true,
      promptArtifactAuditCoupledToLiveStatus: true,
    })
  })
})
