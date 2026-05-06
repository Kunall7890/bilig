import { describe, expect, it } from 'vitest'

import { buildAuditabilityScorecard, validateAuditabilityScorecard } from '../gen-auditability-scorecard.ts'

describe('auditability scorecard', () => {
  it('generates a checked artifact from executable auditability controls', async () => {
    const scorecard = await buildAuditabilityScorecard('2026-05-06T10:00:00.000Z')

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'auditability-posture',
      generatedAt: '2026-05-06T10:00:00.000Z',
      summary: {
        allRequiredControlsPassed: true,
        previewApplyParityPassed: true,
        applyUndoRoundTripPassed: true,
        authoritativeApplyGuardPassed: true,
        historyRevertRedoPassed: true,
        headedBrowserRevertFlowPassed: true,
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
    })
    expect(scorecard.source.externalAuditabilityComparisonArtifact).toBe(
      'packages/benchmarks/baselines/auditability-external-sheets-excel-comparison.json',
    )
    expect(scorecard.controls.map((control) => control.id)).toEqual([
      'agent-preview-apply-parity',
      'agent-apply-undo-roundtrip',
      'authoritative-agent-apply-fails-closed',
      'workbook-history-revert-redo-state',
      'headed-browser-change-review-revert-flow',
      'external-sheets-excel-auditability-comparison',
    ])
    expect(scorecard.controls.every((control) => control.required && control.passed)).toBe(true)
    expect(scorecard.controls.find((control) => control.id === 'external-sheets-excel-auditability-comparison')).toMatchObject({
      category: 'external-comparison',
      coveredControls: [
        'external.googleSheetsAuditabilityDocs',
        'external.microsoftExcelAuditabilityDocs',
        'external.sheetsExcelAuditabilityComparison',
      ],
    })
    expect(scorecard.summary.coveredControls).toEqual([
      'agent.previewDiffParity',
      'agent.authoritativePreviewMismatchFailsClosed',
      'agent.baseRevisionStaleApplyFailsClosed',
      'agent.applyCapturesUndoBundle',
      'agent.undoBundleRestoresSnapshot',
      'history.revertRedoStack',
      'history.revertLinkage',
      'headedBrowser.previewApplyRevertFlow',
      'external.googleSheetsAuditabilityDocs',
      'external.microsoftExcelAuditabilityDocs',
      'external.sheetsExcelAuditabilityComparison',
    ])
    expect(scorecard.summary.uncoveredControls).toEqual([])
  })

  it('rejects stale artifacts missing required auditability controls', async () => {
    const scorecard = await buildAuditabilityScorecard('2026-05-06T10:00:00.000Z')
    const staleScorecard = {
      ...scorecard,
      controls: scorecard.controls.filter((control) => control.id !== 'authoritative-agent-apply-fails-closed'),
    }

    expect(() => validateAuditabilityScorecard(staleScorecard)).toThrow(
      'Auditability scorecard is missing required control: authoritative-agent-apply-fails-closed',
    )
  })
})
