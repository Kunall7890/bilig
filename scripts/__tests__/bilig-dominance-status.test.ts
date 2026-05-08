import { describe, expect, it } from 'vitest'

import { buildBiligDominanceStatus, formatBiligDominanceStatusPathForMessage } from '../bilig-dominance-status.ts'
import { buildSameCorpusProof, type SameCorpusCapture } from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import type { PublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import { buildFixtureInput } from './bilig-dominance-scorecard.fixture.ts'

describe('bilig dominance status', () => {
  it('exposes actionable same-corpus UI proof setup commands', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })

    expect(status.uiSameCorpus).toMatchObject({
      captured: false,
      evidenceKind: 'not-captured',
      requiredWorkloads: ['visible-scroll-response'],
      missingRequiredWorkloads: ['visible-scroll-response'],
      googleSheetsUrl: null,
      googleSheetsUrlEnvVar: 'BILIG_UI_SAME_CORPUS_GOOGLE_SHEETS_URL',
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
      fixture: {
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        localXlsxPath: 'packages/benchmarks/baselines/ui-same-corpus/wide-mixed-250k.xlsx',
      },
      nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
      nextScorecardGenerateCommand: 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json',
      nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
    })
    expect(status.uiSameCorpus.fixture.microsoftExcelWebUrl).toContain('view.officeapps.live.com/op/view.aspx')
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('--google-sheets-url')
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('.cache/ui-responsiveness/same-corpus-capture.json')
    expect(status.uiSameCorpus.nextGoogleSheetsUploadInstruction).toContain('share it to anyone with the link')
  })

  it('fills same-corpus UI proof commands when the Google Sheets URL is known', () => {
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit'
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusGoogleSheetsUrl: googleSheetsUrl,
    })

    expect(status.uiSameCorpus.googleSheetsUrl).toBe(googleSheetsUrl)
    expect(status.uiSameCorpus.missingInputs).toEqual([])
    expect(status.uiSameCorpus.nextPreflightCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<google-sheets-url>')
  })

  it('keeps asking for the Google Sheets URL when captured same-corpus proof is not 10x', () => {
    const fixtureInput = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input: {
        ...fixtureInput,
        uiResponsivenessLiveBrowserScorecard: {
          ...fixtureInput.uiResponsivenessLiveBrowserScorecard,
          sameCorpusProof: buildSameCorpusProof(failingSameCorpusCapture()),
        },
      },
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })

    expect(status.uiSameCorpus).toMatchObject({
      captured: true,
      evidenceKind: 'same-corpus-browser-capture',
      requiredCaseCount: 1,
      tenXMeanAndP95CaseCount: 0,
      tenXRequirementSatisfied: false,
      missingRequiredWorkloads: [],
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
    })
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('<google-sheets-url>')
  })

  it('surfaces financial workbook corpus blockers in dominance status', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: {
        targetWorkbookCount: 5_000,
        sourceCount: 9_824,
        cachedArtifactCount: 0,
        recordedManifestArtifactCount: 0,
        recordedNonPassingCaseCount: 0,
      },
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: true,
      stopMarkerPath: '.agent-coordination/stop.md',
    })

    expect(status.publicWorkbookCorpus).toMatchObject({
      financialWorkbookTargetCount: 5_000,
      financialSourceCount: 9_824,
      financialCachedArtifactCount: 0,
      recordedFinancialManifestArtifactCount: 0,
    })
    expect(status.importExportBlockers).toEqual(
      expect.arrayContaining([
        'financial/accounting corpus cached artifacts below target: 0/5000',
        'financial/accounting corpus recorded verification cases below target: 0/5000',
      ]),
    )
    expect(status.goalStatus).toBe('active-not-achieved')
  })

  it('formats repo-local status paths without exposing the checkout root', () => {
    expect(formatBiligDominanceStatusPathForMessage('/repo/.cache/public-workbook-corpus/manifest.json', '/repo')).toBe(
      '.cache/public-workbook-corpus/manifest.json',
    )
    expect(formatBiligDominanceStatusPathForMessage('/tmp/public-workbook-corpus/manifest.json', '/repo')).toBe(
      '/tmp/public-workbook-corpus/manifest.json',
    )
  })
})

function completePublicWorkbookCorpusStatus(): PublicWorkbookCorpusStatus {
  return {
    targetWorkbookCount: 10_000,
    sourceCount: 10_000,
    cachedArtifactCount: 10_000,
    scorecardCaseCount: 10_000,
    checkpointCaseCount: 10_000,
    recordedManifestArtifactCount: 10_000,
    missingManifestArtifactCount: 0,
    staleRecordedVerificationCount: 0,
    recordedPassedCaseCount: 10_000,
    recordedUnsupportedCaseCount: 0,
    recordedFailedCaseCount: 0,
    recordedErrorCaseCount: 0,
    recordedCoversManifest: true,
    recordedAllCasesPassed: true,
    missingManifestArtifactSample: [],
    staleRecordedVerificationSample: [],
    nextMissingVerificationCommand: null,
    nextMissingVerificationPlanCommand: null,
    nextStaleVerificationCommand: null,
    nextStaleVerificationPlanCommand: null,
    scorecardCoversManifest: true,
    targetComplete: true,
    gaps: [],
  }
}

function completeFinancialCorpusStatus() {
  return {
    targetWorkbookCount: 5_000,
    sourceCount: 5_000,
    cachedArtifactCount: 5_000,
    recordedManifestArtifactCount: 5_000,
    recordedNonPassingCaseCount: 0,
  }
}

function failingSameCorpusCapture(): SameCorpusCapture {
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-capture',
    sampleCount: 3,
    limitations: [],
    cases: [
      {
        id: 'same-corpus-wide-mixed-250k-visible-scroll-response',
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        workload: 'visible-scroll-response',
        bilig: {
          product: 'bilig',
          source: 'e2e/tests/web-shell-scroll-performance.pw.ts',
          operationResponseMsSamples: [200, 200, 200],
          postOperationFrameMsSamples: [12, 12, 12],
          corpusVerification: sameCorpusVerification('bilig-benchmark-state'),
          limitations: [],
        },
        googleSheets: {
          product: 'google-sheets',
          source: 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit',
          operationResponseMsSamples: [100, 100, 100],
          postOperationFrameMsSamples: [16, 16, 16],
          corpusVerification: sameCorpusVerification('google-sheets-xlsx-export'),
          limitations: [],
        },
        microsoftExcelWeb: {
          product: 'microsoft-excel-web',
          source: 'https://view.officeapps.live.com/op/view.aspx?src=sameCorpusWorkbook',
          operationResponseMsSamples: [100, 100, 100],
          postOperationFrameMsSamples: [16, 16, 16],
          corpusVerification: sameCorpusVerification('microsoft-excel-web-source-xlsx'),
          limitations: [],
        },
      },
    ],
  }
}

function sameCorpusVerification(method: SameCorpusCapture['cases'][number]['bilig']['corpusVerification']['method']) {
  return {
    verified: true,
    method,
    sheetName: 'WideGrid',
    materializedCells: 250_000,
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
    ],
  }
}
