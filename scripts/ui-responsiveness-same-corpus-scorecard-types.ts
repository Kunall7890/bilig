import type { NumericSummary } from '../packages/benchmarks/src/stats.js'
import type { SameCorpusCaptureCorpusFingerprint } from './ui-responsiveness-same-corpus-fingerprint.ts'
import type { SameCorpusScenarioProof, sameCorpusUiRenderProofContractVersion } from './ui-responsiveness-same-corpus-proof.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export type UiResponsivenessSameCorpusProduct = 'bilig' | 'google-sheets' | 'microsoft-excel-web'

export interface UiResponsivenessSameCorpusMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: string
  readonly operationResponseMs: NumericSummary
  readonly postOperationFrameMs: NumericSummary
  readonly scrollEventResponseMs?: NumericSummary
  readonly scrollMovementPx?: NumericSummary
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface UiResponsivenessSameCorpusCase {
  readonly id: string
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleCount: number
  readonly bilig: UiResponsivenessSameCorpusMeasurement
  readonly googleSheets: UiResponsivenessSameCorpusMeasurement
  readonly microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement | undefined
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio?: number | undefined
  readonly biligToMicrosoftExcelWebP95Ratio?: number | undefined
  readonly biligToGoogleSheetsScrollEventMeanRatio?: number
  readonly biligToGoogleSheetsScrollEventP95Ratio?: number
  readonly biligToMicrosoftExcelWebScrollEventMeanRatio?: number
  readonly biligToMicrosoftExcelWebScrollEventP95Ratio?: number
  readonly tenXMeanAndP95Metric?: 'operationResponseMs' | 'scrollEventResponseMs'
  readonly scenarioProof: SameCorpusScenarioProof
  readonly tenXMeanAndP95AgainstGoogleSheets: boolean
  readonly tenXMeanAndP95AgainstMicrosoftExcelWeb?: boolean | undefined
  readonly postOperationFrameGuardrailPassed?: boolean
  readonly scrollMovementGuardrailPassed?: boolean
  readonly sourceWorkbookFingerprintGuardrailPassed?: boolean
  readonly passed: boolean
}

export interface UiResponsivenessSameCorpusProof {
  readonly captured: boolean
  readonly evidenceKind: 'same-corpus-browser-capture' | 'not-captured'
  readonly requiredProductCount: number
  readonly requiredCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly coveredCorpusCaseIds: string[]
  readonly runManifest?: UiResponsivenessSameCorpusRunManifest | undefined
  readonly limitations: string[]
  readonly cases: UiResponsivenessSameCorpusCase[]
}

export interface UiResponsivenessSameCorpusRunManifest {
  readonly artifactGenerator: 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts'
  readonly contractVersion: typeof sameCorpusUiRenderProofContractVersion
  readonly requiredProducts: readonly UiResponsivenessSameCorpusProduct[]
  readonly requiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly capturedWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly corpusCaseIds: readonly string[]
  readonly corpusFingerprints: readonly SameCorpusCaptureCorpusFingerprint[]
  readonly productSourceWorkbookFingerprints: readonly SameCorpusProductSourceWorkbookFingerprint[]
  readonly materializedCellCounts: readonly number[]
  readonly sampleCount: number
  readonly caseCount: number
  readonly strictRenderedGridProofCaseCount: number
  readonly legacyInsufficientRenderedGridProofCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly currentContractEvidenceComplete: boolean
  readonly googleSheetsTenXRequirementSatisfied: boolean
  readonly invalidReasons: readonly string[]
}

export interface SameCorpusProductSourceWorkbookFingerprint {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly method: SameCorpusCaptureCorpusVerification['method']
  readonly source: string
  readonly sourceWorkbookSha256: string | null
}

export interface SameCorpusCapture {
  readonly schemaVersion: 1
  readonly suite: 'ui-responsiveness-same-corpus-capture'
  readonly sampleCount: number
  readonly runManifest: SameCorpusCaptureRunManifest
  readonly limitations: string[]
  readonly cases: SameCorpusCaptureCase[]
}

export interface SameCorpusCaptureRunManifest {
  readonly artifactGenerator: 'scripts/capture-ui-responsiveness-same-corpus.ts'
  readonly captureToolVersion: typeof sameCorpusUiCaptureToolVersion
  readonly contractVersion: typeof sameCorpusUiRenderProofContractVersion
  readonly requiredProducts: readonly UiResponsivenessSameCorpusProduct[]
  readonly requiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly capturedWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly corpusCaseIds: readonly string[]
  readonly corpusFingerprints: readonly SameCorpusCaptureCorpusFingerprint[]
  readonly productSourceWorkbookFingerprints: readonly SameCorpusProductSourceWorkbookFingerprint[]
  readonly materializedCellCounts: readonly number[]
  readonly sampleCount: number
  readonly caseCount: number
  readonly strictRenderedGridProofCaseCount: number
  readonly legacyInsufficientRenderedGridProofCaseCount: number
  readonly currentContractEvidenceComplete: boolean
  readonly captureRunSignature: string
  readonly invalidReasons: readonly string[]
}

export interface SameCorpusCaptureCase {
  readonly id: string
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly scenarioProof: SameCorpusScenarioProof
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb?: SameCorpusCaptureMeasurement | undefined
}

export interface SameCorpusCaptureMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: string
  readonly operationResponseMsSamples: number[]
  readonly postOperationFrameMsSamples: number[]
  readonly scrollEventResponseMsSamples?: number[]
  readonly scrollMovementPxSamples?: number[]
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface SameCorpusCaptureVerifiedCell {
  readonly address: string
  readonly expected: string
  readonly actual: string
}

export interface SameCorpusCaptureCorpusVerification {
  readonly verified: boolean
  readonly method: 'bilig-benchmark-state' | 'google-sheets-xlsx-export' | 'microsoft-excel-web-source-xlsx'
  readonly sheetName: string
  readonly materializedCells: number
  readonly corpusFingerprint: SameCorpusCaptureCorpusFingerprint
  readonly sourceWorkbookSha256: string | null
  readonly checkedCells: readonly SameCorpusCaptureVerifiedCell[]
}

export const sameCorpusUiCaptureToolVersion = 'same-corpus-capture-v1'
