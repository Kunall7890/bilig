import type { NumericSummary } from '../packages/benchmarks/src/stats.js'
import type { SameCorpusCaptureCorpusFingerprint } from './ui-responsiveness-same-corpus-fingerprint.ts'
import type {
  SameCorpusPixelGridProof,
  SameCorpusScenarioProof,
  SameCorpusSemanticUiProof,
  SameCorpusScreenshotProof,
  sameCorpusUiRenderProofContractVersion,
} from './ui-responsiveness-same-corpus-proof.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export type UiResponsivenessSameCorpusProduct = 'bilig' | 'google-sheets' | 'microsoft-excel-web'
export type SameCorpusOperationResponseProof = 'load-to-ready' | 'visible-non-scroll-response' | 'visible-scroll-movement'

export interface UiResponsivenessSameCorpusMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: string
  readonly operationResponseMs: NumericSummary
  readonly operationResponseProofs: readonly SameCorpusOperationResponseProof[]
  readonly authoritativeRenderProofMs?: NumericSummary
  readonly postOperationFrameMs: NumericSummary
  readonly scrollEventResponseMs?: NumericSummary
  readonly scrollMovementPx?: NumericSummary
  readonly biligRuntimeProof?: SameCorpusBiligRuntimeProof
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface SameCorpusScenarioCaseFields {
  readonly biligMeanMs: number
  readonly biligP95Ms: number
  readonly googleMeanMs: number
  readonly googleP95Ms: number
  readonly microsoftExcelWebMeanMs?: number
  readonly microsoftExcelWebP95Ms?: number
  /** Google Sheets time divided by Bilig time. Values >= 10 mean Bilig is at least 10x faster. */
  readonly meanRatio: number
  /** Google Sheets p95 divided by Bilig p95. Values >= 10 mean Bilig is at least 10x faster at p95. */
  readonly p95Ratio: number
  /** Microsoft Excel Web time divided by Bilig time. Values >= 10 mean Bilig is at least 10x faster. */
  readonly microsoftExcelWebMeanRatio?: number
  /** Microsoft Excel Web p95 divided by Bilig p95. Values >= 10 mean Bilig is at least 10x faster at p95. */
  readonly microsoftExcelWebP95Ratio?: number
  readonly screenshotProof: SameCorpusScreenshotProof
  readonly pixelGridProof: SameCorpusPixelGridProof
  readonly semanticUiProof: SameCorpusSemanticUiProof
}

export interface UiResponsivenessSameCorpusCase extends SameCorpusScenarioCaseFields {
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
  readonly biligRuntimeProofGuardrailPassed?: boolean
  readonly scrollMovementGuardrailPassed?: boolean
  readonly sourceWorkbookFingerprintGuardrailPassed?: boolean
  readonly operationResponseProofGuardrailPassed?: boolean
  readonly authoritativeRenderProofGuardrailPassed?: boolean
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
  readonly biligProductionRuntimeProofCaseCount: number
  readonly sampleCount: number
  readonly caseCount: number
  readonly scenarioSummaryFieldCaseCount: number
  readonly strictRenderedGridProofCaseCount: number
  readonly visibleOperationResponseProofCaseCount: number
  readonly biligAuthoritativeRenderProofCaseCount: number
  readonly semanticUiProofCaseCount: number
  readonly requiredMutationTargetProofCaseCount: number
  readonly mutationTargetProofCaseCount: number
  readonly requiredMutationTargetProofSampleCount: number
  readonly mutationTargetProofSampleCount: number
  readonly legacyInsufficientRenderedGridProofCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly currentContractEvidenceComplete: boolean
  readonly googleSheetsTenXRequirementSatisfied: boolean
  readonly captureRunSignature: string | null
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
  readonly biligProductionRuntimeProofCaseCount: number
  readonly sampleCount: number
  readonly caseCount: number
  readonly scenarioSummaryFieldCaseCount: number
  readonly strictRenderedGridProofCaseCount: number
  readonly visibleOperationResponseProofCaseCount: number
  readonly biligAuthoritativeRenderProofCaseCount: number
  readonly semanticUiProofCaseCount: number
  readonly requiredMutationTargetProofCaseCount: number
  readonly mutationTargetProofCaseCount: number
  readonly requiredMutationTargetProofSampleCount: number
  readonly mutationTargetProofSampleCount: number
  readonly legacyInsufficientRenderedGridProofCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly currentContractEvidenceComplete: boolean
  readonly googleSheetsTenXRequirementSatisfied: boolean
  readonly captureRunSignature: string
  readonly invalidReasons: readonly string[]
}

export interface SameCorpusCaptureCase extends SameCorpusScenarioCaseFields {
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
  readonly operationResponseProofs: SameCorpusOperationResponseProof[]
  readonly authoritativeRenderProofMsSamples?: number[]
  readonly postOperationFrameMsSamples: number[]
  readonly scrollEventResponseMsSamples?: number[]
  readonly scrollMovementPxSamples?: number[]
  readonly biligRuntimeProof?: SameCorpusBiligRuntimeProof
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface SameCorpusBiligRuntimeProof {
  readonly product: 'bilig'
  readonly source: string
  readonly verificationMethod: 'window.__biligRuntimeBuild'
  readonly requiredBuildKind: 'production'
  readonly actualBuildKind: 'development' | 'production' | 'unknown'
  readonly mode: string
  readonly dev: boolean
  readonly prod: boolean
  readonly remoteSyncEnabled: boolean | null
  readonly entryRoute: string | null
  readonly sampleCount: number
  readonly verified: boolean
  readonly samples: readonly SameCorpusBiligRuntimeProofSample[]
}

export interface SameCorpusBiligRuntimeProofSample {
  readonly sampleIndex: number
  readonly present: boolean
  readonly app: string | null
  readonly buildKind: 'development' | 'production' | 'unknown'
  readonly mode: string
  readonly dev: boolean
  readonly prod: boolean
  readonly remoteSyncEnabled: boolean | null
  readonly entryRoute: string | null
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

export const sameCorpusUiCaptureToolVersion = 'same-corpus-capture-v5'
