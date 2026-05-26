import type { NumericSummary } from '../packages/benchmarks/src/stats.js'
import {
  arrayField,
  asObject,
  booleanField,
  literalField,
  numberField,
  objectField,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'
import type {
  SameCorpusCapture,
  SameCorpusBiligRuntimeProof,
  SameCorpusCaptureCorpusFingerprint,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  SameCorpusCaptureRunManifest,
  SameCorpusProductSourceWorkbookFingerprint,
  SameCorpusCaptureVerifiedCell,
  SameCorpusMutationTargetProofProductSummary,
  SameCorpusMutationTargetProofSampleSummary,
  SameCorpusScenarioCaseFields,
  UiResponsivenessLiveBrowserCase,
  UiResponsivenessLiveBrowserScorecard,
  UiResponsivenessLiveBrowserVendor,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
  UiResponsivenessSameCorpusProof,
  UiResponsivenessSameCorpusRunManifest,
  UiResponsivenessSameCorpusWorkload,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type {
  SameCorpusPixelGridProof,
  SameCorpusProductPixelGridProof,
  SameCorpusScenarioProof,
  SameCorpusProductSemanticUiProof,
  SameCorpusProductSemanticUiProofVerdict,
  SameCorpusSemanticUiProof,
  SameCorpusScreenshotProof,
  SameCorpusMutationTargetProof,
  SameCorpusMutationTargetReadback,
} from './ui-responsiveness-same-corpus-proof.ts'
import { sameCorpusUiRenderProofContractVersion, validateSameCorpusProductSemanticUiProof } from './ui-responsiveness-same-corpus-proof.ts'
import {
  requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount,
  requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount,
} from './ui-responsiveness-same-corpus-mutation-target-proof-summary.ts'
import { sameCorpusUiCaptureToolVersion } from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import type { SameCorpusOperationResponseProof } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { isUiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function parseUiResponsivenessLiveBrowserScorecard(value: Record<string, unknown>): UiResponsivenessLiveBrowserScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const benchmarkViewport = objectField(benchmark, 'viewport')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'ui-responsiveness-live-browser-timing'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-public-browser-playwright'),
      browserEngine: literalField(source, 'browserEngine', 'chromium'),
      measuredOperation: literalField(source, 'measuredOperation', 'public-workbook-load-and-viewport-scroll'),
    },
    benchmark: {
      sampleCount: numberField(benchmark, 'sampleCount'),
      viewport: {
        width: numberField(benchmarkViewport, 'width'),
        height: numberField(benchmarkViewport, 'height'),
      },
      samplingOrder: literalField(benchmark, 'samplingOrder', 'google-sheets-then-microsoft-excel-web'),
    },
    summary: {
      directBrowserTimingCaptured: booleanField(summary, 'directBrowserTimingCaptured'),
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredVendorCount: numberField(summary, 'requiredVendorCount'),
      capturedVendors: stringArrayField(summary, 'capturedVendors').map(parseVendor),
      limitations: stringArrayField(summary, 'limitations'),
    },
    cases: arrayField(value, 'cases').map(parseBrowserCase),
    sameCorpusProof: parseSameCorpusProof(objectField(value, 'sameCorpusProof')),
  }
}

export function parseSameCorpusCapture(value: Record<string, unknown>): SameCorpusCapture {
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'ui-responsiveness-same-corpus-capture'),
    sampleCount: numberField(value, 'sampleCount'),
    runManifest: parseSameCorpusCaptureRunManifest(objectField(value, 'runManifest')),
    limitations: stringArrayField(value, 'limitations'),
    cases: arrayField(value, 'cases').map(parseSameCorpusCaptureCase),
  }
}

function parseBrowserCase(value: unknown): UiResponsivenessLiveBrowserCase {
  const record = asObject(value, 'UI responsiveness live browser case')
  return {
    id: stringField(record, 'id'),
    vendor: parseVendor(stringField(record, 'vendor')),
    product: stringField(record, 'product'),
    sourceUrl: stringField(record, 'sourceUrl'),
    finalUrl: stringField(record, 'finalUrl'),
    title: stringField(record, 'title'),
    accessMode: parseAccessMode(stringField(record, 'accessMode')),
    workload: literalField(record, 'workload', 'open-public-workbook-and-scroll-viewport'),
    sampleCount: numberField(record, 'sampleCount'),
    loadToReadyMs: parseNumericSummary(objectField(record, 'loadToReadyMs')),
    scrollResponseMs: parseNumericSummary(objectField(record, 'scrollResponseMs')),
    postScrollFrameMs: parseNumericSummary(objectField(record, 'postScrollFrameMs')),
    passed: booleanField(record, 'passed'),
    limitations: stringArrayField(record, 'limitations'),
  }
}

function parseSameCorpusProof(value: Record<string, unknown>): UiResponsivenessSameCorpusProof {
  return {
    captured: booleanField(value, 'captured'),
    evidenceKind: parseSameCorpusEvidenceKind(stringField(value, 'evidenceKind')),
    requiredProductCount: numberField(value, 'requiredProductCount'),
    requiredCaseCount: numberField(value, 'requiredCaseCount'),
    tenXMeanAndP95CaseCount: numberField(value, 'tenXMeanAndP95CaseCount'),
    coveredCorpusCaseIds: stringArrayField(value, 'coveredCorpusCaseIds'),
    ...(Object.hasOwn(value, 'runManifest') ? { runManifest: parseSameCorpusRunManifest(objectField(value, 'runManifest')) } : {}),
    limitations: stringArrayField(value, 'limitations'),
    cases: arrayField(value, 'cases').map(parseSameCorpusCase),
  }
}

function parseSameCorpusRunManifest(value: Record<string, unknown>): UiResponsivenessSameCorpusRunManifest {
  return {
    artifactGenerator: literalField(value, 'artifactGenerator', 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts'),
    contractVersion: literalField(value, 'contractVersion', sameCorpusUiRenderProofContractVersion),
    requiredProducts: stringArrayField(value, 'requiredProducts').map(parseSameCorpusProduct),
    requiredWorkloads: stringArrayField(value, 'requiredWorkloads').map(parseSameCorpusWorkload),
    capturedWorkloads: stringArrayField(value, 'capturedWorkloads').map(parseSameCorpusWorkload),
    corpusCaseIds: stringArrayField(value, 'corpusCaseIds'),
    corpusFingerprints: arrayField(value, 'corpusFingerprints').map(parseSameCorpusCorpusFingerprint),
    productSourceWorkbookFingerprints: arrayField(value, 'productSourceWorkbookFingerprints').map(
      parseSameCorpusProductSourceWorkbookFingerprint,
    ),
    materializedCellCounts: arrayField(value, 'materializedCellCounts').map((entry) => {
      if (typeof entry !== 'number' || !Number.isInteger(entry) || entry <= 0) {
        throw new Error('Expected UI responsiveness same-corpus materialized cell counts to contain positive integers')
      }
      return entry
    }),
    biligProductionRuntimeProofCaseCount: numberField(value, 'biligProductionRuntimeProofCaseCount'),
    sampleCount: numberField(value, 'sampleCount'),
    caseCount: numberField(value, 'caseCount'),
    scenarioSummaryFieldCaseCount: numberField(value, 'scenarioSummaryFieldCaseCount'),
    strictRenderedGridProofCaseCount: numberField(value, 'strictRenderedGridProofCaseCount'),
    visibleOperationResponseProofCaseCount: numberField(value, 'visibleOperationResponseProofCaseCount'),
    biligAuthoritativeRenderProofCaseCount: optionalNumberField(value, 'biligAuthoritativeRenderProofCaseCount') ?? 0,
    semanticUiProofCaseCount: optionalNumberField(value, 'semanticUiProofCaseCount') ?? 0,
    requiredMutationTargetProofCaseCount:
      optionalNumberField(value, 'requiredMutationTargetProofCaseCount') ??
      requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount(),
    mutationTargetProofCaseCount: optionalNumberField(value, 'mutationTargetProofCaseCount') ?? 0,
    requiredMutationTargetProofSampleCount:
      optionalNumberField(value, 'requiredMutationTargetProofSampleCount') ??
      requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount(numberField(value, 'sampleCount')),
    mutationTargetProofSampleCount: optionalNumberField(value, 'mutationTargetProofSampleCount') ?? 0,
    mutationTargetProofProductSummaries: Object.hasOwn(value, 'mutationTargetProofProductSummaries')
      ? arrayField(value, 'mutationTargetProofProductSummaries').map(parseSameCorpusMutationTargetProofProductSummary)
      : [],
    legacyInsufficientRenderedGridProofCaseCount: numberField(value, 'legacyInsufficientRenderedGridProofCaseCount'),
    tenXMeanAndP95CaseCount: numberField(value, 'tenXMeanAndP95CaseCount'),
    currentContractEvidenceComplete: booleanField(value, 'currentContractEvidenceComplete'),
    googleSheetsTenXRequirementSatisfied: booleanField(value, 'googleSheetsTenXRequirementSatisfied'),
    captureRunSignature: nullableStringField(value, 'captureRunSignature'),
    invalidReasons: stringArrayField(value, 'invalidReasons'),
  }
}

function parseSameCorpusCaptureRunManifest(value: Record<string, unknown>): SameCorpusCaptureRunManifest {
  return {
    artifactGenerator: literalField(value, 'artifactGenerator', 'scripts/capture-ui-responsiveness-same-corpus.ts'),
    captureToolVersion: literalField(value, 'captureToolVersion', sameCorpusUiCaptureToolVersion),
    contractVersion: literalField(value, 'contractVersion', sameCorpusUiRenderProofContractVersion),
    requiredProducts: stringArrayField(value, 'requiredProducts').map(parseSameCorpusProduct),
    requiredWorkloads: stringArrayField(value, 'requiredWorkloads').map(parseSameCorpusWorkload),
    capturedWorkloads: stringArrayField(value, 'capturedWorkloads').map(parseSameCorpusWorkload),
    corpusCaseIds: stringArrayField(value, 'corpusCaseIds'),
    corpusFingerprints: arrayField(value, 'corpusFingerprints').map(parseSameCorpusCorpusFingerprint),
    productSourceWorkbookFingerprints: arrayField(value, 'productSourceWorkbookFingerprints').map(
      parseSameCorpusProductSourceWorkbookFingerprint,
    ),
    materializedCellCounts: arrayField(value, 'materializedCellCounts').map((entry) => {
      if (typeof entry !== 'number' || !Number.isInteger(entry) || entry <= 0) {
        throw new Error('Expected UI responsiveness same-corpus capture materialized cell counts to contain positive integers')
      }
      return entry
    }),
    biligProductionRuntimeProofCaseCount: numberField(value, 'biligProductionRuntimeProofCaseCount'),
    sampleCount: numberField(value, 'sampleCount'),
    caseCount: numberField(value, 'caseCount'),
    scenarioSummaryFieldCaseCount: numberField(value, 'scenarioSummaryFieldCaseCount'),
    strictRenderedGridProofCaseCount: numberField(value, 'strictRenderedGridProofCaseCount'),
    visibleOperationResponseProofCaseCount: numberField(value, 'visibleOperationResponseProofCaseCount'),
    biligAuthoritativeRenderProofCaseCount: optionalNumberField(value, 'biligAuthoritativeRenderProofCaseCount') ?? 0,
    semanticUiProofCaseCount: optionalNumberField(value, 'semanticUiProofCaseCount') ?? 0,
    requiredMutationTargetProofCaseCount:
      optionalNumberField(value, 'requiredMutationTargetProofCaseCount') ??
      requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount(),
    mutationTargetProofCaseCount: optionalNumberField(value, 'mutationTargetProofCaseCount') ?? 0,
    requiredMutationTargetProofSampleCount:
      optionalNumberField(value, 'requiredMutationTargetProofSampleCount') ??
      requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount(numberField(value, 'sampleCount')),
    mutationTargetProofSampleCount: optionalNumberField(value, 'mutationTargetProofSampleCount') ?? 0,
    mutationTargetProofProductSummaries: Object.hasOwn(value, 'mutationTargetProofProductSummaries')
      ? arrayField(value, 'mutationTargetProofProductSummaries').map(parseSameCorpusMutationTargetProofProductSummary)
      : [],
    legacyInsufficientRenderedGridProofCaseCount: numberField(value, 'legacyInsufficientRenderedGridProofCaseCount'),
    tenXMeanAndP95CaseCount: numberField(value, 'tenXMeanAndP95CaseCount'),
    currentContractEvidenceComplete: booleanField(value, 'currentContractEvidenceComplete'),
    googleSheetsTenXRequirementSatisfied: booleanField(value, 'googleSheetsTenXRequirementSatisfied'),
    captureRunSignature: stringField(value, 'captureRunSignature'),
    invalidReasons: stringArrayField(value, 'invalidReasons'),
  }
}

function parseSameCorpusCase(value: unknown): UiResponsivenessSameCorpusCase {
  const record = asObject(value, 'UI responsiveness same-corpus case')
  const microsoftExcelWeb = Object.hasOwn(record, 'microsoftExcelWeb')
    ? parseSameCorpusMeasurement(objectField(record, 'microsoftExcelWeb'))
    : undefined
  const biligToMicrosoftExcelWebMeanRatio = optionalNumberField(record, 'biligToMicrosoftExcelWebMeanRatio')
  const biligToMicrosoftExcelWebP95Ratio = optionalNumberField(record, 'biligToMicrosoftExcelWebP95Ratio')
  const biligToGoogleSheetsScrollEventMeanRatio = optionalNumberField(record, 'biligToGoogleSheetsScrollEventMeanRatio')
  const biligToGoogleSheetsScrollEventP95Ratio = optionalNumberField(record, 'biligToGoogleSheetsScrollEventP95Ratio')
  const biligToMicrosoftExcelWebScrollEventMeanRatio = optionalNumberField(record, 'biligToMicrosoftExcelWebScrollEventMeanRatio')
  const biligToMicrosoftExcelWebScrollEventP95Ratio = optionalNumberField(record, 'biligToMicrosoftExcelWebScrollEventP95Ratio')
  const tenXMeanAndP95Metric = optionalSameCorpusTenXMetric(record, 'tenXMeanAndP95Metric')
  const postOperationFrameGuardrailPassed = optionalBooleanField(record, 'postOperationFrameGuardrailPassed')
  const operationResponseProofGuardrailPassed = optionalBooleanField(record, 'operationResponseProofGuardrailPassed')
  const authoritativeRenderProofGuardrailPassed = optionalBooleanField(record, 'authoritativeRenderProofGuardrailPassed')
  const biligRuntimeProofGuardrailPassed = optionalBooleanField(record, 'biligRuntimeProofGuardrailPassed')
  const scrollMovementGuardrailPassed = optionalBooleanField(record, 'scrollMovementGuardrailPassed')
  const sourceWorkbookFingerprintGuardrailPassed = optionalBooleanField(record, 'sourceWorkbookFingerprintGuardrailPassed')
  const scenarioCaseFields = parseSameCorpusScenarioCaseFields(record)
  return {
    id: stringField(record, 'id'),
    corpusCaseId: stringField(record, 'corpusCaseId'),
    materializedCells: numberField(record, 'materializedCells'),
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    sampleCount: numberField(record, 'sampleCount'),
    bilig: parseSameCorpusMeasurement(objectField(record, 'bilig')),
    googleSheets: parseSameCorpusMeasurement(objectField(record, 'googleSheets')),
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
    biligToGoogleSheetsMeanRatio: numberField(record, 'biligToGoogleSheetsMeanRatio'),
    biligToGoogleSheetsP95Ratio: numberField(record, 'biligToGoogleSheetsP95Ratio'),
    ...(biligToMicrosoftExcelWebMeanRatio !== undefined ? { biligToMicrosoftExcelWebMeanRatio } : {}),
    ...(biligToMicrosoftExcelWebP95Ratio !== undefined ? { biligToMicrosoftExcelWebP95Ratio } : {}),
    ...(biligToGoogleSheetsScrollEventMeanRatio !== undefined ? { biligToGoogleSheetsScrollEventMeanRatio } : {}),
    ...(biligToGoogleSheetsScrollEventP95Ratio !== undefined ? { biligToGoogleSheetsScrollEventP95Ratio } : {}),
    ...(biligToMicrosoftExcelWebScrollEventMeanRatio !== undefined ? { biligToMicrosoftExcelWebScrollEventMeanRatio } : {}),
    ...(biligToMicrosoftExcelWebScrollEventP95Ratio !== undefined ? { biligToMicrosoftExcelWebScrollEventP95Ratio } : {}),
    ...(tenXMeanAndP95Metric ? { tenXMeanAndP95Metric } : {}),
    ...scenarioCaseFields,
    scenarioProof: parseSameCorpusScenarioProof(objectField(record, 'scenarioProof')),
    tenXMeanAndP95AgainstGoogleSheets: booleanField(record, 'tenXMeanAndP95AgainstGoogleSheets'),
    ...(Object.hasOwn(record, 'tenXMeanAndP95AgainstMicrosoftExcelWeb')
      ? { tenXMeanAndP95AgainstMicrosoftExcelWeb: booleanField(record, 'tenXMeanAndP95AgainstMicrosoftExcelWeb') }
      : {}),
    ...(postOperationFrameGuardrailPassed !== undefined ? { postOperationFrameGuardrailPassed } : {}),
    ...(operationResponseProofGuardrailPassed !== undefined ? { operationResponseProofGuardrailPassed } : {}),
    ...(authoritativeRenderProofGuardrailPassed !== undefined ? { authoritativeRenderProofGuardrailPassed } : {}),
    ...(biligRuntimeProofGuardrailPassed !== undefined ? { biligRuntimeProofGuardrailPassed } : {}),
    ...(scrollMovementGuardrailPassed !== undefined ? { scrollMovementGuardrailPassed } : {}),
    ...(sourceWorkbookFingerprintGuardrailPassed !== undefined ? { sourceWorkbookFingerprintGuardrailPassed } : {}),
    passed: booleanField(record, 'passed'),
  }
}

function parseSameCorpusMeasurement(value: Record<string, unknown>): UiResponsivenessSameCorpusMeasurement {
  const scrollEventResponseMs = optionalNumericSummary(value, 'scrollEventResponseMs')
  const scrollMovementPx = optionalNumericSummary(value, 'scrollMovementPx')
  const authoritativeRenderProofMs = optionalNumericSummary(value, 'authoritativeRenderProofMs')
  const biligRuntimeProof = Object.hasOwn(value, 'biligRuntimeProof')
    ? parseBiligRuntimeProof(objectField(value, 'biligRuntimeProof'))
    : undefined
  return {
    product: parseSameCorpusProduct(stringField(value, 'product')),
    source: stringField(value, 'source'),
    operationResponseMs: parseNumericSummary(objectField(value, 'operationResponseMs')),
    operationResponseProofs: stringArrayField(value, 'operationResponseProofs').map(parseSameCorpusOperationResponseProof),
    ...(authoritativeRenderProofMs ? { authoritativeRenderProofMs } : {}),
    postOperationFrameMs: parseNumericSummary(objectField(value, 'postOperationFrameMs')),
    ...(scrollEventResponseMs ? { scrollEventResponseMs } : {}),
    ...(scrollMovementPx ? { scrollMovementPx } : {}),
    ...(biligRuntimeProof ? { biligRuntimeProof } : {}),
    corpusVerification: parseSameCorpusVerification(objectField(value, 'corpusVerification')),
    limitations: stringArrayField(value, 'limitations'),
  }
}

function parseSameCorpusCaptureCase(value: unknown): SameCorpusCaptureCase {
  const record = asObject(value, 'UI responsiveness same-corpus capture case')
  const microsoftExcelWeb = Object.hasOwn(record, 'microsoftExcelWeb')
    ? parseSameCorpusCaptureMeasurement(objectField(record, 'microsoftExcelWeb'), 'microsoft-excel-web')
    : undefined
  const scenarioCaseFields = parseSameCorpusScenarioCaseFields(record)
  return {
    id: stringField(record, 'id'),
    corpusCaseId: stringField(record, 'corpusCaseId'),
    materializedCells: numberField(record, 'materializedCells'),
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    ...scenarioCaseFields,
    scenarioProof: parseSameCorpusScenarioProof(objectField(record, 'scenarioProof')),
    bilig: parseSameCorpusCaptureMeasurement(objectField(record, 'bilig'), 'bilig'),
    googleSheets: parseSameCorpusCaptureMeasurement(objectField(record, 'googleSheets'), 'google-sheets'),
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
  }
}

function parseSameCorpusScenarioCaseFields(value: Record<string, unknown>): SameCorpusScenarioCaseFields {
  const microsoftExcelWebMeanMs = optionalNumberField(value, 'microsoftExcelWebMeanMs')
  const microsoftExcelWebP95Ms = optionalNumberField(value, 'microsoftExcelWebP95Ms')
  const microsoftExcelWebMeanRatio = optionalNumberField(value, 'microsoftExcelWebMeanRatio')
  const microsoftExcelWebP95Ratio = optionalNumberField(value, 'microsoftExcelWebP95Ratio')
  return {
    biligMeanMs: numberField(value, 'biligMeanMs'),
    biligP95Ms: numberField(value, 'biligP95Ms'),
    googleMeanMs: numberField(value, 'googleMeanMs'),
    googleP95Ms: numberField(value, 'googleP95Ms'),
    ...(microsoftExcelWebMeanMs !== undefined ? { microsoftExcelWebMeanMs } : {}),
    ...(microsoftExcelWebP95Ms !== undefined ? { microsoftExcelWebP95Ms } : {}),
    meanRatio: numberField(value, 'meanRatio'),
    p95Ratio: numberField(value, 'p95Ratio'),
    ...(microsoftExcelWebMeanRatio !== undefined ? { microsoftExcelWebMeanRatio } : {}),
    ...(microsoftExcelWebP95Ratio !== undefined ? { microsoftExcelWebP95Ratio } : {}),
    screenshotProof: parseSameCorpusScreenshotProof(objectField(value, 'screenshotProof')),
    pixelGridProof: parseSameCorpusPixelGridProof(objectField(value, 'pixelGridProof')),
    semanticUiProof: parseOptionalSameCorpusSemanticUiProof(value),
  }
}

function parseSameCorpusCaptureMeasurement(
  value: Record<string, unknown>,
  product: UiResponsivenessSameCorpusProduct,
): SameCorpusCaptureMeasurement {
  const parsedProduct = parseSameCorpusProduct(stringField(value, 'product'))
  if (parsedProduct !== product) {
    throw new Error(`UI responsiveness same-corpus capture product mismatch: expected ${product}, got ${parsedProduct}`)
  }
  return {
    product: parsedProduct,
    source: stringField(value, 'source'),
    operationResponseMsSamples: numericArrayField(value, 'operationResponseMsSamples'),
    operationResponseProofs: stringArrayField(value, 'operationResponseProofs').map(parseSameCorpusOperationResponseProof),
    ...(Object.hasOwn(value, 'authoritativeRenderProofMsSamples')
      ? { authoritativeRenderProofMsSamples: numericArrayField(value, 'authoritativeRenderProofMsSamples') }
      : {}),
    postOperationFrameMsSamples: numericArrayField(value, 'postOperationFrameMsSamples'),
    ...(Object.hasOwn(value, 'scrollEventResponseMsSamples')
      ? { scrollEventResponseMsSamples: numericArrayField(value, 'scrollEventResponseMsSamples') }
      : {}),
    ...(Object.hasOwn(value, 'scrollMovementPxSamples')
      ? { scrollMovementPxSamples: numericArrayField(value, 'scrollMovementPxSamples') }
      : {}),
    ...(Object.hasOwn(value, 'biligRuntimeProof')
      ? { biligRuntimeProof: parseBiligRuntimeProof(objectField(value, 'biligRuntimeProof')) }
      : {}),
    corpusVerification: parseSameCorpusVerification(objectField(value, 'corpusVerification')),
    limitations: stringArrayField(value, 'limitations'),
  }
}

function parseSameCorpusOperationResponseProof(value: string): SameCorpusOperationResponseProof {
  if (value === 'load-to-ready' || value === 'visible-non-scroll-response' || value === 'visible-scroll-movement') {
    return value
  }
  throw new Error(`Unexpected same-corpus operation response proof: ${value}`)
}

function parseBiligRuntimeProof(value: Record<string, unknown>): SameCorpusBiligRuntimeProof {
  return {
    product: literalField(value, 'product', 'bilig'),
    source: stringField(value, 'source'),
    verificationMethod: literalField(value, 'verificationMethod', 'window.__biligRuntimeBuild'),
    requiredBuildKind: literalField(value, 'requiredBuildKind', 'production'),
    actualBuildKind: parseBiligRuntimeBuildKind(stringField(value, 'actualBuildKind')),
    mode: stringField(value, 'mode'),
    dev: booleanField(value, 'dev'),
    prod: booleanField(value, 'prod'),
    remoteSyncEnabled: nullableBooleanField(value, 'remoteSyncEnabled'),
    entryRoute: nullableStringField(value, 'entryRoute'),
    sampleCount: numberField(value, 'sampleCount'),
    verified: booleanField(value, 'verified'),
    samples: arrayField(value, 'samples').map(parseBiligRuntimeProofSample),
  }
}

function parseBiligRuntimeProofSample(value: unknown): SameCorpusBiligRuntimeProof['samples'][number] {
  const record = asObject(value, 'UI responsiveness same-corpus Bilig runtime proof sample')
  return {
    sampleIndex: numberField(record, 'sampleIndex'),
    present: booleanField(record, 'present'),
    app: nullableStringField(record, 'app'),
    buildKind: parseBiligRuntimeBuildKind(stringField(record, 'buildKind')),
    mode: stringField(record, 'mode'),
    dev: booleanField(record, 'dev'),
    prod: booleanField(record, 'prod'),
    remoteSyncEnabled: nullableBooleanField(record, 'remoteSyncEnabled'),
    entryRoute: nullableStringField(record, 'entryRoute'),
  }
}

function parseBiligRuntimeBuildKind(value: string): SameCorpusBiligRuntimeProof['actualBuildKind'] {
  if (value === 'development' || value === 'production' || value === 'unknown') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus Bilig runtime build kind: ${value}`)
}

function parseSameCorpusVerification(value: Record<string, unknown>): SameCorpusCaptureCorpusVerification {
  return {
    verified: booleanField(value, 'verified'),
    method: parseSameCorpusVerificationMethod(stringField(value, 'method')),
    sheetName: stringField(value, 'sheetName'),
    materializedCells: numberField(value, 'materializedCells'),
    corpusFingerprint: parseSameCorpusCorpusFingerprint(objectField(value, 'corpusFingerprint')),
    sourceWorkbookSha256: nullableStringField(value, 'sourceWorkbookSha256'),
    checkedCells: arrayField(value, 'checkedCells').map(parseSameCorpusVerifiedCell),
  }
}

function parseSameCorpusCorpusFingerprint(value: unknown): SameCorpusCaptureCorpusFingerprint {
  const record = asObject(value, 'UI responsiveness same-corpus benchmark fingerprint')
  const primaryViewport = objectField(record, 'primaryViewport')
  return {
    version: literalField(record, 'version', 'same-corpus-fingerprint-v1'),
    corpusCaseId: stringField(record, 'corpusCaseId'),
    workbookName: stringField(record, 'workbookName'),
    sheetCount: numberField(record, 'sheetCount'),
    materializedCells: numberField(record, 'materializedCells'),
    primaryViewport: {
      sheetName: stringField(primaryViewport, 'sheetName'),
      rowStart: numberField(primaryViewport, 'rowStart'),
      rowEnd: numberField(primaryViewport, 'rowEnd'),
      colStart: numberField(primaryViewport, 'colStart'),
      colEnd: numberField(primaryViewport, 'colEnd'),
    },
    snapshotSha256: stringField(record, 'snapshotSha256'),
  }
}

function parseSameCorpusProductSourceWorkbookFingerprint(value: unknown): SameCorpusProductSourceWorkbookFingerprint {
  const record = asObject(value, 'UI responsiveness same-corpus product source workbook fingerprint')
  return {
    product: parseSameCorpusProduct(stringField(record, 'product')),
    method: parseSameCorpusVerificationMethod(stringField(record, 'method')),
    source: stringField(record, 'source'),
    sourceWorkbookSha256: nullableStringField(record, 'sourceWorkbookSha256'),
  }
}

function parseSameCorpusScenarioProof(value: Record<string, unknown>): SameCorpusScenarioProof {
  const microsoftExcelWebMeanMs = optionalNumberField(value, 'microsoftExcelWebMeanMs')
  const microsoftExcelWebP95Ms = optionalNumberField(value, 'microsoftExcelWebP95Ms')
  const microsoftExcelWebMeanRatio = optionalNumberField(value, 'microsoftExcelWebMeanRatio')
  const microsoftExcelWebP95Ratio = optionalNumberField(value, 'microsoftExcelWebP95Ratio')
  return {
    biligMeanMs: numberField(value, 'biligMeanMs'),
    biligP95Ms: numberField(value, 'biligP95Ms'),
    googleMeanMs: numberField(value, 'googleMeanMs'),
    googleP95Ms: numberField(value, 'googleP95Ms'),
    ...(microsoftExcelWebMeanMs !== undefined ? { microsoftExcelWebMeanMs } : {}),
    ...(microsoftExcelWebP95Ms !== undefined ? { microsoftExcelWebP95Ms } : {}),
    meanRatio: numberField(value, 'meanRatio'),
    p95Ratio: numberField(value, 'p95Ratio'),
    ...(microsoftExcelWebMeanRatio !== undefined ? { microsoftExcelWebMeanRatio } : {}),
    ...(microsoftExcelWebP95Ratio !== undefined ? { microsoftExcelWebP95Ratio } : {}),
    screenshotProof: parseSameCorpusScreenshotProof(objectField(value, 'screenshotProof')),
    pixelGridProof: parseSameCorpusPixelGridProof(objectField(value, 'pixelGridProof')),
    semanticUiProof: parseOptionalSameCorpusSemanticUiProof(value),
  }
}

function parseOptionalSameCorpusSemanticUiProof(value: Record<string, unknown>): SameCorpusSemanticUiProof {
  if (Object.hasOwn(value, 'semanticUiProof')) {
    return parseSameCorpusSemanticUiProof(objectField(value, 'semanticUiProof'))
  }
  const pixelGridProof = parseSameCorpusPixelGridProof(objectField(value, 'pixelGridProof'))
  const products = pixelGridProof.products.map((entry) => missingSameCorpusSemanticUiProductProof(entry.product))
  return {
    captured: false,
    requiredProducts: pixelGridProof.requiredProducts,
    products,
    productVerdicts: products.map(validateSameCorpusProductSemanticUiProof),
    missingProducts: [...pixelGridProof.requiredProducts],
  }
}

function parseSameCorpusScreenshotProof(value: Record<string, unknown>): SameCorpusScreenshotProof {
  return {
    captured: booleanField(value, 'captured'),
    requiredProducts: stringArrayField(value, 'requiredProducts').map(parseSameCorpusProduct),
    artifactPaths: stringArrayField(value, 'artifactPaths'),
    missingProducts: stringArrayField(value, 'missingProducts').map(parseSameCorpusProduct),
  }
}

function parseSameCorpusPixelGridProof(value: Record<string, unknown>): SameCorpusPixelGridProof {
  return {
    captured: booleanField(value, 'captured'),
    requiredProducts: stringArrayField(value, 'requiredProducts').map(parseSameCorpusProduct),
    products: arrayField(value, 'products').map(parseSameCorpusProductPixelGridProof),
    productVerdicts: arrayField(value, 'productVerdicts').map(parseSameCorpusProductPixelGridProofVerdict),
    missingProducts: stringArrayField(value, 'missingProducts').map(parseSameCorpusProduct),
  }
}

function parseSameCorpusProductPixelGridProof(value: unknown): SameCorpusProductPixelGridProof {
  const record = asObject(value, 'UI responsiveness same-corpus product pixel grid proof')
  return {
    product: parseSameCorpusProduct(stringField(record, 'product')),
    captured: booleanField(record, 'captured'),
    method: parseSameCorpusPixelGridMethod(stringField(record, 'method')),
    viewportPixelWidth: numberField(record, 'viewportPixelWidth'),
    viewportPixelHeight: numberField(record, 'viewportPixelHeight'),
    evidence: stringArrayField(record, 'evidence'),
  }
}

function parseSameCorpusProductPixelGridProofVerdict(value: unknown) {
  const record = asObject(value, 'UI responsiveness same-corpus product pixel grid proof verdict')
  return {
    product: parseSameCorpusProduct(stringField(record, 'product')),
    evidenceStatus: parseSameCorpusProductPixelGridEvidenceStatus(stringField(record, 'evidenceStatus')),
    acceptedForCurrentScorecard: booleanField(record, 'acceptedForCurrentScorecard'),
    contractVersion: nullableStringField(record, 'contractVersion'),
    requiredContractVersion: literalField(record, 'requiredContractVersion', sameCorpusUiRenderProofContractVersion),
    invalidReasons: stringArrayField(record, 'invalidReasons'),
  }
}

function parseSameCorpusSemanticUiProof(value: Record<string, unknown>): SameCorpusSemanticUiProof {
  return {
    captured: booleanField(value, 'captured'),
    requiredProducts: stringArrayField(value, 'requiredProducts').map(parseSameCorpusProduct),
    products: arrayField(value, 'products').map(parseSameCorpusProductSemanticUiProof),
    productVerdicts: arrayField(value, 'productVerdicts').map(parseSameCorpusProductSemanticUiProofVerdict),
    missingProducts: stringArrayField(value, 'missingProducts').map(parseSameCorpusProduct),
  }
}

function parseSameCorpusProductSemanticUiProof(value: unknown): SameCorpusProductSemanticUiProof {
  const record = asObject(value, 'UI responsiveness same-corpus product semantic UI proof')
  return {
    product: parseSameCorpusProduct(stringField(record, 'product')),
    captured: booleanField(record, 'captured'),
    method: parseSameCorpusSemanticUiProofMethod(stringField(record, 'method')),
    sheetName: stringField(record, 'sheetName'),
    sheetId: nullableStringField(record, 'sheetId'),
    selectedRange: nullableStringField(record, 'selectedRange'),
    checkedCells: arrayField(record, 'checkedCells').map(parseSameCorpusVerifiedCell),
    authoritativeRenderRevision: nullableStringField(record, 'authoritativeRenderRevision'),
    visibleRenderRevision: nullableStringField(record, 'visibleRenderRevision'),
    screenshotSha256: nullableStringField(record, 'screenshotSha256'),
    mutationTargetProofs: Object.hasOwn(record, 'mutationTargetProofs')
      ? arrayField(record, 'mutationTargetProofs').map(parseSameCorpusMutationTargetProof)
      : [],
    evidence: stringArrayField(record, 'evidence'),
  }
}

function parseSameCorpusMutationTargetProof(value: unknown): SameCorpusMutationTargetProof {
  const record = asObject(value, 'UI responsiveness same-corpus mutation target proof')
  return {
    sampleIndex: numberField(record, 'sampleIndex'),
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    intendedOperation: parseSameCorpusMutatingWorkload(stringField(record, 'intendedOperation')),
    intendedPayload: parseSameCorpusMutationTargetIntendedPayload(objectField(record, 'intendedPayload')),
    sheetName: stringField(record, 'sheetName'),
    targetRange: stringField(record, 'targetRange'),
    before: parseSameCorpusMutationTargetReadback(objectField(record, 'before')),
    after: parseSameCorpusMutationTargetReadback(objectField(record, 'after')),
    restored: parseSameCorpusMutationTargetReadback(objectField(record, 'restored')),
    visibleAfter: Object.hasOwn(record, 'visibleAfter')
      ? parseSameCorpusMutationTargetReadback(objectField(record, 'visibleAfter'))
      : missingSameCorpusMutationTargetReadback(),
    visibleRestored: Object.hasOwn(record, 'visibleRestored')
      ? parseSameCorpusMutationTargetReadback(objectField(record, 'visibleRestored'))
      : missingSameCorpusMutationTargetReadback(),
    authoritativeReadbackRevision: nullableStringField(record, 'authoritativeReadbackRevision'),
    visibleRenderRevision: nullableStringField(record, 'visibleRenderRevision'),
    screenshotPath: nullableStringField(record, 'screenshotPath'),
    screenshotSha256: nullableStringField(record, 'screenshotSha256'),
    undoRestoreStatus: parseSameCorpusMutationUndoRestoreStatus(stringField(record, 'undoRestoreStatus')),
  }
}

function parseSameCorpusMutationTargetIntendedPayload(value: Record<string, unknown>): SameCorpusMutationTargetProof['intendedPayload'] {
  const kind = stringField(value, 'kind')
  if (kind === 'cell-value') {
    return { kind, value: stringField(value, 'value') }
  }
  if (kind === 'formula') {
    return { kind, formula: stringField(value, 'formula') }
  }
  if (kind === 'fill-color') {
    return {
      kind,
      expectedFillColor: Object.hasOwn(value, 'expectedFillColor') ? stringField(value, 'expectedFillColor') : '',
      swatchLabel: stringField(value, 'swatchLabel'),
    }
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation target payload kind: ${kind}`)
}

function missingSameCorpusMutationTargetReadback(): SameCorpusMutationTargetReadback {
  return {
    fillColor: null,
    formula: null,
    source: 'unknown',
    value: null,
    visibleText: null,
  }
}

function parseSameCorpusMutationTargetReadback(value: Record<string, unknown>): SameCorpusMutationTargetReadback {
  const batchId = optionalNumberField(value, 'batchId')
  const capturedRevision = Object.hasOwn(value, 'capturedRevision') ? nullableStringField(value, 'capturedRevision') : undefined
  const visibleSceneProofSha256 = Object.hasOwn(value, 'visibleSceneProofSha256')
    ? nullableStringField(value, 'visibleSceneProofSha256')
    : undefined
  return {
    value: nullableStringField(value, 'value'),
    formula: nullableStringField(value, 'formula'),
    fillColor: nullableStringField(value, 'fillColor'),
    visibleText: nullableStringField(value, 'visibleText'),
    source: Object.hasOwn(value, 'source') ? parseSameCorpusMutationTargetReadbackSource(stringField(value, 'source')) : 'unknown',
    ...(batchId !== undefined ? { batchId } : {}),
    ...(capturedRevision !== undefined ? { capturedRevision } : {}),
    ...(visibleSceneProofSha256 !== undefined ? { visibleSceneProofSha256 } : {}),
  }
}

function parseSameCorpusMutationTargetReadbackSource(value: string): SameCorpusMutationTargetReadback['source'] {
  if (value === 'bilig-authoritative-range' || value === 'visible-formula-bar' || value === 'visible-grid-cell' || value === 'unknown') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation readback source: ${value}`)
}

function parseSameCorpusMutatingWorkload(value: string): SameCorpusMutationTargetProof['intendedOperation'] {
  if (value === 'edit-visible-cell' || value === 'formula-edit' || value === 'fill-format-change') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutating workload: ${value}`)
}

function parseSameCorpusMutationUndoRestoreStatus(value: string): SameCorpusMutationTargetProof['undoRestoreStatus'] {
  if (value === 'verified' || value === 'missing' || value === 'failed') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus mutation undo restore status: ${value}`)
}

function parseSameCorpusMutationTargetProofProductSummary(value: unknown): SameCorpusMutationTargetProofProductSummary {
  const record = asObject(value, 'UI responsiveness same-corpus mutation target proof product summary')
  return {
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    product: parseSameCorpusProduct(stringField(record, 'product')),
    requiredSampleCount: numberField(record, 'requiredSampleCount'),
    rawSampleCount: numberField(record, 'rawSampleCount'),
    acceptedSampleCount: numberField(record, 'acceptedSampleCount'),
    accepted: booleanField(record, 'accepted'),
    samples: arrayField(record, 'samples').map(parseSameCorpusMutationTargetProofSampleSummary),
    invalidReasons: stringArrayField(record, 'invalidReasons'),
  }
}

function parseSameCorpusMutationTargetProofSampleSummary(value: unknown): SameCorpusMutationTargetProofSampleSummary {
  const record = asObject(value, 'UI responsiveness same-corpus mutation target proof sample summary')
  return {
    sampleIndex: numberField(record, 'sampleIndex'),
    present: booleanField(record, 'present'),
    accepted: booleanField(record, 'accepted'),
    targetRange: nullableStringField(record, 'targetRange'),
    screenshotPath: nullableStringField(record, 'screenshotPath'),
    screenshotSha256: nullableStringField(record, 'screenshotSha256'),
    invalidReasons: Object.hasOwn(record, 'invalidReasons') ? stringArrayField(record, 'invalidReasons') : [],
  }
}

function parseSameCorpusProductSemanticUiProofVerdict(value: unknown): SameCorpusProductSemanticUiProofVerdict {
  const record = asObject(value, 'UI responsiveness same-corpus product semantic UI proof verdict')
  return {
    product: parseSameCorpusProduct(stringField(record, 'product')),
    evidenceStatus: parseSameCorpusProductSemanticUiEvidenceStatus(stringField(record, 'evidenceStatus')),
    acceptedForCurrentScorecard: booleanField(record, 'acceptedForCurrentScorecard'),
    invalidReasons: stringArrayField(record, 'invalidReasons'),
  }
}

function missingSameCorpusSemanticUiProductProof(product: UiResponsivenessSameCorpusProduct): SameCorpusProductSemanticUiProof {
  return {
    product,
    captured: false,
    method: parseSameCorpusSemanticUiProofMethod(
      product === 'bilig'
        ? 'bilig-visible-semantic-readback'
        : product === 'google-sheets'
          ? 'google-sheets-visible-semantic-readback'
          : 'excel-web-visible-semantic-readback',
    ),
    sheetName: '',
    sheetId: null,
    selectedRange: null,
    checkedCells: [],
    authoritativeRenderRevision: null,
    visibleRenderRevision: null,
    screenshotSha256: null,
    mutationTargetProofs: [],
    evidence: ['semantic UI proof was not captured'],
  }
}

function parseSameCorpusVerifiedCell(value: unknown): SameCorpusCaptureVerifiedCell {
  const record = asObject(value, 'UI responsiveness same-corpus verified cell')
  return {
    address: stringField(record, 'address'),
    expected: stringField(record, 'expected'),
    actual: stringField(record, 'actual'),
  }
}

function parseNumericSummary(value: Record<string, unknown>): NumericSummary {
  const confidence95 = objectField(value, 'confidence95')
  return {
    samples: arrayField(value, 'samples').map((entry) => {
      if (typeof entry !== 'number' || !Number.isFinite(entry)) {
        throw new Error('Expected numeric summary samples to contain finite numbers')
      }
      return entry
    }),
    min: numberField(value, 'min'),
    median: numberField(value, 'median'),
    p95: numberField(value, 'p95'),
    max: numberField(value, 'max'),
    mean: numberField(value, 'mean'),
    standardDeviation: numberField(value, 'standardDeviation'),
    relativeStandardDeviation: numberField(value, 'relativeStandardDeviation'),
    standardError: numberField(value, 'standardError'),
    confidence95: {
      low: numberField(confidence95, 'low'),
      high: numberField(confidence95, 'high'),
    },
  }
}

function numericArrayField(value: Record<string, unknown>, key: string): number[] {
  return arrayField(value, key).map((entry) => {
    if (typeof entry !== 'number' || !Number.isFinite(entry) || entry < 0) {
      throw new Error(`Expected ${key} to contain finite non-negative numbers`)
    }
    return entry
  })
}

function optionalNumericSummary(value: Record<string, unknown>, key: string): NumericSummary | undefined {
  return Object.hasOwn(value, key) ? parseNumericSummary(objectField(value, key)) : undefined
}

function optionalNumberField(value: Record<string, unknown>, key: string): number | undefined {
  return Object.hasOwn(value, key) ? numberField(value, key) : undefined
}

function optionalBooleanField(value: Record<string, unknown>, key: string): boolean | undefined {
  return Object.hasOwn(value, key) ? booleanField(value, key) : undefined
}

function nullableStringField(value: Record<string, unknown>, key: string): string | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${key} to be a string or null`)
  }
  return fieldValue
}

function nullableBooleanField(value: Record<string, unknown>, key: string): boolean | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${key} to be a boolean or null`)
  }
  return fieldValue
}

function optionalSameCorpusTenXMetric(
  value: Record<string, unknown>,
  key: string,
): UiResponsivenessSameCorpusCase['tenXMeanAndP95Metric'] | undefined {
  if (!Object.hasOwn(value, key)) {
    return undefined
  }
  const metric = stringField(value, key)
  if (metric === 'operationResponseMs' || metric === 'scrollEventResponseMs') {
    return metric
  }
  throw new Error(`Unexpected UI responsiveness same-corpus 10x metric: ${metric}`)
}

function parseVendor(value: string): UiResponsivenessLiveBrowserVendor {
  if (value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser vendor: ${value}`)
}

function parseSameCorpusEvidenceKind(value: string): UiResponsivenessSameCorpusProof['evidenceKind'] {
  if (value === 'same-corpus-browser-capture' || value === 'not-captured') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus evidence kind: ${value}`)
}

function parseSameCorpusProduct(value: string): UiResponsivenessSameCorpusProduct {
  if (value === 'bilig' || value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus product: ${value}`)
}

function parseSameCorpusVerificationMethod(value: string): SameCorpusCaptureCorpusVerification['method'] {
  if (value === 'bilig-benchmark-state' || value === 'google-sheets-xlsx-export' || value === 'microsoft-excel-web-source-xlsx') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus verification method: ${value}`)
}

function parseSameCorpusPixelGridMethod(value: string): SameCorpusProductPixelGridProof['method'] {
  if (value === 'typegpu-visible-canvas' || value === 'google-sheets-visible-grid' || value === 'excel-web-visible-grid') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus pixel grid proof method: ${value}`)
}

function parseSameCorpusProductPixelGridEvidenceStatus(value: string): 'current-contract' | 'legacy-insufficient' | 'missing' | 'invalid' {
  if (value === 'current-contract' || value === 'legacy-insufficient' || value === 'missing' || value === 'invalid') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus pixel grid proof evidence status: ${value}`)
}

function parseSameCorpusSemanticUiProofMethod(value: string): SameCorpusProductSemanticUiProof['method'] {
  if (
    value === 'bilig-visible-semantic-readback' ||
    value === 'google-sheets-visible-semantic-readback' ||
    value === 'excel-web-visible-semantic-readback'
  ) {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus semantic UI proof method: ${value}`)
}

function parseSameCorpusProductSemanticUiEvidenceStatus(value: string): 'current-contract' | 'missing' | 'invalid' {
  if (value === 'current-contract' || value === 'missing' || value === 'invalid') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus semantic UI proof evidence status: ${value}`)
}

function parseSameCorpusWorkload(value: string): UiResponsivenessSameCorpusWorkload {
  if (isUiResponsivenessSameCorpusWorkload(value)) {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus workload: ${value}`)
}

function parseAccessMode(value: string): UiResponsivenessLiveBrowserCase['accessMode'] {
  if (value === 'public-comment-only' || value === 'public-view-only' || value === 'public-office-web-viewer') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser access mode: ${value}`)
}
