import { existsSync } from 'node:fs'

import type { BuildScorecardInput } from './bilig-dominance-scorecard-types.ts'
import { localCiResourceGuardOverrideEnv, type LocalCiResourceGuardStatus } from './ci-local-resource-guard.ts'
import { readJsonObject } from './json-scorecard-helpers.ts'
import type {
  SameCorpusCapture,
  UiResponsivenessSameCorpusRunManifest,
  UiResponsivenessSameCorpusWorkload,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { parseSameCorpusCapture } from './ui-responsiveness-live-browser-scorecard-parse.ts'
import { buildSameCorpusProof, validateSameCorpusCaptureRunManifest } from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import { sameCorpusScenarioSummaryFieldsCurrent } from './ui-responsiveness-same-corpus-scenario-fields.ts'
import { sameCorpusUiSpeedGaps, type SameCorpusUiSpeedGap } from './ui-responsiveness-same-corpus-speed-gaps.ts'
import {
  isUiResponsivenessSameCorpusWorkload,
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
} from './ui-responsiveness-same-corpus-workloads.ts'
import type { SameCorpusPublicAccessCheck } from './ui-responsiveness-same-corpus-public-access-check.ts'
import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'

export type UiSameCorpusGoogleSheetsUrlSource = 'argument-or-environment' | 'public-access-check' | 'checked-in-capture' | 'missing'

export interface UiSameCorpusStatus {
  readonly captured: boolean
  readonly evidenceKind: 'same-corpus-browser-capture' | 'not-captured'
  readonly requiredProductCount: number
  readonly requiredCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly tenXRequirementSatisfied: boolean
  readonly runManifest: UiResponsivenessSameCorpusRunManifest | null
  readonly renderProofContractVersion: string | null
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
  readonly currentContractEvidenceComplete: boolean
  readonly googleSheetsTenXRequirementSatisfied: boolean
  readonly runManifestInvalidReasons: readonly string[]
  readonly speedGaps: readonly SameCorpusUiSpeedGap[]
  readonly requiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly scrollEventEvidenceCaseCount: number
  readonly casesMissingScrollEventEvidence: readonly string[]
  readonly coveredCorpusCaseIds: readonly string[]
  readonly limitations: readonly string[]
  readonly fixture: UiSameCorpusFixtureStatus
  readonly googleSheetsUrl: string | null
  readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
  readonly googleSheetsUrlEnvVar: string
  readonly microsoftExcelWebEditableUrl: string | null
  readonly microsoftExcelWebEditableUrlEnvVar: string
  readonly publicAccessCheckPath: string
  readonly captureArtifact: UiSameCorpusCaptureArtifactStatus
  readonly missingInputs: readonly string[]
  readonly nextFixtureCheckCommand: string
  readonly nextPublicAccessCheckCommand: string
  readonly nextGoogleSheetsStorageStateCommand: string | null
  readonly nextMicrosoftExcelWebStorageStateCommand: string | null
  readonly nextGoogleSheetsUploadInstruction: string | null
  readonly nextMicrosoftExcelWebUploadInstruction: string | null
  readonly nextPreflightCommand: string | null
  readonly nextAuthenticatedPreflightCommand: string | null
  readonly nextCaptureCommand: string | null
  readonly nextAuthenticatedCaptureCommand: string | null
  readonly blockedCommands: readonly string[]
  readonly browserCaptureGuard: UiSameCorpusBrowserCaptureGuardStatus
  readonly nextScorecardGenerateCommand: string | null
  readonly nextDominanceCheckCommand: string
}

export interface UiSameCorpusFixtureStatus {
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly materializedCells: number
  readonly localXlsxPath: string
  readonly publicGithubRawUrl: string
  readonly publicForgejoRawUrl: string
  readonly microsoftExcelWebUrl: string
}

export interface UiSameCorpusBrowserCaptureGuardStatus {
  readonly active: boolean
  readonly activeMarkerPaths: readonly string[]
  readonly overrideEnvVar: string
  readonly overridePrefix: string | null
  readonly nextPreflightRequiresOverride: boolean
  readonly nextCaptureRequiresOverride: boolean
}

export interface UiSameCorpusCaptureArtifactStatus {
  readonly path: string
  readonly exists: boolean
  readonly parseable: boolean
  readonly currentRunManifest: boolean
  readonly readyForScorecardGeneration: boolean
  readonly sampleCount: number | null
  readonly caseCount: number | null
  readonly scenarioSummaryFieldCaseCount: number | null
  readonly strictRenderedGridProofCaseCount: number | null
  readonly visibleOperationResponseProofCaseCount: number | null
  readonly biligAuthoritativeRenderProofCaseCount: number | null
  readonly semanticUiProofCaseCount: number | null
  readonly requiredMutationTargetProofCaseCount: number | null
  readonly mutationTargetProofCaseCount: number | null
  readonly requiredMutationTargetProofSampleCount: number | null
  readonly mutationTargetProofSampleCount: number | null
  readonly legacyInsufficientRenderedGridProofCaseCount: number | null
  readonly tenXMeanAndP95CaseCount: number | null
  readonly currentContractEvidenceComplete: boolean | null
  readonly googleSheetsTenXRequirementSatisfied: boolean | null
  readonly capturedWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly captureRunSignature: string | null
  readonly readinessErrors: readonly string[]
  readonly runManifestInvalidReasons: readonly string[]
  readonly legacyCapture: UiSameCorpusLegacyCaptureArtifactStatus | null
}

export interface UiSameCorpusLegacyCaptureArtifactStatus {
  readonly schemaVersion: number | null
  readonly suite: string | null
  readonly sampleCount: number | null
  readonly caseCount: number | null
  readonly capturedWorkloads: readonly string[]
  readonly missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly pixelGridProofCaseCount: number | null
  readonly tenXMeanAndP95CaseCount: number | null
  readonly googleSheetsTenXRequirementSatisfied: boolean | null
  readonly contractGap: 'missing-run-manifest' | 'not-current-same-corpus-capture'
}

export const uiSameCorpusGoogleSheetsUrlEnvVar = 'BILIG_UI_SAME_CORPUS_GOOGLE_SHEETS_URL'
export const uiSameCorpusMicrosoftExcelWebUrlEnvVar = 'BILIG_UI_SAME_CORPUS_MICROSOFT_EXCEL_WEB_URL'

const defaultUiSameCorpusId: WorkbookBenchmarkCorpusId = 'wide-mixed-250k'
const requiredUiSameCorpusWorkloads = requiredUiResponsivenessSameCorpusWorkloads

export function buildUiSameCorpusStatus(
  input: BuildScorecardInput,
  args: {
    readonly googleSheetsUrl: string | null
    readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
    readonly localCiResourceGuardStatus: LocalCiResourceGuardStatus
    readonly microsoftExcelWebEditableUrl: string | null
    readonly publicAccessCheckPath: string
    readonly captureArtifact?: UiSameCorpusCaptureArtifactStatus
  },
): UiSameCorpusStatus {
  const proof = input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof
  const runManifest = proof.runManifest ?? null
  const fixture = uiSameCorpusFixtureStatus(defaultUiSameCorpusId)
  const coveredWorkloads = new Set(proof.cases.map((entry) => entry.workload))
  const missingRequiredWorkloads = requiredUiSameCorpusWorkloads.filter((workload) => !coveredWorkloads.has(workload))
  const requiredProofCases = proof.cases.filter((entry) => requiredUiSameCorpusWorkloads.includes(entry.workload))
  const scrollEvidenceRequiredProofCases = requiredProofCases.filter((entry) =>
    uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload),
  )
  const casesMissingScrollEventEvidence = scrollEvidenceRequiredProofCases
    .filter((entry) => !uiSameCorpusCaseHasScrollEventEvidence(entry))
    .map((entry) => entry.id)
  const scrollEventEvidenceCaseCount = Math.max(0, scrollEvidenceRequiredProofCases.length - casesMissingScrollEventEvidence.length)
  const tenXRequirementSatisfied = uiSameCorpusTenXRequirementSatisfied(proof, missingRequiredWorkloads, casesMissingScrollEventEvidence)
  const googleSheetsUrlArgument = args.googleSheetsUrl ?? '<google-sheets-url>'
  const productionBiligUrlArgument = '<production-bilig-url>'
  const microsoftExcelWebUrlArgument = args.microsoftExcelWebEditableUrl ?? '<microsoft-excel-web-editable-url>'
  const browserCaptureGuard = buildBrowserCaptureGuardStatus(args.localCiResourceGuardStatus)
  const missingInputs = args.googleSheetsUrl || tenXRequirementSatisfied ? [] : ['googleSheetsUrlForUploadedSameCorpusWorkbook']
  const nextGoogleSheetsUploadInstruction = missingInputs.includes('googleSheetsUrlForUploadedSameCorpusWorkbook')
    ? `Upload ${fixture.localXlsxPath} to Google Sheets as a native Google Sheet, share it to anyone with the link, then pass its edit URL as --google-sheets-url.`
    : null
  const nextMicrosoftExcelWebUploadInstruction = missingInputs.includes('microsoftExcelWebEditableUrlForUploadedSameCorpusWorkbook')
    ? `Upload ${fixture.localXlsxPath} to OneDrive or Microsoft 365, open it as an editable Excel Web workbook, then pass its browser URL as --microsoft-excel-web-url. The Office viewer URL is only valid for public XLSX identity checks.`
    : null
  const nextGoogleSheetsStorageStateCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--save-storage-state',
    '.cache/ui-responsiveness/google-sheets-storage-state.json',
    '--auth-product',
    'google-sheets',
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--corpus',
    fixture.corpusCaseId,
  ]
    .map(shellQuote)
    .join(' ')
  const nextMicrosoftExcelWebStorageStateCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--save-storage-state',
    '.cache/ui-responsiveness/microsoft-excel-web-storage-state.json',
    '--auth-product',
    'microsoft-excel-web',
    '--microsoft-excel-web-url',
    microsoftExcelWebUrlArgument,
    '--corpus',
    fixture.corpusCaseId,
  ]
    .map(shellQuote)
    .join(' ')
  const nextPreflightCommand = ['pnpm', 'ui:same-corpus:capture', '--', '--preflight', '--google-sheets-url', googleSheetsUrlArgument]
    .map(shellQuote)
    .join(' ')
  const nextAuthenticatedPreflightCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--preflight',
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--google-sheets-storage-state',
    '.cache/ui-responsiveness/google-sheets-storage-state.json',
  ]
    .map(shellQuote)
    .join(' ')
  const nextCaptureCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--output',
    '.cache/ui-responsiveness/same-corpus-capture.json',
    '--bilig-url',
    productionBiligUrlArgument,
    '--google-sheets-url',
    googleSheetsUrlArgument,
  ]
    .map(shellQuote)
    .join(' ')
  const nextAuthenticatedCaptureCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--output',
    '.cache/ui-responsiveness/same-corpus-capture.json',
    '--bilig-url',
    productionBiligUrlArgument,
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--google-sheets-storage-state',
    '.cache/ui-responsiveness/google-sheets-storage-state.json',
  ]
    .map(shellQuote)
    .join(' ')
  const nextScorecardGenerateCommand = 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json'
  return {
    captured: proof.captured,
    evidenceKind: proof.evidenceKind,
    requiredProductCount: proof.requiredProductCount,
    requiredCaseCount: proof.requiredCaseCount,
    tenXMeanAndP95CaseCount: proof.tenXMeanAndP95CaseCount,
    tenXRequirementSatisfied,
    runManifest,
    renderProofContractVersion: runManifest?.contractVersion ?? null,
    scenarioSummaryFieldCaseCount:
      runManifest?.scenarioSummaryFieldCaseCount ?? proof.cases.filter(sameCorpusScenarioSummaryFieldsCurrent).length,
    strictRenderedGridProofCaseCount:
      runManifest?.strictRenderedGridProofCaseCount ?? proof.cases.filter((entry) => entry.scenarioProof.pixelGridProof.captured).length,
    visibleOperationResponseProofCaseCount:
      runManifest?.visibleOperationResponseProofCaseCount ??
      proof.cases.filter((entry) => entry.operationResponseProofGuardrailPassed === true).length,
    biligAuthoritativeRenderProofCaseCount:
      runManifest?.biligAuthoritativeRenderProofCaseCount ??
      proof.cases.filter((entry) => entry.authoritativeRenderProofGuardrailPassed === true).length,
    semanticUiProofCaseCount:
      runManifest?.semanticUiProofCaseCount ?? proof.cases.filter((entry) => entry.scenarioProof.semanticUiProof.captured).length,
    requiredMutationTargetProofCaseCount: runManifest?.requiredMutationTargetProofCaseCount ?? 0,
    mutationTargetProofCaseCount: runManifest?.mutationTargetProofCaseCount ?? 0,
    requiredMutationTargetProofSampleCount: runManifest?.requiredMutationTargetProofSampleCount ?? 0,
    mutationTargetProofSampleCount: runManifest?.mutationTargetProofSampleCount ?? 0,
    legacyInsufficientRenderedGridProofCaseCount:
      runManifest?.legacyInsufficientRenderedGridProofCaseCount ?? proof.cases.filter(hasLegacyInsufficientRenderedGridProof).length,
    currentContractEvidenceComplete: runManifest?.currentContractEvidenceComplete ?? false,
    googleSheetsTenXRequirementSatisfied: runManifest?.googleSheetsTenXRequirementSatisfied ?? false,
    runManifestInvalidReasons: runManifest?.invalidReasons ?? ['same-corpus UI proof is missing a run manifest'],
    speedGaps: sameCorpusUiSpeedGaps(proof),
    requiredWorkloads: requiredUiSameCorpusWorkloads,
    missingRequiredWorkloads,
    scrollEventEvidenceCaseCount,
    casesMissingScrollEventEvidence,
    coveredCorpusCaseIds: proof.coveredCorpusCaseIds,
    limitations: proof.limitations,
    fixture,
    googleSheetsUrl: args.googleSheetsUrl,
    googleSheetsUrlSource: args.googleSheetsUrlSource,
    googleSheetsUrlEnvVar: uiSameCorpusGoogleSheetsUrlEnvVar,
    microsoftExcelWebEditableUrl: args.microsoftExcelWebEditableUrl,
    microsoftExcelWebEditableUrlEnvVar: uiSameCorpusMicrosoftExcelWebUrlEnvVar,
    publicAccessCheckPath: args.publicAccessCheckPath,
    captureArtifact:
      args.captureArtifact ?? buildMissingUiSameCorpusCaptureArtifactStatus('.cache/ui-responsiveness/same-corpus-capture.json'),
    missingInputs,
    nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
    nextPublicAccessCheckCommand: [
      'pnpm',
      'ui:same-corpus:public-check',
      '--',
      '--output',
      args.publicAccessCheckPath,
      '--google-sheets-url',
      googleSheetsUrlArgument,
      '--microsoft-excel-web-url',
      fixture.microsoftExcelWebUrl,
    ]
      .map(shellQuote)
      .join(' '),
    nextGoogleSheetsStorageStateCommand: browserCaptureGuard.active ? null : nextGoogleSheetsStorageStateCommand,
    nextMicrosoftExcelWebStorageStateCommand: browserCaptureGuard.active ? null : nextMicrosoftExcelWebStorageStateCommand,
    nextGoogleSheetsUploadInstruction,
    nextMicrosoftExcelWebUploadInstruction,
    nextPreflightCommand: browserCaptureGuard.active ? null : nextPreflightCommand,
    nextAuthenticatedPreflightCommand: browserCaptureGuard.active ? null : nextAuthenticatedPreflightCommand,
    nextCaptureCommand: browserCaptureGuard.active ? null : nextCaptureCommand,
    nextAuthenticatedCaptureCommand: browserCaptureGuard.active ? null : nextAuthenticatedCaptureCommand,
    blockedCommands: browserCaptureGuard.active
      ? [
          nextGoogleSheetsStorageStateCommand,
          nextMicrosoftExcelWebStorageStateCommand,
          nextPreflightCommand,
          nextAuthenticatedPreflightCommand,
          nextCaptureCommand,
          nextAuthenticatedCaptureCommand,
          nextScorecardGenerateCommand,
        ].map(localCiResourceGuardOverrideCommand)
      : [],
    browserCaptureGuard,
    nextScorecardGenerateCommand: browserCaptureGuard.active ? null : nextScorecardGenerateCommand,
    nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
  }
}

export function readUiSameCorpusCaptureArtifactStatus(args: {
  readonly path: string
  readonly displayPath?: string
}): UiSameCorpusCaptureArtifactStatus {
  const displayPath = args.displayPath ?? args.path
  if (!existsSync(args.path)) {
    return buildMissingUiSameCorpusCaptureArtifactStatus(displayPath)
  }
  const raw = readJsonObject(args.path)
  try {
    const capture = parseSameCorpusCapture(raw)
    return buildUiSameCorpusCaptureArtifactStatus(displayPath, capture)
  } catch (error) {
    return buildInvalidUiSameCorpusCaptureArtifactStatus(
      displayPath,
      true,
      false,
      buildUiSameCorpusCaptureReadinessErrors(errorMessage(error), inspectLegacySameCorpusCaptureArtifact(raw)),
      inspectLegacySameCorpusCaptureArtifact(raw),
    )
  }
}

export function buildMissingUiSameCorpusCaptureArtifactStatus(path: string): UiSameCorpusCaptureArtifactStatus {
  return buildInvalidUiSameCorpusCaptureArtifactStatus(path, false, false, 'same-corpus capture artifact is missing')
}

function buildUiSameCorpusCaptureArtifactStatus(path: string, capture: SameCorpusCapture): UiSameCorpusCaptureArtifactStatus {
  try {
    validateSameCorpusCaptureRunManifest(capture)
  } catch (error) {
    return buildParsedUiSameCorpusCaptureArtifactStatus(path, capture, false, false, [errorMessage(error)], [])
  }

  try {
    const proof = buildSameCorpusProof(capture)
    const runManifest = proof.runManifest
    return {
      path,
      exists: true,
      parseable: true,
      currentRunManifest: true,
      readyForScorecardGeneration: true,
      sampleCount: capture.sampleCount,
      caseCount: capture.cases.length,
      scenarioSummaryFieldCaseCount: runManifest?.scenarioSummaryFieldCaseCount ?? null,
      strictRenderedGridProofCaseCount: runManifest?.strictRenderedGridProofCaseCount ?? null,
      visibleOperationResponseProofCaseCount: runManifest?.visibleOperationResponseProofCaseCount ?? null,
      biligAuthoritativeRenderProofCaseCount: runManifest?.biligAuthoritativeRenderProofCaseCount ?? null,
      semanticUiProofCaseCount: runManifest?.semanticUiProofCaseCount ?? null,
      requiredMutationTargetProofCaseCount: runManifest?.requiredMutationTargetProofCaseCount ?? null,
      mutationTargetProofCaseCount: runManifest?.mutationTargetProofCaseCount ?? null,
      requiredMutationTargetProofSampleCount: runManifest?.requiredMutationTargetProofSampleCount ?? null,
      mutationTargetProofSampleCount: runManifest?.mutationTargetProofSampleCount ?? null,
      legacyInsufficientRenderedGridProofCaseCount: runManifest?.legacyInsufficientRenderedGridProofCaseCount ?? null,
      tenXMeanAndP95CaseCount: proof.tenXMeanAndP95CaseCount,
      currentContractEvidenceComplete: runManifest?.currentContractEvidenceComplete ?? null,
      googleSheetsTenXRequirementSatisfied: runManifest?.googleSheetsTenXRequirementSatisfied ?? null,
      capturedWorkloads: [...capture.runManifest.capturedWorkloads],
      missingRequiredWorkloads: requiredUiSameCorpusWorkloads.filter(
        (workload) => !capture.runManifest.capturedWorkloads.includes(workload),
      ),
      captureRunSignature: capture.runManifest.captureRunSignature,
      readinessErrors: [],
      runManifestInvalidReasons: runManifest?.invalidReasons ?? [],
      legacyCapture: null,
    }
  } catch (error) {
    return buildParsedUiSameCorpusCaptureArtifactStatus(
      path,
      capture,
      true,
      false,
      [errorMessage(error)],
      [...capture.runManifest.invalidReasons],
    )
  }
}

function buildParsedUiSameCorpusCaptureArtifactStatus(
  path: string,
  capture: SameCorpusCapture,
  currentRunManifest: boolean,
  readyForScorecardGeneration: boolean,
  readinessErrors: readonly string[],
  runManifestInvalidReasons: readonly string[],
): UiSameCorpusCaptureArtifactStatus {
  return {
    path,
    exists: true,
    parseable: true,
    currentRunManifest,
    readyForScorecardGeneration,
    sampleCount: capture.sampleCount,
    caseCount: capture.cases.length,
    scenarioSummaryFieldCaseCount: capture.runManifest.scenarioSummaryFieldCaseCount,
    strictRenderedGridProofCaseCount: capture.runManifest.strictRenderedGridProofCaseCount,
    visibleOperationResponseProofCaseCount: capture.runManifest.visibleOperationResponseProofCaseCount,
    biligAuthoritativeRenderProofCaseCount: capture.runManifest.biligAuthoritativeRenderProofCaseCount,
    semanticUiProofCaseCount: capture.runManifest.semanticUiProofCaseCount,
    requiredMutationTargetProofCaseCount: capture.runManifest.requiredMutationTargetProofCaseCount,
    mutationTargetProofCaseCount: capture.runManifest.mutationTargetProofCaseCount,
    requiredMutationTargetProofSampleCount: capture.runManifest.requiredMutationTargetProofSampleCount,
    mutationTargetProofSampleCount: capture.runManifest.mutationTargetProofSampleCount,
    legacyInsufficientRenderedGridProofCaseCount: capture.runManifest.legacyInsufficientRenderedGridProofCaseCount,
    tenXMeanAndP95CaseCount: capture.runManifest.tenXMeanAndP95CaseCount,
    currentContractEvidenceComplete: capture.runManifest.currentContractEvidenceComplete,
    googleSheetsTenXRequirementSatisfied: capture.runManifest.googleSheetsTenXRequirementSatisfied,
    capturedWorkloads: [...capture.runManifest.capturedWorkloads],
    missingRequiredWorkloads: requiredUiSameCorpusWorkloads.filter((workload) => !capture.runManifest.capturedWorkloads.includes(workload)),
    captureRunSignature: capture.runManifest.captureRunSignature,
    readinessErrors,
    runManifestInvalidReasons,
    legacyCapture: null,
  }
}

function buildInvalidUiSameCorpusCaptureArtifactStatus(
  path: string,
  exists: boolean,
  parseable: boolean,
  readinessErrors: string | readonly string[],
  legacyCapture: UiSameCorpusLegacyCaptureArtifactStatus | null = null,
): UiSameCorpusCaptureArtifactStatus {
  return {
    path,
    exists,
    parseable,
    currentRunManifest: false,
    readyForScorecardGeneration: false,
    sampleCount: null,
    caseCount: null,
    scenarioSummaryFieldCaseCount: null,
    strictRenderedGridProofCaseCount: null,
    visibleOperationResponseProofCaseCount: null,
    biligAuthoritativeRenderProofCaseCount: null,
    semanticUiProofCaseCount: null,
    requiredMutationTargetProofCaseCount: null,
    mutationTargetProofCaseCount: null,
    requiredMutationTargetProofSampleCount: null,
    mutationTargetProofSampleCount: null,
    legacyInsufficientRenderedGridProofCaseCount: null,
    tenXMeanAndP95CaseCount: null,
    currentContractEvidenceComplete: null,
    googleSheetsTenXRequirementSatisfied: null,
    capturedWorkloads: [],
    missingRequiredWorkloads: [...requiredUiSameCorpusWorkloads],
    captureRunSignature: null,
    readinessErrors: typeof readinessErrors === 'string' ? [readinessErrors] : [...readinessErrors],
    runManifestInvalidReasons: [],
    legacyCapture,
  }
}

function inspectLegacySameCorpusCaptureArtifact(value: Record<string, unknown>): UiSameCorpusLegacyCaptureArtifactStatus | null {
  const suite = typeof value['suite'] === 'string' ? value['suite'] : null
  const schemaVersion =
    typeof value['schemaVersion'] === 'number' && Number.isFinite(value['schemaVersion']) ? value['schemaVersion'] : null
  const casesValue = value['cases']
  const cases = Array.isArray(casesValue) ? casesValue.filter(isRecord) : []
  const capturedWorkloads = uniqueStrings(
    cases.map((entry) => (typeof entry['workload'] === 'string' ? entry['workload'] : null)).filter((entry) => entry !== null),
  )
  const capturedSameCorpusWorkloads = new Set(capturedWorkloads.filter(isUiResponsivenessSameCorpusWorkload))
  const missingRequiredWorkloads = requiredUiSameCorpusWorkloads.filter((workload) => !capturedSameCorpusWorkloads.has(workload))
  const tenXMeanAndP95CaseCount = cases.filter(legacyCaseHasTenXMeanAndP95).length
  const pixelGridProofCaseCount = cases.filter(legacyCaseHasPixelGridProof).length
  const isSameCorpusCapture =
    suite === 'ui-responsiveness-same-corpus-capture' || cases.some((entry) => typeof entry['workload'] === 'string')
  if (!isSameCorpusCapture) {
    return null
  }
  return {
    schemaVersion,
    suite,
    sampleCount: typeof value['sampleCount'] === 'number' && Number.isFinite(value['sampleCount']) ? value['sampleCount'] : null,
    caseCount: Array.isArray(casesValue) ? cases.length : null,
    capturedWorkloads,
    missingRequiredWorkloads,
    pixelGridProofCaseCount,
    tenXMeanAndP95CaseCount,
    googleSheetsTenXRequirementSatisfied:
      cases.length > 0 &&
      missingRequiredWorkloads.length === 0 &&
      tenXMeanAndP95CaseCount === requiredUiSameCorpusWorkloads.length &&
      pixelGridProofCaseCount === requiredUiSameCorpusWorkloads.length,
    contractGap: Object.hasOwn(value, 'runManifest') ? 'not-current-same-corpus-capture' : 'missing-run-manifest',
  }
}

function buildUiSameCorpusCaptureReadinessErrors(
  parseError: string,
  legacyCapture: UiSameCorpusLegacyCaptureArtifactStatus | null,
): readonly string[] {
  if (!legacyCapture) {
    return [parseError]
  }
  const contractGap =
    legacyCapture.contractGap === 'missing-run-manifest'
      ? 'legacy same-corpus capture artifact is missing the current runManifest contract'
      : 'same-corpus capture artifact does not match the current runManifest contract'
  return [contractGap, parseError]
}

function legacyCaseHasTenXMeanAndP95(value: Record<string, unknown>): boolean {
  return value['tenXMeanAndP95AgainstGoogleSheets'] === true
}

function legacyCaseHasPixelGridProof(value: Record<string, unknown>): boolean {
  const directProof = isRecord(value['pixelGridProof']) ? value['pixelGridProof'] : null
  const scenarioProof = isRecord(value['scenarioProof']) ? value['scenarioProof'] : null
  const nestedProof = scenarioProof && isRecord(scenarioProof['pixelGridProof']) ? scenarioProof['pixelGridProof'] : null
  const proof = directProof ?? nestedProof
  return proof?.['captured'] === true
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function uniqueStrings(values: readonly string[]): readonly string[] {
  return [...new Set(values)]
}

export function resolveUiSameCorpusGoogleSheetsUrl(args: {
  readonly corpusCaseId?: WorkbookBenchmarkCorpusId
  readonly explicitGoogleSheetsUrl: string | null
  readonly publicAccessCheck: SameCorpusPublicAccessCheck | null
  readonly sameCorpusProof: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof']
}): {
  readonly googleSheetsUrl: string | null
  readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
} {
  const corpusCaseId = args.corpusCaseId ?? defaultUiSameCorpusId
  if (args.explicitGoogleSheetsUrl) {
    return {
      googleSheetsUrl: args.explicitGoogleSheetsUrl,
      googleSheetsUrlSource: 'argument-or-environment',
    }
  }
  const verifiedPublicAccessUrl = verifiedGoogleSheetsUrlFromPublicAccessCheck(args.publicAccessCheck, corpusCaseId)
  if (verifiedPublicAccessUrl) {
    return {
      googleSheetsUrl: verifiedPublicAccessUrl,
      googleSheetsUrlSource: 'public-access-check',
    }
  }
  const verifiedCheckedInCaptureUrl = verifiedGoogleSheetsUrlFromSameCorpusProof(args.sameCorpusProof, corpusCaseId)
  if (verifiedCheckedInCaptureUrl) {
    return {
      googleSheetsUrl: verifiedCheckedInCaptureUrl,
      googleSheetsUrlSource: 'checked-in-capture',
    }
  }
  return {
    googleSheetsUrl: null,
    googleSheetsUrlSource: 'missing',
  }
}

function buildBrowserCaptureGuardStatus(status: LocalCiResourceGuardStatus): UiSameCorpusBrowserCaptureGuardStatus {
  const active = status.activeMarkerPaths.length > 0
  return {
    active,
    activeMarkerPaths: status.activeMarkerPaths,
    overrideEnvVar: localCiResourceGuardOverrideEnv,
    overridePrefix: active ? `${localCiResourceGuardOverrideEnv}=1` : null,
    nextPreflightRequiresOverride: active,
    nextCaptureRequiresOverride: active,
  }
}

function verifiedGoogleSheetsUrlFromPublicAccessCheck(
  check: SameCorpusPublicAccessCheck | null,
  corpusCaseId: WorkbookBenchmarkCorpusId,
): string | null {
  if (!check || check.corpusCaseId !== corpusCaseId) {
    return null
  }
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  if (check.materializedCells !== corpus.materializedCellCount) {
    return null
  }
  const product = check.products.find((entry) => entry.product === 'google-sheets')
  return product?.corpusVerification.verified ? product.source : null
}

function verifiedGoogleSheetsUrlFromSameCorpusProof(
  proof: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof'],
  corpusCaseId: WorkbookBenchmarkCorpusId,
): string | null {
  if (!proof.captured || proof.evidenceKind !== 'same-corpus-browser-capture' || !proof.coveredCorpusCaseIds.includes(corpusCaseId)) {
    return null
  }
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  const corpusCases = proof.cases.filter((entry) => entry.corpusCaseId === corpusCaseId)
  if (corpusCases.length === 0) {
    return null
  }
  let url: string | null = null
  for (const entry of corpusCases) {
    const googleSheets = entry.googleSheets
    const source = googleSheets.source.trim()
    if (
      source.length === 0 ||
      !googleSheets.corpusVerification.verified ||
      googleSheets.corpusVerification.method !== 'google-sheets-xlsx-export' ||
      googleSheets.corpusVerification.materializedCells !== corpus.materializedCellCount
    ) {
      return null
    }
    if (url !== null && url !== source) {
      return null
    }
    url = source
  }
  return url
}

function uiSameCorpusTenXRequirementSatisfied(
  proof: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof'],
  missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[],
  casesMissingScrollEventEvidence: readonly string[],
): boolean {
  return (
    proof.captured &&
    proof.evidenceKind === 'same-corpus-browser-capture' &&
    proof.runManifest?.currentContractEvidenceComplete === true &&
    proof.runManifest.googleSheetsTenXRequirementSatisfied &&
    proof.runManifest.invalidReasons.length === 0 &&
    proof.runManifest.scenarioSummaryFieldCaseCount === proof.requiredCaseCount &&
    proof.runManifest.strictRenderedGridProofCaseCount === proof.requiredCaseCount &&
    proof.runManifest.visibleOperationResponseProofCaseCount === proof.requiredCaseCount &&
    proof.runManifest.mutationTargetProofCaseCount === proof.runManifest.requiredMutationTargetProofCaseCount &&
    proof.runManifest.mutationTargetProofSampleCount === proof.runManifest.requiredMutationTargetProofSampleCount &&
    proof.requiredProductCount === 2 &&
    proof.requiredCaseCount > 0 &&
    proof.cases.length === proof.requiredCaseCount &&
    proof.tenXMeanAndP95CaseCount === proof.requiredCaseCount &&
    missingRequiredWorkloads.length === 0 &&
    casesMissingScrollEventEvidence.length === 0 &&
    proof.cases.every((entry) => entry.tenXMeanAndP95AgainstGoogleSheets && entry.passed)
  )
}

function uiSameCorpusCaseHasScrollEventEvidence(
  entry: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof']['cases'][number],
): boolean {
  return (
    entry.tenXMeanAndP95Metric === 'scrollEventResponseMs' &&
    Boolean(entry.bilig.scrollEventResponseMs) &&
    Boolean(entry.googleSheets.scrollEventResponseMs) &&
    Boolean(entry.bilig.scrollMovementPx) &&
    Boolean(entry.googleSheets.scrollMovementPx)
  )
}

function hasLegacyInsufficientRenderedGridProof(entry: {
  readonly scenarioProof: {
    readonly pixelGridProof: {
      readonly productVerdicts?: readonly {
        readonly evidenceStatus?: string
      }[]
    }
  }
}): boolean {
  return entry.scenarioProof.pixelGridProof.productVerdicts?.some((verdict) => verdict.evidenceStatus === 'legacy-insufficient') ?? false
}

function uiSameCorpusFixtureStatus(corpusCaseId: WorkbookBenchmarkCorpusId): UiSameCorpusFixtureStatus {
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  const localXlsxPath = `packages/benchmarks/baselines/ui-same-corpus/${corpus.id}.xlsx`
  const publicGithubRawUrl = `https://raw.githubusercontent.com/proompteng/bilig/main/${localXlsxPath}`
  return {
    corpusCaseId,
    materializedCells: corpus.materializedCellCount,
    localXlsxPath,
    publicGithubRawUrl,
    publicForgejoRawUrl: `https://code.proompteng.ai/kalmyk/bilig/raw/branch/main/${localXlsxPath}`,
    microsoftExcelWebUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(publicGithubRawUrl)}`,
  }
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function localCiResourceGuardOverrideCommand(command: string): string {
  if (command.includes(`${localCiResourceGuardOverrideEnv}=1`)) {
    return command
  }
  return `${localCiResourceGuardOverrideEnv}=1 ${command}`
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
