import { join, resolve } from 'node:path'

import type { BuildScorecardInput } from './bilig-dominance-scorecard-types.ts'
import { parseCompetitiveArtifact, parseFormulaDominanceSnapshot, parseSurfaceSnapshot } from './bilig-dominance-scorecard-parsers.ts'
import { parseAuditabilityScorecard } from './gen-auditability-scorecard.ts'
import { parseAutomationScorecard } from './gen-automation-scorecard.ts'
import { parseCalculationSemanticsScorecard } from './gen-calculation-semantics-scorecard.ts'
import { parseCollaborationScorecard } from './gen-collaboration-scorecard.ts'
import { parseGoogleSheetsLiveCalculationScorecard } from './gen-google-sheets-live-calculation-scorecard.ts'
import { parseGoogleSheetsLiveLargeWorkbookScorecard } from './gen-google-sheets-live-large-workbook-scorecard.ts'
import { parseGoogleSheetsLiveRecalculationScorecard } from './gen-google-sheets-live-recalculation-scorecard.ts'
import { parseGoogleSheetsLiveStructuralScorecard } from './gen-google-sheets-live-structural-scorecard.ts'
import { parseImportExportFidelityScorecard } from './gen-import-export-fidelity-scorecard.ts'
import { parseLargeWorkbookSloScorecard } from './gen-large-workbook-slo-scorecard.ts'
import { parseMicrosoftExcelLiveCalculationScorecard } from './gen-microsoft-excel-live-calculation-scorecard.ts'
import { parseMicrosoftExcelLiveLargeWorkbookScorecard } from './gen-microsoft-excel-live-large-workbook-scorecard.ts'
import { parseMicrosoftExcelLiveRecalculationScorecard } from './gen-microsoft-excel-live-recalculation-scorecard.ts'
import { parseMicrosoftExcelLiveStructuralScorecard } from './gen-microsoft-excel-live-structural-scorecard.ts'
import { parseReliabilityScorecard } from './gen-reliability-scorecard.ts'
import { parseSecurityPostureScorecard } from './gen-security-posture-scorecard.ts'
import { parseUiResponsivenessLiveBrowserScorecard } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { readJsonObject } from './json-scorecard-helpers.ts'

export const rootDir = resolve(new URL('..', import.meta.url).pathname)
export const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'bilig-dominance-scorecard.json')

const auditabilityScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'auditability-scorecard.json')
const automationScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'automation-scorecard.json')
const calculationSemanticsScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'calculation-semantics-scorecard.json')
const collaborationScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'collaboration-scorecard.json')
const competitiveArtifactPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json')
const formulaSnapshotPath = join(rootDir, 'packages', 'formula', 'src', '__tests__', 'fixtures', 'formula-dominance-snapshot.json')
const googleSheetsLiveCalculationScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'google-sheets-live-calculation-scorecard.json',
)
const googleSheetsLiveRecalculationScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'google-sheets-live-recalculation-scorecard.json',
)
const googleSheetsLiveStructuralScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'google-sheets-live-structural-scorecard.json',
)
const googleSheetsLiveLargeWorkbookScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'google-sheets-live-large-workbook-scorecard.json',
)
const microsoftExcelLiveCalculationScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'microsoft-excel-live-calculation-scorecard.json',
)
const microsoftExcelLiveRecalculationScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'microsoft-excel-live-recalculation-scorecard.json',
)
const microsoftExcelLiveLargeWorkbookScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'microsoft-excel-live-large-workbook-scorecard.json',
)
const microsoftExcelLiveStructuralScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'microsoft-excel-live-structural-scorecard.json',
)
const importExportFidelityScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'import-export-fidelity-scorecard.json')
const largeWorkbookSloScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'large-workbook-slo-scorecard.json')
const uiResponsivenessLiveBrowserScorecardPath = join(
  rootDir,
  'packages',
  'benchmarks',
  'baselines',
  'ui-responsiveness-live-browser-scorecard.json',
)
const reliabilityScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'reliability-scorecard.json')
const securityPostureScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'security-posture-scorecard.json')
const surfaceSnapshotPath = join(rootDir, 'packages', 'headless', 'src', '__tests__', 'fixtures', 'hyperformula-surface.json')

export function loadBiligDominanceScorecardInput(): BuildScorecardInput {
  return {
    auditabilityScorecard: parseAuditabilityScorecard(readJsonObject(auditabilityScorecardPath)),
    auditabilityScorecardPath: toRepoPath(auditabilityScorecardPath),
    automationScorecard: parseAutomationScorecard(readJsonObject(automationScorecardPath)),
    automationScorecardPath: toRepoPath(automationScorecardPath),
    calculationSemanticsScorecard: parseCalculationSemanticsScorecard(readJsonObject(calculationSemanticsScorecardPath)),
    calculationSemanticsScorecardPath: toRepoPath(calculationSemanticsScorecardPath),
    collaborationScorecard: parseCollaborationScorecard(readJsonObject(collaborationScorecardPath)),
    collaborationScorecardPath: toRepoPath(collaborationScorecardPath),
    competitiveArtifact: parseCompetitiveArtifact(readJsonObject(competitiveArtifactPath)),
    competitiveArtifactPath: toRepoPath(competitiveArtifactPath),
    formulaSnapshot: parseFormulaDominanceSnapshot(readJsonObject(formulaSnapshotPath)),
    formulaSnapshotPath: toRepoPath(formulaSnapshotPath),
    googleSheetsLiveCalculationScorecard: parseGoogleSheetsLiveCalculationScorecard(
      readJsonObject(googleSheetsLiveCalculationScorecardPath),
    ),
    googleSheetsLiveCalculationScorecardPath: toRepoPath(googleSheetsLiveCalculationScorecardPath),
    googleSheetsLiveRecalculationScorecard: parseGoogleSheetsLiveRecalculationScorecard(
      readJsonObject(googleSheetsLiveRecalculationScorecardPath),
    ),
    googleSheetsLiveRecalculationScorecardPath: toRepoPath(googleSheetsLiveRecalculationScorecardPath),
    googleSheetsLiveStructuralScorecard: parseGoogleSheetsLiveStructuralScorecard(readJsonObject(googleSheetsLiveStructuralScorecardPath)),
    googleSheetsLiveStructuralScorecardPath: toRepoPath(googleSheetsLiveStructuralScorecardPath),
    googleSheetsLiveLargeWorkbookScorecard: parseGoogleSheetsLiveLargeWorkbookScorecard(
      readJsonObject(googleSheetsLiveLargeWorkbookScorecardPath),
    ),
    googleSheetsLiveLargeWorkbookScorecardPath: toRepoPath(googleSheetsLiveLargeWorkbookScorecardPath),
    microsoftExcelLiveCalculationScorecard: parseMicrosoftExcelLiveCalculationScorecard(
      readJsonObject(microsoftExcelLiveCalculationScorecardPath),
    ),
    microsoftExcelLiveCalculationScorecardPath: toRepoPath(microsoftExcelLiveCalculationScorecardPath),
    microsoftExcelLiveRecalculationScorecard: parseMicrosoftExcelLiveRecalculationScorecard(
      readJsonObject(microsoftExcelLiveRecalculationScorecardPath),
    ),
    microsoftExcelLiveRecalculationScorecardPath: toRepoPath(microsoftExcelLiveRecalculationScorecardPath),
    microsoftExcelLiveLargeWorkbookScorecard: parseMicrosoftExcelLiveLargeWorkbookScorecard(
      readJsonObject(microsoftExcelLiveLargeWorkbookScorecardPath),
    ),
    microsoftExcelLiveLargeWorkbookScorecardPath: toRepoPath(microsoftExcelLiveLargeWorkbookScorecardPath),
    microsoftExcelLiveStructuralScorecard: parseMicrosoftExcelLiveStructuralScorecard(
      readJsonObject(microsoftExcelLiveStructuralScorecardPath),
    ),
    microsoftExcelLiveStructuralScorecardPath: toRepoPath(microsoftExcelLiveStructuralScorecardPath),
    importExportFidelityScorecard: parseImportExportFidelityScorecard(readJsonObject(importExportFidelityScorecardPath)),
    importExportFidelityScorecardPath: toRepoPath(importExportFidelityScorecardPath),
    largeWorkbookSloScorecard: parseLargeWorkbookSloScorecard(readJsonObject(largeWorkbookSloScorecardPath)),
    largeWorkbookSloScorecardPath: toRepoPath(largeWorkbookSloScorecardPath),
    uiResponsivenessLiveBrowserScorecard: parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(uiResponsivenessLiveBrowserScorecardPath),
    ),
    uiResponsivenessLiveBrowserScorecardPath: toRepoPath(uiResponsivenessLiveBrowserScorecardPath),
    reliabilityScorecard: parseReliabilityScorecard(readJsonObject(reliabilityScorecardPath)),
    reliabilityScorecardPath: toRepoPath(reliabilityScorecardPath),
    securityPostureScorecard: parseSecurityPostureScorecard(readJsonObject(securityPostureScorecardPath)),
    securityPostureScorecardPath: toRepoPath(securityPostureScorecardPath),
    surfaceSnapshot: parseSurfaceSnapshot(readJsonObject(surfaceSnapshotPath)),
    surfaceSnapshotPath: toRepoPath(surfaceSnapshotPath),
  }
}

function toRepoPath(path: string): string {
  return path.startsWith(`${rootDir}/`) ? path.slice(rootDir.length + 1) : path
}
