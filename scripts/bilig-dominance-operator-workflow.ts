import { readFileSync } from 'node:fs'
import { join } from 'node:path'

export interface OperatorWorkflowEvidence {
  readonly packageJsonPath: string
  readonly runCiPath: string
  readonly auditGeneratorPath: string
  readonly scorecardGeneratorPath: string
  readonly dominanceGenerateScriptPresent: boolean
  readonly dominanceCheckScriptPresent: boolean
  readonly dominanceAuditCheckScriptPresent: boolean
  readonly googleSheetsTenXClaimGateScriptPresent: boolean
  readonly publicClaimsCheckScriptPresent: boolean
  readonly runCiDominanceCheckPresent: boolean
  readonly runCiDominanceAuditCheckPresent: boolean
  readonly runCiPublicClaimsCheckPresent: boolean
  readonly generatedSourceChecksSerialized: boolean
  readonly blanketClaimPolicyCoupledToCompletionAudit: boolean
  readonly promptArtifactAuditCoupledToLiveStatus: boolean
}

export function loadOperatorWorkflowEvidence(rootDir: string): OperatorWorkflowEvidence {
  const packageJsonPath = join(rootDir, 'package.json')
  const runCiPath = join(rootDir, 'scripts', 'run-ci.ts')
  const auditGeneratorPath = join(rootDir, 'scripts', 'bilig-dominance-audit.ts')
  const scorecardGeneratorPath = join(rootDir, 'scripts', 'gen-bilig-dominance-scorecard.ts')
  const packageJson = parsePackageJson(readFileSync(packageJsonPath, 'utf8'))
  const runCiSource = readFileSync(runCiPath, 'utf8')
  const auditGeneratorSource = readFileSync(auditGeneratorPath, 'utf8')
  const scorecardGeneratorSource = readFileSync(scorecardGeneratorPath, 'utf8')
  return {
    packageJsonPath: 'package.json',
    runCiPath: 'scripts/run-ci.ts',
    auditGeneratorPath: 'scripts/bilig-dominance-audit.ts',
    scorecardGeneratorPath: 'scripts/gen-bilig-dominance-scorecard.ts',
    dominanceGenerateScriptPresent: packageJson.scripts['dominance:generate'] === 'bun scripts/gen-bilig-dominance-scorecard.ts',
    dominanceCheckScriptPresent: packageJson.scripts['dominance:check'] === 'bun scripts/gen-bilig-dominance-scorecard.ts --check',
    dominanceAuditCheckScriptPresent: packageJson.scripts['dominance:audit:check'] === 'bun scripts/bilig-dominance-audit.ts --check',
    googleSheetsTenXClaimGateScriptPresent:
      packageJson.scripts['google-sheets-10x:claim:check'] === 'bun scripts/google-sheets-10x-claim-gate.ts',
    publicClaimsCheckScriptPresent: packageJson.scripts['claims:check'] === 'bun scripts/check-public-claims.ts',
    runCiDominanceCheckPresent: runCiSource.includes(
      "bunScript('bilig dominance scorecard check', 'scripts/gen-bilig-dominance-scorecard.ts', '--check')",
    ),
    runCiDominanceAuditCheckPresent: runCiSource.includes(
      "bunScript('bilig dominance audit check', 'scripts/bilig-dominance-audit.ts', '--check')",
    ),
    runCiPublicClaimsCheckPresent: runCiSource.includes("bunScript('public claims check', 'scripts/check-public-claims.ts')"),
    generatedSourceChecksSerialized: runCiSource.includes('Keep generated-source checks serialized'),
    blanketClaimPolicyCoupledToCompletionAudit:
      scorecardGeneratorSource.includes("goalStatus: completionAudit.allCriteriaPassed ? 'achieved' : 'active-not-achieved'") &&
      scorecardGeneratorSource.includes('blanketTenXClaimAllowed: completionAudit.allCriteriaPassed') &&
      scorecardGeneratorSource.includes('unmetRequirements: completionAudit.unmetRequirements'),
    promptArtifactAuditCoupledToLiveStatus:
      auditGeneratorSource.includes('buildBiligDominanceStatusFromArgs') &&
      auditGeneratorSource.includes('livePublicWorkbookCorpus') &&
      auditGeneratorSource.includes('validateBiligDominancePromptArtifactAudit'),
  }
}

export function operatorWorkflowGaps(evidence: OperatorWorkflowEvidence): string[] {
  return [
    ...(evidence.dominanceGenerateScriptPresent ? [] : ['package.json is missing the dominance:generate script']),
    ...(evidence.dominanceCheckScriptPresent ? [] : ['package.json is missing the dominance:check script']),
    ...(evidence.dominanceAuditCheckScriptPresent ? [] : ['package.json is missing the dominance:audit:check script']),
    ...(evidence.googleSheetsTenXClaimGateScriptPresent ? [] : ['package.json is missing the google-sheets-10x:claim:check script']),
    ...(evidence.publicClaimsCheckScriptPresent ? [] : ['package.json is missing the claims:check script']),
    ...(evidence.runCiDominanceCheckPresent ? [] : ['run-ci does not execute dominance:check']),
    ...(evidence.runCiDominanceAuditCheckPresent ? [] : ['run-ci does not execute dominance:audit:check']),
    ...(evidence.runCiPublicClaimsCheckPresent ? [] : ['run-ci does not execute claims:check']),
    ...(evidence.generatedSourceChecksSerialized ? [] : ['generated-source CI checks are not serialized']),
    ...(evidence.blanketClaimPolicyCoupledToCompletionAudit ? [] : ['blanket 10x claim policy is not coupled to completion audit results']),
    ...(evidence.promptArtifactAuditCoupledToLiveStatus ? [] : ['prompt-to-artifact audit is not coupled to live dominance status']),
  ]
}

function parsePackageJson(source: string): { readonly scripts: Record<string, string> } {
  const parsed: unknown = JSON.parse(source)
  if (!isRecord(parsed)) {
    throw new Error('package.json must be an object')
  }
  const scripts = parsed['scripts']
  if (!isRecord(scripts)) {
    throw new Error('package.json scripts must be an object')
  }
  return {
    scripts: Object.fromEntries(Object.entries(scripts).filter((entry): entry is [string, string] => typeof entry[1] === 'string')),
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}
