import { arrayField, asObject, booleanField, stringArrayField, stringField } from './json-scorecard-helpers.ts'

export const externalSecurityComparisonArtifactRepoPath =
  'packages/benchmarks/baselines/security-external-sheets-excel-comparison.json' as const

export const externalSecurityComparisonCoveredControls = [
  'external.googleSheetsSecurityDocs',
  'external.microsoftExcelSecurityDocs',
  'external.sheetsExcelSecurityComparison',
] as const

const requiredDimensionIds = [
  'workbook-code-execution',
  'external-network-egress',
  'workbook-protection-boundaries',
  'enterprise-data-protection',
] as const

type ExternalSecurityVendor = 'google-sheets' | 'microsoft-excel'
type ExternalSecurityDimensionId = (typeof requiredDimensionIds)[number]

export interface ExternalSecurityOfficialSource {
  readonly id: string
  readonly vendor: ExternalSecurityVendor
  readonly title: string
  readonly url: string
  readonly retrievedAt: string
  readonly evidence: string[]
}

export interface ExternalSecurityDimension {
  readonly id: ExternalSecurityDimensionId
  readonly title: string
  readonly required: boolean
  readonly passed: boolean
  readonly sourceIds: string[]
  readonly biligEvidence: string[]
  readonly googleSheetsEvidence: string[]
  readonly microsoftExcelEvidence: string[]
  readonly comparisonResult: string
  readonly residualLimitations: string[]
}

export interface ExternalSecurityComparisonArtifact {
  readonly schemaVersion: 1
  readonly suite: 'external-sheets-excel-security-comparison'
  readonly generatedAt: string
  readonly sourceBasis: string
  readonly officialSources: ExternalSecurityOfficialSource[]
  readonly summary: {
    readonly comparisonCaptured: boolean
    readonly requiredDimensionsPassed: boolean
    readonly coveredControls: string[]
    readonly limitations: string[]
  }
  readonly dimensions: ExternalSecurityDimension[]
}

export function parseExternalSecurityComparisonArtifact(value: unknown): ExternalSecurityComparisonArtifact {
  const record = asObject(value, 'external security comparison artifact')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'external-sheets-excel-security-comparison') {
    throw new Error('Unexpected external security comparison artifact header')
  }
  const summary = asObject(record['summary'], 'external security comparison summary')

  return {
    schemaVersion: 1,
    suite: 'external-sheets-excel-security-comparison',
    generatedAt: stringField(record, 'generatedAt'),
    sourceBasis: stringField(record, 'sourceBasis'),
    officialSources: arrayField(record, 'officialSources').map(parseOfficialSource),
    summary: {
      comparisonCaptured: booleanField(summary, 'comparisonCaptured'),
      requiredDimensionsPassed: booleanField(summary, 'requiredDimensionsPassed'),
      coveredControls: stringArrayField(summary, 'coveredControls'),
      limitations: stringArrayField(summary, 'limitations'),
    },
    dimensions: arrayField(record, 'dimensions').map(parseDimension),
  }
}

export function validateExternalSecurityComparisonArtifact(artifact: ExternalSecurityComparisonArtifact): string[] {
  const findings: string[] = []
  const sourceIds = new Set(artifact.officialSources.map((source) => source.id))

  if (!isIsoDateTime(artifact.generatedAt)) {
    findings.push('generatedAt must be an ISO timestamp')
  }
  if (!/^official-public-docs-reviewed-\d{4}-\d{2}-\d{2}$/.test(artifact.sourceBasis)) {
    findings.push('sourceBasis must identify the official-doc review date')
  }
  if (!artifact.summary.comparisonCaptured) {
    findings.push('summary does not mark the comparison as captured')
  }
  if (!artifact.summary.requiredDimensionsPassed) {
    findings.push('summary does not mark required dimensions as passed')
  }
  if (!arrayEquals(artifact.summary.coveredControls, externalSecurityComparisonCoveredControls)) {
    findings.push('summary covered controls do not match the external comparison contract')
  }
  if (artifact.summary.limitations.length === 0) {
    findings.push('summary must disclose comparison limitations')
  }
  if (!artifact.officialSources.some((source) => source.vendor === 'google-sheets')) {
    findings.push('missing official Google Sheets source')
  }
  if (!artifact.officialSources.some((source) => source.vendor === 'microsoft-excel')) {
    findings.push('missing official Microsoft Excel source')
  }

  for (const source of artifact.officialSources) {
    if (source.evidence.length === 0) {
      findings.push(`source ${source.id} has no evidence summaries`)
    }
    if (!isValidOfficialSourceUrl(source.vendor, source.url)) {
      findings.push(`source ${source.id} is not an accepted official ${source.vendor} URL`)
    }
    if (!isIsoDate(source.retrievedAt)) {
      findings.push(`source ${source.id} retrievedAt must be YYYY-MM-DD`)
    }
  }

  for (const requiredDimensionId of requiredDimensionIds) {
    const dimension = artifact.dimensions.find((entry) => entry.id === requiredDimensionId)
    if (!dimension) {
      findings.push(`missing required external security dimension: ${requiredDimensionId}`)
      continue
    }
    findings.push(...validateDimension(dimension, artifact.officialSources, sourceIds))
  }

  return findings
}

function parseOfficialSource(value: unknown): ExternalSecurityOfficialSource {
  const record = asObject(value, 'external security official source')
  return {
    id: stringField(record, 'id'),
    vendor: parseVendor(stringField(record, 'vendor')),
    title: stringField(record, 'title'),
    url: stringField(record, 'url'),
    retrievedAt: stringField(record, 'retrievedAt'),
    evidence: stringArrayField(record, 'evidence'),
  }
}

function parseDimension(value: unknown): ExternalSecurityDimension {
  const record = asObject(value, 'external security dimension')
  return {
    id: parseDimensionId(stringField(record, 'id')),
    title: stringField(record, 'title'),
    required: booleanField(record, 'required'),
    passed: booleanField(record, 'passed'),
    sourceIds: stringArrayField(record, 'sourceIds'),
    biligEvidence: stringArrayField(record, 'biligEvidence'),
    googleSheetsEvidence: stringArrayField(record, 'googleSheetsEvidence'),
    microsoftExcelEvidence: stringArrayField(record, 'microsoftExcelEvidence'),
    comparisonResult: stringField(record, 'comparisonResult'),
    residualLimitations: stringArrayField(record, 'residualLimitations'),
  }
}

function validateDimension(
  dimension: ExternalSecurityDimension,
  officialSources: readonly ExternalSecurityOfficialSource[],
  sourceIds: ReadonlySet<string>,
): string[] {
  const findings: string[] = []
  const referencedSources = officialSources.filter((source) => dimension.sourceIds.includes(source.id))

  if (!dimension.required) {
    findings.push(`dimension ${dimension.id} is not marked required`)
  }
  if (!dimension.passed) {
    findings.push(`dimension ${dimension.id} is not marked passed`)
  }
  if (dimension.biligEvidence.length === 0) {
    findings.push(`dimension ${dimension.id} has no bilig evidence`)
  }
  if (dimension.googleSheetsEvidence.length === 0) {
    findings.push(`dimension ${dimension.id} has no Google Sheets evidence`)
  }
  if (dimension.microsoftExcelEvidence.length === 0) {
    findings.push(`dimension ${dimension.id} has no Microsoft Excel evidence`)
  }
  if (dimension.comparisonResult.trim().length === 0) {
    findings.push(`dimension ${dimension.id} has no comparison result`)
  }
  if (!referencedSources.some((source) => source.vendor === 'google-sheets')) {
    findings.push(`dimension ${dimension.id} does not cite an official Google Sheets source`)
  }
  if (!referencedSources.some((source) => source.vendor === 'microsoft-excel')) {
    findings.push(`dimension ${dimension.id} does not cite an official Microsoft Excel source`)
  }
  for (const sourceId of dimension.sourceIds) {
    if (!sourceIds.has(sourceId)) {
      findings.push(`dimension ${dimension.id} references unknown source ${sourceId}`)
    }
  }

  return findings
}

function parseVendor(value: string): ExternalSecurityVendor {
  if (value === 'google-sheets' || value === 'microsoft-excel') {
    return value
  }
  throw new Error(`Unexpected external security vendor: ${value}`)
}

function parseDimensionId(value: string): ExternalSecurityDimensionId {
  switch (value) {
    case 'workbook-code-execution':
    case 'external-network-egress':
    case 'workbook-protection-boundaries':
    case 'enterprise-data-protection':
      return value
  }
  throw new Error(`Unexpected external security dimension: ${value}`)
}

function isValidOfficialSourceUrl(vendor: ExternalSecurityVendor, url: string): boolean {
  let parsed: URL
  try {
    parsed = new URL(url)
  } catch {
    return false
  }

  if (parsed.protocol !== 'https:') {
    return false
  }
  if (vendor === 'google-sheets') {
    return parsed.hostname === 'support.google.com' || parsed.hostname === 'developers.google.com'
  }
  return parsed.hostname === 'support.microsoft.com' || parsed.hostname === 'learn.microsoft.com'
}

function isIsoDate(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(value)
}

function isIsoDateTime(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value)
}

function arrayEquals(left: readonly string[], right: readonly string[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index])
}
