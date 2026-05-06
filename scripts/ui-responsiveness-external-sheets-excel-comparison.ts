import { arrayField, asObject, booleanField, stringArrayField, stringField } from './json-scorecard-helpers.ts'

export const externalUiResponsivenessComparisonArtifactRepoPath =
  'packages/benchmarks/baselines/ui-responsiveness-external-sheets-excel-comparison.json' as const

export const externalUiResponsivenessComparisonCoveredFeatures = [
  'external.googleSheetsUiResponsivenessDocs',
  'external.microsoftExcelUiResponsivenessDocs',
  'external.sheetsExcelUiResponsivenessComparison',
] as const

const requiredDimensionIds = [
  'visible-edit-and-calculation-responsiveness',
  'scroll-selection-filtering-responsiveness',
  'browser-workbook-resource-boundaries',
  'live-incumbent-ui-timing-gap',
] as const

type ExternalUiResponsivenessVendor = 'google-sheets' | 'microsoft-excel'
type ExternalUiResponsivenessDimensionId = (typeof requiredDimensionIds)[number]

export interface ExternalUiResponsivenessOfficialSource {
  readonly id: string
  readonly vendor: ExternalUiResponsivenessVendor
  readonly title: string
  readonly url: string
  readonly retrievedAt: string
  readonly evidence: string[]
}

export interface ExternalUiResponsivenessDimension {
  readonly id: ExternalUiResponsivenessDimensionId
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

export interface ExternalUiResponsivenessComparisonArtifact {
  readonly schemaVersion: 1
  readonly suite: 'external-sheets-excel-ui-responsiveness-comparison'
  readonly generatedAt: string
  readonly sourceBasis: string
  readonly officialSources: ExternalUiResponsivenessOfficialSource[]
  readonly summary: {
    readonly comparisonCaptured: boolean
    readonly requiredDimensionsPassed: boolean
    readonly coveredFeatures: string[]
    readonly limitations: string[]
  }
  readonly dimensions: ExternalUiResponsivenessDimension[]
}

export function parseExternalUiResponsivenessComparisonArtifact(value: unknown): ExternalUiResponsivenessComparisonArtifact {
  const record = asObject(value, 'external UI responsiveness comparison artifact')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'external-sheets-excel-ui-responsiveness-comparison') {
    throw new Error('Unexpected external UI responsiveness comparison artifact header')
  }
  const summary = asObject(record['summary'], 'external UI responsiveness comparison summary')

  return {
    schemaVersion: 1,
    suite: 'external-sheets-excel-ui-responsiveness-comparison',
    generatedAt: stringField(record, 'generatedAt'),
    sourceBasis: stringField(record, 'sourceBasis'),
    officialSources: arrayField(record, 'officialSources').map(parseOfficialSource),
    summary: {
      comparisonCaptured: booleanField(summary, 'comparisonCaptured'),
      requiredDimensionsPassed: booleanField(summary, 'requiredDimensionsPassed'),
      coveredFeatures: stringArrayField(summary, 'coveredFeatures'),
      limitations: stringArrayField(summary, 'limitations'),
    },
    dimensions: arrayField(record, 'dimensions').map(parseDimension),
  }
}

export function validateExternalUiResponsivenessComparisonArtifact(artifact: ExternalUiResponsivenessComparisonArtifact): string[] {
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
  if (!arrayEquals(artifact.summary.coveredFeatures, externalUiResponsivenessComparisonCoveredFeatures)) {
    findings.push('summary covered features do not match the external UI responsiveness comparison contract')
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
    if (!isValidOfficialSourceUrl(source.url)) {
      findings.push(`source ${source.id} is not an accepted official ${source.vendor} URL`)
    }
    if (!isIsoDate(source.retrievedAt)) {
      findings.push(`source ${source.id} retrievedAt must be YYYY-MM-DD`)
    }
  }

  for (const requiredDimensionId of requiredDimensionIds) {
    const dimension = artifact.dimensions.find((entry) => entry.id === requiredDimensionId)
    if (!dimension) {
      findings.push(`missing required external UI responsiveness dimension: ${requiredDimensionId}`)
      continue
    }
    findings.push(...validateDimension(dimension, artifact.officialSources, sourceIds))
  }

  return findings
}

function parseOfficialSource(value: unknown): ExternalUiResponsivenessOfficialSource {
  const record = asObject(value, 'external UI responsiveness official source')
  return {
    id: stringField(record, 'id'),
    vendor: parseVendor(stringField(record, 'vendor')),
    title: stringField(record, 'title'),
    url: stringField(record, 'url'),
    retrievedAt: stringField(record, 'retrievedAt'),
    evidence: stringArrayField(record, 'evidence'),
  }
}

function parseDimension(value: unknown): ExternalUiResponsivenessDimension {
  const record = asObject(value, 'external UI responsiveness dimension')
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
  dimension: ExternalUiResponsivenessDimension,
  officialSources: readonly ExternalUiResponsivenessOfficialSource[],
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
  if (dimension.residualLimitations.length === 0) {
    findings.push(`dimension ${dimension.id} must disclose residual limitations`)
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

function parseVendor(value: string): ExternalUiResponsivenessVendor {
  if (value === 'google-sheets' || value === 'microsoft-excel') {
    return value
  }
  throw new Error(`Unexpected external UI responsiveness vendor: ${value}`)
}

function parseDimensionId(value: string): ExternalUiResponsivenessDimensionId {
  switch (value) {
    case 'visible-edit-and-calculation-responsiveness':
    case 'scroll-selection-filtering-responsiveness':
    case 'browser-workbook-resource-boundaries':
    case 'live-incumbent-ui-timing-gap':
      return value
  }
  throw new Error(`Unexpected external UI responsiveness dimension: ${value}`)
}

function isValidOfficialSourceUrl(url: string): boolean {
  let parsed: URL
  try {
    parsed = new URL(url)
  } catch {
    return false
  }

  return parsed.protocol === 'https:' && (parsed.hostname === 'support.google.com' || parsed.hostname === 'support.microsoft.com')
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
