interface PublicWorkbookCorpusScorecard {
  readonly generatedAt: string
  readonly summary: PublicWorkbookCorpusSummary
  readonly cases: readonly PublicWorkbookCorpusCase[]
}

interface PublicWorkbookCorpusSummary {
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedWorkbookCount: number
  readonly importedWorkbookCount: number
  readonly passedWorkbookCount: number
  readonly failedWorkbookCount: number
  readonly errorWorkbookCount: number
  readonly unsupportedWorkbookCount: number
  readonly formulaOracleComparisonCount: number
  readonly formulaOracleMatchCount: number
  readonly structuralSmokeRunCount: number
  readonly allCachedWorkbooksPassed: boolean
  readonly remainingToTarget: number
}

interface PublicWorkbookCorpusCase {
  readonly id: string
  readonly sourceUrl: string
  readonly fileName: string
  readonly byteSize: number
  readonly license: {
    readonly spdxId: string
    readonly title: string
    readonly evidenceUrl: string
  }
  readonly status: string
  readonly passed: boolean
  readonly featureCounts: PublicWorkbookCorpusFeatureCounts
  readonly unsupportedFeatureClassifications?: readonly string[] | undefined
}

export interface PublicWorkbookCorpusFeatureCounts {
  readonly sheetCount: number
  readonly cellCount: number
  readonly formulaCellCount: number
  readonly valueCellCount: number
  readonly definedNameCount: number
  readonly tableCount: number
  readonly chartCount: number
  readonly pivotCount: number
  readonly mergeCount: number
  readonly styleRangeCount: number
  readonly conditionalFormatCount: number
  readonly dataValidationCount: number
  readonly macroPayloadCount: number
  readonly warningCount: number
}

export interface PublicWorkbookCorpusLicenseEvidence {
  readonly spdxId: string
  readonly title: string
  readonly evidenceUrl: string
  readonly workbookCount: number
}

export interface PublicWorkbookCorpusWorkbookEvidence {
  readonly id: string
  readonly fileName: string
  readonly status: string
  readonly passed: boolean
  readonly sourceUrl: string
  readonly licenseSpdxId: string
  readonly byteSize: number
  readonly cellCount: number
  readonly formulaCellCount: number
  readonly sheetCount: number
  readonly unsupportedFeatureClassifications: readonly string[]
}

export interface PublicWorkbookCorpusEvidence {
  readonly artifactPath: string
  readonly reportPath: string
  readonly generatedAt: string
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedWorkbookCount: number
  readonly importedWorkbookCount: number
  readonly passedWorkbookCount: number
  readonly unsupportedWorkbookCount: number
  readonly failedWorkbookCount: number
  readonly errorWorkbookCount: number
  readonly formulaOracleComparisonCount: number
  readonly formulaOracleMatchCount: number
  readonly structuralSmokeRunCount: number
  readonly allCachedWorkbooksPassed: boolean
  readonly remainingToTarget: number
  readonly totalByteSize: number
  readonly featureTotals: PublicWorkbookCorpusFeatureCounts
  readonly licenseCount: number
  readonly licenses: readonly PublicWorkbookCorpusLicenseEvidence[]
  readonly largestWorkbooks: readonly PublicWorkbookCorpusWorkbookEvidence[]
  readonly unsupportedWorkbooks: readonly PublicWorkbookCorpusWorkbookEvidence[]
  readonly limitations: readonly string[]
}

const featureCountKeys = [
  'sheetCount',
  'cellCount',
  'formulaCellCount',
  'valueCellCount',
  'definedNameCount',
  'tableCount',
  'chartCount',
  'pivotCount',
  'mergeCount',
  'styleRangeCount',
  'conditionalFormatCount',
  'dataValidationCount',
  'macroPayloadCount',
  'warningCount',
] as const satisfies readonly (keyof PublicWorkbookCorpusFeatureCounts)[]

export function buildPublicWorkbookCorpusEvidence(args: {
  readonly scorecard: unknown
  readonly artifactPath: string
  readonly reportPath: string
}): PublicWorkbookCorpusEvidence {
  const scorecard = readScorecard(args.scorecard)
  const featureTotals = sumFeatureCounts(scorecard.cases)
  const licenses = summarizeLicenses(scorecard.cases)
  const workbookEvidence = scorecard.cases.map(toWorkbookEvidence)
  const summary = scorecard.summary

  return {
    artifactPath: args.artifactPath,
    reportPath: args.reportPath,
    generatedAt: scorecard.generatedAt,
    targetWorkbookCount: summary.targetWorkbookCount,
    sourceCount: summary.sourceCount,
    cachedWorkbookCount: summary.cachedWorkbookCount,
    importedWorkbookCount: summary.importedWorkbookCount,
    passedWorkbookCount: summary.passedWorkbookCount,
    unsupportedWorkbookCount: summary.unsupportedWorkbookCount,
    failedWorkbookCount: summary.failedWorkbookCount,
    errorWorkbookCount: summary.errorWorkbookCount,
    formulaOracleComparisonCount: summary.formulaOracleComparisonCount,
    formulaOracleMatchCount: summary.formulaOracleMatchCount,
    structuralSmokeRunCount: summary.structuralSmokeRunCount,
    allCachedWorkbooksPassed: summary.allCachedWorkbooksPassed,
    remainingToTarget: summary.remainingToTarget,
    totalByteSize: scorecard.cases.reduce((sum, entry) => sum + entry.byteSize, 0),
    featureTotals,
    licenseCount: licenses.length,
    licenses,
    largestWorkbooks: workbookEvidence.toSorted((left, right) => right.cellCount - left.cellCount).slice(0, 6),
    unsupportedWorkbooks: workbookEvidence.filter((entry) => entry.status === 'unsupported'),
    limitations: [
      'This is the checked-in 22-workbook public scorecard, not the broader 10,000-workbook or 5,000-financial-workbook objective.',
      'Unsupported rows are resource-budget classifications with evidence, not hidden failures.',
      'Formula-oracle matches only cover workbooks with meaningful cached formula comparisons in this scorecard.',
      'Run the commands above before using this report as current release evidence.',
    ],
  }
}

export function renderPublicWorkbookCorpusReport(evidence: PublicWorkbookCorpusEvidence): string {
  return `${frontMatter()}

# Public Workbook Corpus Report

Status: generated public evidence from \`${evidence.artifactPath}\`.

This report is intentionally scoped. It publishes the checked-in public workbook
corpus baseline that ships with the repository: \`${formatInteger(evidence.cachedWorkbookCount)}\` cached
public workbooks selected from \`${formatInteger(evidence.sourceCount)}\` source candidates. It is not the
larger active financial or 10,000-workbook corpus goal.

## Reproduce The Report

\`\`\`sh
pnpm public-workbook-corpus:status
pnpm public:evidence:check
\`\`\`

For a no-key package smoke outside a repo clone, run:

\`\`\`sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
\`\`\`

The package smoke proves the stale-XLSX recalculation door. The corpus commands
above prove this checked repository scorecard and generated report.

## Scorecard Summary

| Metric | Value |
| --- | ---: |
| Source candidates | \`${formatInteger(evidence.sourceCount)}\` |
| Cached public workbooks | \`${formatInteger(evidence.cachedWorkbookCount)}\` |
| Imported workbooks | \`${formatInteger(evidence.importedWorkbookCount)}\` |
| Passed workbooks | \`${formatInteger(evidence.passedWorkbookCount)}\` |
| Resource-limited unsupported workbooks | \`${formatInteger(evidence.unsupportedWorkbookCount)}\` |
| Failed workbooks | \`${formatInteger(evidence.failedWorkbookCount)}\` |
| Error workbooks | \`${formatInteger(evidence.errorWorkbookCount)}\` |
| Formula-oracle matches | \`${formatInteger(evidence.formulaOracleMatchCount)}/${formatInteger(evidence.formulaOracleComparisonCount)}\` |
| Structural smoke runs | \`${formatInteger(evidence.structuralSmokeRunCount)}\` |
| All cached workbooks passed gate | \`${String(evidence.allCachedWorkbooksPassed)}\` |
| Remaining to checked target | \`${formatInteger(evidence.remainingToTarget)}\` |
| Generated at | \`${evidence.generatedAt}\` |

## Workbook Shape

| Feature | Count |
| --- | ---: |
| Workbook bytes | \`${formatInteger(evidence.totalByteSize)}\` |
| Sheets | \`${formatInteger(evidence.featureTotals.sheetCount)}\` |
| Cells | \`${formatInteger(evidence.featureTotals.cellCount)}\` |
| Formula cells | \`${formatInteger(evidence.featureTotals.formulaCellCount)}\` |
| Defined names | \`${formatInteger(evidence.featureTotals.definedNameCount)}\` |
| Tables | \`${formatInteger(evidence.featureTotals.tableCount)}\` |
| Charts | \`${formatInteger(evidence.featureTotals.chartCount)}\` |
| Pivots | \`${formatInteger(evidence.featureTotals.pivotCount)}\` |
| Merges | \`${formatInteger(evidence.featureTotals.mergeCount)}\` |
| Conditional formats | \`${formatInteger(evidence.featureTotals.conditionalFormatCount)}\` |
| Warnings | \`${formatInteger(evidence.featureTotals.warningCount)}\` |

## Largest Cases

| Workbook | Status | Cells | Formula cells | License |
| --- | --- | ---: | ---: | --- |
${evidence.largestWorkbooks.map(formatWorkbookRow).join('\n')}

## Licenses

| SPDX | Workbooks | Evidence |
| --- | ---: | --- |
${evidence.licenses.map(formatLicenseRow).join('\n')}

## Resource-Limited Cases

The scorecard classifies \`${formatInteger(evidence.unsupportedWorkbookCount)}\` workbooks as unsupported
because their footprint exceeds the configured
round-trip or structural-smoke resource budget. Those rows still have source,
license, hash, import, and classification evidence.

| Workbook | Cells | Classification |
| --- | ---: | --- |
${evidence.unsupportedWorkbooks.map(formatUnsupportedRow).join('\n')}

## What This Proves

- The checked public baseline imports \`${formatInteger(evidence.importedWorkbookCount)}\` cached public workbook
  files without failures or errors.
- It covers \`${formatInteger(evidence.featureTotals.cellCount)}\` cells, \`${formatInteger(evidence.featureTotals.formulaCellCount)}\`
  formula cells, \`${formatInteger(evidence.featureTotals.sheetCount)}\` sheets, and \`${formatInteger(evidence.licenseCount)}\` source
  license families.
- Formula-oracle rows in this baseline match cached workbook evidence
  \`${formatInteger(evidence.formulaOracleMatchCount)}/${formatInteger(evidence.formulaOracleComparisonCount)}\`.
- Resource-budget skips are visible instead of being counted as silent passes.

## What This Does Not Prove

${evidence.limitations.map((limitation) => `- ${limitation}`).join('\n')}

Read [where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
before treating any corpus scorecard as a blanket Excel-parity claim. Use the
[XLSX corpus verifier walkthrough](xlsx-corpus-verifier-walkthrough.md) when
you need to run the same boundary against private workbooks.
`
}

function frontMatter(): string {
  return `---
title: Public workbook corpus report
published: true
description: Generated Bilig public workbook corpus scorecard with workbook counts, formula cells, license evidence, and explicit limitations.
tags: xlsx, public-data, formulas, compatibility, workpaper
canonical_url: https://proompteng.github.io/bilig/public-workbook-corpus-report.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---`
}

function readScorecard(value: unknown): PublicWorkbookCorpusScorecard {
  const record = asRecord(value, 'public workbook corpus scorecard')
  const cases = readArray(record, 'cases', 'public workbook corpus scorecard').map((entry, index) =>
    readCase(entry, `public workbook corpus scorecard.cases[${index.toString()}]`),
  )
  return {
    generatedAt: readString(record, 'generatedAt', 'public workbook corpus scorecard'),
    summary: readSummary(record['summary'], 'public workbook corpus scorecard.summary'),
    cases,
  }
}

function readSummary(value: unknown, context: string): PublicWorkbookCorpusSummary {
  const record = asRecord(value, context)
  return {
    targetWorkbookCount: readNumber(record, 'targetWorkbookCount', context),
    sourceCount: readNumber(record, 'sourceCount', context),
    cachedWorkbookCount: readNumber(record, 'cachedWorkbookCount', context),
    importedWorkbookCount: readNumber(record, 'importedWorkbookCount', context),
    passedWorkbookCount: readNumber(record, 'passedWorkbookCount', context),
    failedWorkbookCount: readNumber(record, 'failedWorkbookCount', context),
    errorWorkbookCount: readNumber(record, 'errorWorkbookCount', context),
    unsupportedWorkbookCount: readNumber(record, 'unsupportedWorkbookCount', context),
    formulaOracleComparisonCount: readNumber(record, 'formulaOracleComparisonCount', context),
    formulaOracleMatchCount: readNumber(record, 'formulaOracleMatchCount', context),
    structuralSmokeRunCount: readNumber(record, 'structuralSmokeRunCount', context),
    allCachedWorkbooksPassed: readBoolean(record, 'allCachedWorkbooksPassed', context),
    remainingToTarget: readNumber(record, 'remainingToTarget', context),
  }
}

function readCase(value: unknown, context: string): PublicWorkbookCorpusCase {
  const record = asRecord(value, context)
  const featureCounts = readFeatureCounts(record['featureCounts'], `${context}.featureCounts`)
  const license = asRecord(record['license'], `${context}.license`)
  return {
    id: readString(record, 'id', context),
    sourceUrl: readString(record, 'sourceUrl', context),
    fileName: readString(record, 'fileName', context),
    byteSize: readNumber(record, 'byteSize', context),
    license: {
      spdxId: readString(license, 'spdxId', `${context}.license`),
      title: readString(license, 'title', `${context}.license`),
      evidenceUrl: readString(license, 'evidenceUrl', `${context}.license`),
    },
    status: readString(record, 'status', context),
    passed: readBoolean(record, 'passed', context),
    featureCounts,
    unsupportedFeatureClassifications:
      record['unsupportedFeatureClassifications'] === undefined
        ? []
        : readStringArray(record, 'unsupportedFeatureClassifications', context),
  }
}

function readFeatureCounts(value: unknown, context: string): PublicWorkbookCorpusFeatureCounts {
  const record = asRecord(value, context)
  return {
    sheetCount: readNumber(record, 'sheetCount', context),
    cellCount: readNumber(record, 'cellCount', context),
    formulaCellCount: readNumber(record, 'formulaCellCount', context),
    valueCellCount: readNumber(record, 'valueCellCount', context),
    definedNameCount: readNumber(record, 'definedNameCount', context),
    tableCount: readNumber(record, 'tableCount', context),
    chartCount: readNumber(record, 'chartCount', context),
    pivotCount: readNumber(record, 'pivotCount', context),
    mergeCount: readNumber(record, 'mergeCount', context),
    styleRangeCount: readNumber(record, 'styleRangeCount', context),
    conditionalFormatCount: readNumber(record, 'conditionalFormatCount', context),
    dataValidationCount: readNumber(record, 'dataValidationCount', context),
    macroPayloadCount: readNumber(record, 'macroPayloadCount', context),
    warningCount: readNumber(record, 'warningCount', context),
  }
}

function sumFeatureCounts(cases: readonly PublicWorkbookCorpusCase[]): PublicWorkbookCorpusFeatureCounts {
  const totals = emptyFeatureCounts()
  for (const entry of cases) {
    for (const key of featureCountKeys) {
      totals[key] += entry.featureCounts[key]
    }
  }
  return totals
}

function emptyFeatureCounts(): PublicWorkbookCorpusFeatureCounts {
  return {
    sheetCount: 0,
    cellCount: 0,
    formulaCellCount: 0,
    valueCellCount: 0,
    definedNameCount: 0,
    tableCount: 0,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: 0,
    styleRangeCount: 0,
    conditionalFormatCount: 0,
    dataValidationCount: 0,
    macroPayloadCount: 0,
    warningCount: 0,
  }
}

function summarizeLicenses(cases: readonly PublicWorkbookCorpusCase[]): readonly PublicWorkbookCorpusLicenseEvidence[] {
  const licenses = new Map<string, PublicWorkbookCorpusLicenseEvidence>()
  for (const entry of cases) {
    const current = licenses.get(entry.license.spdxId)
    licenses.set(entry.license.spdxId, {
      spdxId: entry.license.spdxId,
      title: entry.license.title,
      evidenceUrl: entry.license.evidenceUrl,
      workbookCount: (current?.workbookCount ?? 0) + 1,
    })
  }
  return [...licenses.values()].toSorted((left, right) => left.spdxId.localeCompare(right.spdxId))
}

function toWorkbookEvidence(entry: PublicWorkbookCorpusCase): PublicWorkbookCorpusWorkbookEvidence {
  return {
    id: entry.id,
    fileName: entry.fileName,
    status: entry.status,
    passed: entry.passed,
    sourceUrl: entry.sourceUrl,
    licenseSpdxId: entry.license.spdxId,
    byteSize: entry.byteSize,
    cellCount: entry.featureCounts.cellCount,
    formulaCellCount: entry.featureCounts.formulaCellCount,
    sheetCount: entry.featureCounts.sheetCount,
    unsupportedFeatureClassifications: entry.unsupportedFeatureClassifications ?? [],
  }
}

function formatWorkbookRow(entry: PublicWorkbookCorpusWorkbookEvidence): string {
  return `| ${markdownCell(entry.fileName)} | \`${entry.status}\` | \`${formatInteger(entry.cellCount)}\` | \`${formatInteger(entry.formulaCellCount)}\` | \`${entry.licenseSpdxId}\` |`
}

function formatLicenseRow(entry: PublicWorkbookCorpusLicenseEvidence): string {
  return `| \`${entry.spdxId}\` | \`${formatInteger(entry.workbookCount)}\` | [${markdownCell(entry.title)}](${entry.evidenceUrl}) |`
}

function formatUnsupportedRow(entry: PublicWorkbookCorpusWorkbookEvidence): string {
  const classifications =
    entry.unsupportedFeatureClassifications.length === 0
      ? 'resource budget classification'
      : entry.unsupportedFeatureClassifications.map((classification) => `\`${classification}\``).join('<br>')
  return `| ${markdownCell(entry.fileName)} | \`${formatInteger(entry.cellCount)}\` | ${classifications} |`
}

function formatInteger(value: number): string {
  return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(value)
}

function markdownCell(value: string): string {
  return value.replace(/\|/gu, '\\|')
}

function asRecord(value: unknown, context: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${context} must be an object`)
  }
  return Object.fromEntries(Object.entries(value))
}

function readString(record: Record<string, unknown>, key: string, context: string): string {
  const value = record[key]
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`${context}.${key} must be a non-empty string`)
  }
  return value
}

function readNumber(record: Record<string, unknown>, key: string, context: string): number {
  const value = record[key]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`${context}.${key} must be a finite number`)
  }
  return value
}

function readBoolean(record: Record<string, unknown>, key: string, context: string): boolean {
  const value = record[key]
  if (typeof value !== 'boolean') {
    throw new Error(`${context}.${key} must be a boolean`)
  }
  return value
}

function readArray(record: Record<string, unknown>, key: string, context: string): readonly unknown[] {
  const value = record[key]
  if (!Array.isArray(value)) {
    throw new Error(`${context}.${key} must be an array`)
  }
  return value
}

function readStringArray(record: Record<string, unknown>, key: string, context: string): readonly string[] {
  const value = record[key]
  if (!Array.isArray(value) || value.some((entry) => typeof entry !== 'string')) {
    throw new Error(`${context}.${key} must be an array of strings`)
  }
  return [...value]
}
