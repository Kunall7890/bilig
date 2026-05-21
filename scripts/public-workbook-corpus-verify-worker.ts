import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

import {
  tryInspectLargeSimpleXlsxHeadless,
  type LargeSimpleXlsxHeadlessInspectResult,
} from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import type { LargeSimpleXlsxImportStats } from '../packages/excel-import/src/xlsx-large-simple-import.js'
import {
  readXlsxZipEntryMetadata,
  readXlsxZipEntriesLazyFromByteSource,
  type XlsxZipByteSource,
} from '../packages/excel-import/src/xlsx-zip.js'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookFeatureCounts } from './public-workbook-corpus-types.ts'
import { defaultSelfRssCheckIntervalMs, startSelfRssGuard } from './public-workbook-corpus-process.ts'
import { largeSimpleImportPhaseTelemetryEvidence } from './public-workbook-corpus-large-simple-evidence.ts'
import { FileBackedXlsxZipByteSource, isZipWorkbookSource, sha256XlsxZipByteSourceHex } from './public-workbook-corpus-xlsx-byte-source.ts'

const verificationWorkerPhasePrefix = 'bilig-public-workbook-verify-phase='
const defaultVerifyTimeoutMs = 180_000
const defaultVerifyMaxRssBytes = 1536 * 1024 * 1024
const defaultVerifyMaxCellCount = 1_500_000
const compactLargeSimplePreflightMinPackageBytes = 2 * 1024 * 1024
const compactLargeSimplePreflightMinWorksheetCompressedBytes = 256 * 1024

const verifyMaxRssBytes = capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes))
const stopSelfRssGuard = startSelfRssGuard(verifyMaxRssBytes, 'Workbook verification worker')

try {
  const cacheDir = readStringArg('--cache-dir', '.cache/public-workbook-corpus')
  const artifactId = readStringArg('--artifact-id', '')
  if (!artifactId) {
    throw new Error('Expected --artifact-id for verify-artifact-worker')
  }
  const artifact = await readWorkerArtifact(artifactId)
  if (!artifact) {
    throw new Error(`Manifest does not contain public workbook artifact ${artifactId}`)
  }
  const runStructuralSmoke = readFlagArg('--structural-smoke')
  const maxCellCount = readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount)
  const workerOptions = {
    timeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
    maxRssBytes: verifyMaxRssBytes,
    rssCheckIntervalMs: defaultSelfRssCheckIntervalMs,
    onPhase: writeWorkerPhase,
  }
  const result =
    tryVerifyCompactLargeSimpleArtifact(artifact, cacheDir, runStructuralSmoke, maxCellCount) ??
    (await (async () => {
      const { verifyCachedWorkbookArtifact } = await import('./public-workbook-corpus-verify.ts')
      return verifyCachedWorkbookArtifact(artifact, cacheDir, runStructuralSmoke, maxCellCount, workerOptions)
    })())
  process.stdout.write(`${JSON.stringify(result)}\n`)
} finally {
  stopSelfRssGuard()
}

function tryVerifyCompactLargeSimpleArtifact(
  artifact: PublicWorkbookArtifact,
  cacheDir: string,
  runStructuralSmoke: boolean,
  maxCellCount: number,
): PublicWorkbookCorpusCase | null {
  const cachePath = join(cacheDir, artifact.cachePath)
  if (!existsSync(cachePath)) {
    return null
  }
  writeWorkerPhase('read-cache')
  const source = new FileBackedXlsxZipByteSource(cachePath)
  try {
    if (sha256XlsxZipByteSourceHex(source) !== artifact.sha256 || !isZipWorkbookSource(source)) {
      return null
    }
    collectGarbage()
    if (artifact.byteSize < compactLargeSimplePreflightMinPackageBytes && !hasCompactLargeSimpleWorksheetPayload(source)) {
      return null
    }
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      return null
    }
    writeWorkerPhase('import-xlsx')
    const imported = tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, artifact.fileName, zip, {
      afterWorksheetScan: collectGarbage,
      minByteLength: 0,
      releaseOwnedSourceBytes: () => ({ ownedSourceBytesBeforeRelease: source.byteLength, ownedSourceBytesAfterRelease: 0 }),
      releaseZipSource: true,
    })
    if (!imported) {
      return null
    }
    return buildCompactLargeSimpleCaseFromInspect(artifact, imported, runStructuralSmoke, maxCellCount)
  } finally {
    source.release()
  }
}

function hasCompactLargeSimpleWorksheetPayload(source: XlsxZipByteSource): boolean {
  const entries = readXlsxZipEntryMetadata(source)
  if (!entries) {
    return true
  }
  const worksheetCompressedBytes = entries.reduce(
    (sum, entry) => (/^xl\/worksheets\/[^/]+\.xml$/u.test(entry.path) ? sum + entry.compressedSize : sum),
    0,
  )
  return worksheetCompressedBytes >= compactLargeSimplePreflightMinWorksheetCompressedBytes
}

function buildCompactLargeSimpleCaseFromInspect(
  artifact: PublicWorkbookArtifact,
  imported: LargeSimpleXlsxHeadlessInspectResult,
  runStructuralSmoke: boolean,
  maxCellCount: number,
): PublicWorkbookCorpusCase | null {
  const featureCounts = featureCountsFromLargeSimpleStats(imported.stats)
  if (
    featureCounts.cellCount <= 100_000 ||
    featureCounts.cellCount > maxCellCount ||
    (featureCounts.formulaCellCount > 0 && featureCounts.formulaCellCount <= 2_000) ||
    !roundTripWouldBeResourceSkipped(artifact, featureCounts) ||
    (runStructuralSmoke && !structuralSmokeWouldBeResourceSkipped(featureCounts))
  ) {
    return null
  }
  const formulaOracleSkipped = featureCounts.formulaCellCount > 2_000
  const roundTripEvidence = `Round-trip projection skipped because workbook footprint exceeds verifier resource budget: cell-count ${String(
    featureCounts.cellCount,
  )} > 100000`
  const formulaOracleEvidence = `Formula oracle skipped because workbook has ${String(
    featureCounts.formulaCellCount,
  )} formulas, above verifier budget 2000.`
  const structuralSmokeEvidence = `Structural smoke skipped because workbook footprint exceeds verifier resource budget: cell-count ${String(
    featureCounts.cellCount,
  )} > 100000`
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status: 'unsupported',
    passed: true,
    featureCounts,
    workbookMetadata: {
      workbookName: imported.workbookName,
      sheetNames: imported.sheetNames,
      dimensions: imported.stats.dimensions.map((dimension) => ({
        sheetName: dimension.sheetName,
        rowCount: dimension.usedRange ? dimension.usedRange.endRow + 1 : dimension.rowCount,
        columnCount: dimension.usedRange ? dimension.usedRange.endColumn + 1 : dimension.columnCount,
        nonEmptyCellCount: dimension.nonEmptyCellCount,
        usedRange: dimension.usedRange,
      })),
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [
      ...(formulaOracleSkipped ? ['xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>2000formulas'] : []),
      'xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells',
      ...(runStructuralSmoke ? ['xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>100000cells'] : []),
    ],
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      `sheets=${String(featureCounts.sheetCount)}`,
      `cells=${String(featureCounts.cellCount)}`,
      `formulas=${String(featureCounts.formulaCellCount)}`,
      ...largeSimpleImportPhaseTelemetryEvidence(imported.stats),
      'resource-limit-classifier=2026-05-17-native-streaming-xlsx-footprint',
      ...(formulaOracleSkipped
        ? [
            'rss-limit-phase=formula-oracle',
            formulaOracleEvidence,
            `formula-oracle-formula-count=${String(featureCounts.formulaCellCount)}`,
          ]
        : []),
      'rss-limit-phase=round-trip',
      roundTripEvidence,
      ...(runStructuralSmoke ? ['rss-limit-phase=structural-smoke', structuralSmokeEvidence] : []),
    ],
  }
}

async function readWorkerArtifact(artifactId: string): Promise<PublicWorkbookArtifact | undefined> {
  const artifactJsonBase64 = readStringArg('--artifact-json-base64', '')
  if (artifactJsonBase64) {
    const artifact = parseWorkerArtifact(JSON.parse(Buffer.from(artifactJsonBase64, 'base64').toString('utf8')))
    if (artifact.id !== artifactId) {
      throw new Error(`Worker artifact id mismatch: expected ${artifactId}, received ${artifact.id}`)
    }
    return artifact
  }
  const manifestPath = readStringArg('--manifest', '.cache/public-workbook-corpus/manifest.json')
  const { parsePublicWorkbookManifestJson } = await import('./public-workbook-corpus-json.ts')
  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8')))
  return manifest.artifacts.find((entry) => entry.id === artifactId)
}

function readStringArg(name: string, fallback: string): string {
  // Keep worker argument parsing local so the low-RSS verifier does not load the broad public corpus CLI module.
  let value: string | null = null
  let count = 0
  for (const [index, arg] of process.argv.entries()) {
    const parsed = readArgValueForName(name, arg, index)
    if (parsed === null) {
      continue
    }
    count += 1
    assertArgSpecifiedOnce(name, count)
    value = parsed
  }
  return value ?? fallback
}

function readNumberArg(name: string, fallback: number): number {
  const raw = readStringArg(name, String(fallback))
  const parsed = Number(raw)
  if (!/^\d+$/u.test(raw) || parsed <= 0 || !Number.isSafeInteger(parsed)) {
    throw new Error(`Expected ${name} to be a positive integer`)
  }
  return parsed
}

function readMegabytesArg(name: string, fallbackBytes: number): number {
  const raw = readStringArg(name, String(Math.ceil(fallbackBytes / 1024 / 1024)))
  const parsed = Number(raw)
  if (!/^\d+$/u.test(raw) || parsed <= 0 || !Number.isSafeInteger(parsed) || parsed > Math.floor(Number.MAX_SAFE_INTEGER / 1024 / 1024)) {
    throw new Error(`Expected ${name} to be a positive integer number of MiB`)
  }
  return parsed * 1024 * 1024
}

function readFlagArg(name: string): boolean {
  let value = false
  let count = 0
  for (const [index, arg] of process.argv.entries()) {
    if (arg === name) {
      count += 1
      assertArgSpecifiedOnce(name, count)
      const next = process.argv[index + 1]
      value = next === undefined || next.startsWith('--') ? true : readBooleanArgValue(name, next)
      continue
    }
    const inlinePrefix = `${name}=`
    if (!arg.startsWith(inlinePrefix)) {
      continue
    }
    count += 1
    assertArgSpecifiedOnce(name, count)
    value = readBooleanArgValue(name, arg.slice(inlinePrefix.length))
  }
  return value
}

function assertArgSpecifiedOnce(name: string, count: number): void {
  if (count > 1) {
    throw new Error(`Expected ${name} to be specified once`)
  }
}

function readArgValueForName(name: string, arg: string, index: number): string | null {
  if (arg === name) {
    return readArgValueAt(name, index)
  }
  const inlinePrefix = `${name}=`
  if (!arg.startsWith(inlinePrefix)) {
    return null
  }
  const value = arg.slice(inlinePrefix.length)
  if (value.trim().length === 0) {
    throw new Error(`Expected ${name} to have a value`)
  }
  return value
}

function readArgValueAt(name: string, index: number): string {
  const next = process.argv[index + 1]
  if (next === undefined || next.trim().length === 0 || next.startsWith('--')) {
    throw new Error(`Expected ${name} to have a value`)
  }
  return next
}

function readBooleanArgValue(name: string, raw: string): boolean {
  if (raw === 'true') {
    return true
  }
  if (raw === 'false') {
    return false
  }
  throw new Error(`Expected ${name} to be true or false`)
}

function parseWorkerArtifact(value: unknown): PublicWorkbookArtifact {
  const record = readRecord(value)
  const topicEvidence = readOptionalStringArray(record, 'topicEvidence')
  return {
    id: readRequiredString(record, 'id'),
    sourceId: readRequiredString(record, 'sourceId'),
    sourceUrl: readRequiredString(record, 'sourceUrl'),
    downloadUrl: readRequiredString(record, 'downloadUrl'),
    fileName: readRequiredString(record, 'fileName'),
    cachePath: readRequiredString(record, 'cachePath'),
    sha256: readRequiredString(record, 'sha256'),
    byteSize: readRequiredInteger(record, 'byteSize'),
    workbookFingerprint: readRequiredString(record, 'workbookFingerprint'),
    fetchedAt: readRequiredString(record, 'fetchedAt'),
    license: parseWorkerLicense(record['license']),
    ...(topicEvidence ? { topicEvidence } : {}),
  }
}

function parseWorkerLicense(value: unknown): PublicWorkbookArtifact['license'] {
  const record = readRecord(value)
  return {
    spdxId: readNullableString(record, 'spdxId'),
    title: readRequiredString(record, 'title'),
    evidenceUrl: readNullableString(record, 'evidenceUrl'),
  }
}

function readRecord(value: unknown): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('Expected object')
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

function readRequiredString(record: Record<string, unknown>, key: string): string {
  const value = record[key]
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`Expected ${key} to be a non-empty string`)
  }
  return value.trim()
}

function readNullableString(record: Record<string, unknown>, key: string): string | null {
  const value = record[key]
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : null
}

function readRequiredInteger(record: Record<string, unknown>, key: string): number {
  const value = record[key]
  if (!Number.isInteger(value)) {
    throw new Error(`Expected ${key} to be an integer`)
  }
  return value
}

function readOptionalStringArray(record: Record<string, unknown>, key: string): readonly string[] | undefined {
  const value = record[key]
  if (value === undefined) {
    return undefined
  }
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected ${key} to be a string array`)
  }
  return value.map((entry) => entry.trim()).filter((entry) => entry.length > 0)
}

function writeWorkerPhase(phase: string): void {
  process.stderr.write(`${verificationWorkerPhasePrefix}${phase}\n`)
}

function capVerifyMaxRssBytes(value: number): number {
  const normalizedValue = Math.max(1, Math.trunc(value))
  if (normalizedValue > defaultVerifyMaxRssBytes) {
    throw new Error(
      `Public workbook corpus verification RSS limits above ${String(Math.ceil(defaultVerifyMaxRssBytes / 1024 / 1024))} MiB are disabled because workbook workers can hang interactive hosts.`,
    )
  }
  return normalizedValue
}

function collectGarbage(): void {
  if (typeof Bun !== 'undefined' && typeof Bun.gc === 'function') {
    Bun.gc(true)
    return
  }
  const gc = Reflect.get(globalThis, 'gc')
  if (typeof gc === 'function') {
    gc()
  }
}

function roundTripWouldBeResourceSkipped(artifact: PublicWorkbookArtifact, featureCounts: PublicWorkbookFeatureCounts): boolean {
  return featureCounts.cellCount > 100_000 || (featureCounts.sheetCount >= 30 && artifact.byteSize > 2 * 1024 * 1024)
}

function structuralSmokeWouldBeResourceSkipped(featureCounts: PublicWorkbookFeatureCounts): boolean {
  return featureCounts.cellCount > 100_000 || featureCounts.sheetCount > 80
}

function featureCountsFromLargeSimpleStats(stats: LargeSimpleXlsxImportStats): PublicWorkbookFeatureCounts {
  return {
    sheetCount: stats.sheetCount,
    cellCount: stats.cellCount,
    formulaCellCount: stats.formulaCellCount,
    valueCellCount: stats.valueCellCount,
    definedNameCount: stats.definedNameCount,
    tableCount: stats.tableCount,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: stats.mergeCount,
    styleRangeCount: 0,
    conditionalFormatCount: stats.conditionalFormatCount,
    dataValidationCount: stats.dataValidationCount ?? 0,
    macroPayloadCount: 0,
    warningCount: stats.warningCount,
  }
}
