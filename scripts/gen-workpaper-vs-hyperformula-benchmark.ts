#!/usr/bin/env bun

import { execFileSync, spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula.ts'
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  buildExpandedComparativeBenchmarkReport,
  parseExpandedBenchmarkCliOptions,
  runWorkPaperVsHyperFormulaExpandedBenchmarkSuite,
  type ExpandedComparativeBenchmarkWorkload,
  type ExpandedComparativeBenchmarkResult,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts'
import {
  buildExpandedCompetitiveFamilyReport,
  type ExpandedCompetitiveFamilySummary,
  type ExpandedCompetitiveScorecardSummary,
} from '../packages/benchmarks/src/report-competitive-families.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

interface ExpandedCompetitiveBenchmarkArtifact {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula'
  generatedAt: string
  host: {
    arch: string
    nodeVersion: string
    platform: string
  }
  benchmark: {
    sampleCount: number
    warmupCount: number
  }
  engines: {
    hyperformula: {
      commit: string
      licenseKey: string
      metadataSource: 'fallback' | 'local-checkout'
      packageName: 'hyperformula'
      sourcePath: string
      version: string
    }
    workpaper: {
      packageName: '@bilig/headless'
      sourcePath: string
      version: string
    }
  }
  families: readonly ExpandedCompetitiveFamilySummary[]
  scorecard: ExpandedCompetitiveScorecardSummary
  results: ExpandedComparativeBenchmarkResult[]
}

interface ArtifactShapeInput {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula'
  results: Array<{
    category: string
    comparable: boolean
    comparison?: Record<string, unknown>
    engines: {
      hyperformula: Record<string, unknown>
      workpaper: Record<string, unknown>
    }
    fixture: Record<string, unknown>
    note?: string
    workload: string
  }>
}

interface ArtifactReportInput {
  schemaVersion: 1
  suite: 'workpaper-vs-hyperformula'
  families: readonly ExpandedCompetitiveFamilySummary[]
  scorecard: ExpandedCompetitiveScorecardSummary
  results: ExpandedComparativeBenchmarkResult[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json')
const localHyperFormulaRoot = '/Users/gregkonush/github.com/hyperformula'
const rawCliArgs = process.argv.slice(2)
const isCheckMode = rawCliArgs.includes('--check')
const isEmitResultsMode = rawCliArgs.includes('--emit-results')
const benchmarkCliOptions = parseExpandedBenchmarkCliOptions(rawCliArgs.filter((arg) => arg !== '--check' && arg !== '--emit-results'))
const workpaperSourcePath = 'packages/headless'
const expandedComparativeWorkloadSet: ReadonlySet<string> = new Set(EXPANDED_COMPARATIVE_WORKLOADS)

const sampleCount = benchmarkCliOptions.sampleCount ?? DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT
const warmupCount = benchmarkCliOptions.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT
const jobs = benchmarkCliOptions.jobs ?? 1

if (isCheckMode) {
  if (benchmarkCliOptions.workloads !== undefined) {
    throw new Error('--workload is only supported when generating benchmark worker results')
  }
  if (!existsSync(outputPath)) {
    throw new Error('WorkPaper competitive benchmark artifact is missing. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts')
  }

  const artifactJson = readFileSync(outputPath, 'utf8')
  const artifactRecord = parseJsonRecord(artifactJson)
  assertEngineSourcePath(artifactRecord, 'workpaper', workpaperSourcePath)
  assertBenchmarkSettings(artifactRecord, {
    sampleCount,
    warmupCount,
  })
  assertNoRawBenchmarkSampleArrays(artifactRecord)
  const existing = parseArtifactForShape(artifactJson)
  const actualShape = normalizeArtifactShape(existing)
  const actualWorkloads = actualShape.workloads.map((workload) => workload.workload)
  if (JSON.stringify(actualWorkloads) !== JSON.stringify([...EXPANDED_COMPARATIVE_WORKLOADS])) {
    throw new Error(
      'WorkPaper competitive benchmark artifact workload coverage is out of date. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts',
    )
  }
  const expectedShape = normalizeArtifactShape({
    schemaVersion: 1,
    suite: 'workpaper-vs-hyperformula',
    results: [...EXPANDED_COMPARATIVE_WORKLOADS].map((workload) =>
      isLeadershipWorkload(workload) ? leadershipShape(workload, [], []) : comparableShape(workload, [], []),
    ),
  })

  if (JSON.stringify(actualShape) !== JSON.stringify(expectedShape)) {
    throw new Error(
      'WorkPaper competitive benchmark artifact shape is out of date. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts',
    )
  }

  assertArtifactReportDerivedFromResults(parseArtifactForReport(artifactJson))

  console.log(
    JSON.stringify(
      {
        mode: 'check',
        outputPath,
        workloads: actualShape.workloads.map((workload) => workload.workload),
      },
      null,
      2,
    ),
  )
  process.exit(0)
}

if (!isEmitResultsMode && benchmarkCliOptions.workloads !== undefined) {
  throw new Error('--workload is only supported when generating benchmark worker results')
}

const workpaperVersion = readPackageVersion(join(rootDir, 'packages', 'headless', 'package.json'))
const hyperformulaMetadata = readHyperFormulaMetadata(localHyperFormulaRoot)
const results = await runBenchmarkResults({
  jobs,
  sampleCount,
  warmupCount,
  ...(benchmarkCliOptions.workloads ? { workloads: benchmarkCliOptions.workloads } : {}),
})

if (isEmitResultsMode) {
  const workerResultsPath = process.env['BILIG_BENCHMARK_WORKER_RESULTS_PATH']
  if (workerResultsPath) {
    mkdirSync(dirname(workerResultsPath), { recursive: true })
    writeFileSync(workerResultsPath, JSON.stringify(results))
    console.log(JSON.stringify({ resultsPath: workerResultsPath }))
  } else {
    console.log(JSON.stringify(results))
  }
  process.exit(0)
}

const benchmarkReport = buildExpandedComparativeBenchmarkReport(results)

const artifact: ExpandedCompetitiveBenchmarkArtifact = {
  schemaVersion: 1,
  suite: 'workpaper-vs-hyperformula',
  generatedAt: new Date().toISOString(),
  host: {
    arch: process.arch,
    nodeVersion: process.version,
    platform: process.platform,
  },
  benchmark: {
    sampleCount,
    warmupCount,
  },
  engines: {
    workpaper: {
      packageName: '@bilig/headless',
      sourcePath: workpaperSourcePath,
      version: workpaperVersion,
    },
    hyperformula: hyperformulaMetadata,
  },
  families: benchmarkReport.families,
  scorecard: benchmarkReport.scorecard,
  results,
}

mkdirSync(dirname(outputPath), { recursive: true })
writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(compactRawBenchmarkSampleArrays(artifact), null, 2)}\n`))
console.log(
  JSON.stringify(
    {
      mode: 'write',
      outputPath,
      workloads: artifact.results.map((result) => result.workload),
    },
    null,
    2,
  ),
)

async function runBenchmarkResults(options: {
  jobs: number
  sampleCount: number
  warmupCount: number
  workloads?: readonly ExpandedComparativeBenchmarkWorkload[]
}): Promise<ExpandedComparativeBenchmarkResult[]> {
  const workloads = options.workloads ?? EXPANDED_COMPARATIVE_WORKLOADS
  if (workloads.length === 0) {
    return []
  }
  if (options.jobs === 1 || workloads.length === 1) {
    return runWorkPaperVsHyperFormulaExpandedBenchmarkSuite({
      sampleCount: options.sampleCount,
      warmupCount: options.warmupCount,
      workloads,
    })
  }

  const shards = shardWorkloads(workloads, Math.min(options.jobs, workloads.length))
  const shardResults = await Promise.all(
    shards.map((shard, index) =>
      runBenchmarkWorker({
        shardIndex: index,
        sampleCount: options.sampleCount,
        warmupCount: options.warmupCount,
        workloads: shard,
      }),
    ),
  )
  return orderBenchmarkResults(shardResults.flat(), workloads)
}

function shardWorkloads(
  workloads: readonly ExpandedComparativeBenchmarkWorkload[],
  shardCount: number,
): readonly (readonly ExpandedComparativeBenchmarkWorkload[])[] {
  const shards: ExpandedComparativeBenchmarkWorkload[][] = Array.from({ length: shardCount }, () => [])
  workloads.forEach((workload, index) => {
    shards[index % shardCount].push(workload)
  })
  return shards.filter((shard) => shard.length > 0)
}

async function runBenchmarkWorker(options: {
  shardIndex: number
  sampleCount: number
  warmupCount: number
  workloads: readonly ExpandedComparativeBenchmarkWorkload[]
}): Promise<ExpandedComparativeBenchmarkResult[]> {
  const scriptPath = fileURLToPath(import.meta.url)
  const args = [
    scriptPath,
    '--emit-results',
    '--sample-count',
    String(options.sampleCount),
    '--warmup-count',
    String(options.warmupCount),
    ...options.workloads.flatMap((workload) => ['--workload', workload]),
  ]
  const resultsPath = join(rootDir, '.cache', 'benchmark-workers', `workpaper-vs-hyperformula-${process.pid}-${options.shardIndex}.json`)
  const { stdout } = await spawnForOutput(process.execPath, args, {
    BILIG_BENCHMARK_WORKER_RESULTS_PATH: resultsPath,
  })
  const workerMessage = JSON.parse(stdout) as unknown
  const parsed =
    isRecord(workerMessage) && workerMessage['resultsPath'] === resultsPath
      ? (JSON.parse(readFileSync(resultsPath, 'utf8')) as unknown)
      : workerMessage
  if (!Array.isArray(parsed)) {
    throw new Error('Benchmark worker returned a non-array result payload')
  }
  const workerResults: ExpandedComparativeBenchmarkResult[] = []
  for (const value of parsed) {
    if (!isExpandedComparativeBenchmarkResult(value)) {
      throw new Error('Benchmark worker returned an invalid result payload')
    }
    workerResults.push(value)
  }
  return workerResults
}

function spawnForOutput(command: string, args: readonly string[], env: Readonly<Record<string, string>> = {}): Promise<{ stdout: string }> {
  return new Promise((resolvePromise, reject) => {
    const child = spawn(command, args, {
      cwd: rootDir,
      env: {
        ...process.env,
        ...env,
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
    })
    child.on('error', reject)
    child.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(`Benchmark worker exited with code ${String(code)}${stderr ? `:\n${stderr}` : ''}`))
        return
      }
      resolvePromise({ stdout })
    })
  })
}

function orderBenchmarkResults(
  benchmarkResults: readonly ExpandedComparativeBenchmarkResult[],
  expectedWorkloads: readonly ExpandedComparativeBenchmarkWorkload[],
): ExpandedComparativeBenchmarkResult[] {
  const resultsByWorkload = new Map<ExpandedComparativeBenchmarkWorkload, ExpandedComparativeBenchmarkResult>()
  for (const result of benchmarkResults) {
    if (resultsByWorkload.has(result.workload)) {
      throw new Error(`Benchmark worker returned duplicate workload: ${result.workload}`)
    }
    resultsByWorkload.set(result.workload, result)
  }
  return expectedWorkloads.map((workload) => {
    const result = resultsByWorkload.get(workload)
    if (!result) {
      throw new Error(`Benchmark worker did not return workload: ${workload}`)
    }
    return result
  })
}

function isExpandedComparativeBenchmarkResult(value: unknown): value is ExpandedComparativeBenchmarkResult {
  if (!isRecord(value)) {
    return false
  }
  const workload = value.workload
  const category = value.category
  return (
    typeof workload === 'string' &&
    isExpandedComparativeBenchmarkWorkload(workload) &&
    (category === 'directly-comparable' || category === 'leadership')
  )
}

function isExpandedComparativeBenchmarkWorkload(value: string): value is ExpandedComparativeBenchmarkWorkload {
  return expandedComparativeWorkloadSet.has(value)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function comparableShape(workload: string, fixtureKeys: string[], verificationKeys: string[]): ArtifactShapeInput['results'][number] {
  return {
    workload,
    category: 'directly-comparable',
    comparable: true,
    fixture: Object.fromEntries(fixtureKeys.map((key) => [key, 'placeholder'])),
    comparison: {
      fasterEngine: 'workpaper',
      meanSpeedup: 1,
      workpaperToHyperFormulaMeanRatio: 1,
      workpaperToHyperFormulaMedianRatio: 1,
      workpaperToHyperFormulaP95Ratio: 1,
      maxRelativeNoise: 0,
      confidenceIntervalOverlaps: false,
      resultConfidence: 'decisive',
      decisiveFasterEngine: 'workpaper',
      verificationEquivalent: true,
    },
    engines: {
      workpaper: measuredEngineShape(verificationKeys),
      hyperformula: measuredEngineShape(verificationKeys),
    },
  }
}

function leadershipShape(workload: string, fixtureKeys: string[], verificationKeys: string[]): ArtifactShapeInput['results'][number] {
  return {
    workload,
    category: 'leadership',
    comparable: false,
    fixture: Object.fromEntries(fixtureKeys.map((key) => [key, 'placeholder'])),
    note: 'placeholder',
    engines: {
      workpaper: measuredEngineShape(verificationKeys),
      hyperformula: {
        evidence: [],
        reason: '',
        status: 'unsupported',
      },
    },
  }
}

function measuredEngineShape(verificationKeys: string[]): Record<string, unknown> {
  return {
    status: 'supported',
    elapsedMs: numericSummaryShape(),
    memoryDeltaBytes: {
      arrayBuffersBytes: numericSummaryShape(),
      externalBytes: numericSummaryShape(),
      heapTotalBytes: numericSummaryShape(),
      heapUsedBytes: numericSummaryShape(),
      rssBytes: numericSummaryShape(),
    },
    verification: Object.fromEntries(verificationKeys.map((key) => [key, 'placeholder'])),
  }
}

function numericSummaryShape(): Record<string, unknown> {
  return {
    samples: [],
    min: 0,
    median: 0,
    p95: 0,
    max: 0,
    mean: 0,
    standardDeviation: 0,
    relativeStandardDeviation: 0,
    standardError: 0,
    confidence95: { low: 0, high: 0 },
  }
}

function isLeadershipWorkload(workload: string): boolean {
  return (
    workload === 'lookup-reverse-search' ||
    workload === 'dynamic-array-filter' ||
    workload === 'dynamic-array-sort' ||
    workload === 'dynamic-array-unique'
  )
}

function normalizeArtifactShape(input: ArtifactShapeInput) {
  return {
    schemaVersion: input.schemaVersion,
    suite: input.suite,
    workloads: input.results.map((result) => ({
      workload: result.workload,
      category: result.category,
      comparable: result.comparable,
      hasComparison: result.comparison !== undefined,
      hasNote: result.note !== undefined,
      hyperformulaStatus: result.engines.hyperformula.status,
    })),
  }
}

function parseArtifactForShape(json: string): ArtifactShapeInput {
  const parsed = parseJsonRecord(json)
  if (!isArtifactShapeInput(parsed)) {
    throw new Error('Competitive benchmark artifact has an unexpected format')
  }
  return parsed
}

function parseArtifactForReport(json: string): ArtifactReportInput {
  const parsed = parseJsonRecord(json)
  if (!isArtifactReportInput(parsed)) {
    throw new Error('Competitive benchmark artifact report has an unexpected format')
  }
  return parsed
}

function assertArtifactReportDerivedFromResults(reportArtifact: ArtifactReportInput): void {
  const expectedReport = buildExpandedCompetitiveFamilyReport(reportArtifact.results)
  const mismatchedSections: string[] = []
  if (!reportsMatch(reportArtifact.families, expectedReport.families)) {
    mismatchedSections.push('families')
  }
  if (!reportsMatch(reportArtifact.scorecard, expectedReport.scorecard)) {
    mismatchedSections.push('scorecard')
  }
  if (mismatchedSections.length === 0) {
    return
  }

  throw new Error(
    `WorkPaper competitive benchmark artifact ${mismatchedSections.join(
      ' and ',
    )} do not match the raw benchmark results. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts`,
  )
}

function assertEngineSourcePath(artifactRecord: Record<string, unknown>, engineName: string, expectedSourcePath: string): void {
  const engines = recordField(artifactRecord, 'engines')
  const engine = recordField(engines, engineName)
  const actualSourcePath = engine.sourcePath
  if (actualSourcePath !== expectedSourcePath) {
    throw new Error(
      `WorkPaper competitive benchmark ${engineName} sourcePath is stale. Expected ${expectedSourcePath}, got ${String(
        actualSourcePath,
      )}. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts`,
    )
  }
}

function assertBenchmarkSettings(
  artifactRecord: Record<string, unknown>,
  expected: { readonly sampleCount: number; readonly warmupCount: number },
): void {
  const benchmark = recordField(artifactRecord, 'benchmark')
  const actualSampleCount = benchmark.sampleCount
  const actualWarmupCount = benchmark.warmupCount
  if (actualSampleCount === expected.sampleCount && actualWarmupCount === expected.warmupCount) {
    return
  }

  throw new Error(
    `WorkPaper competitive benchmark sample settings are stale. Expected sampleCount=${String(
      expected.sampleCount,
    )} warmupCount=${String(expected.warmupCount)}, got sampleCount=${String(actualSampleCount)} warmupCount=${String(
      actualWarmupCount,
    )}. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts`,
  )
}

function assertNoRawBenchmarkSampleArrays(artifactRecord: Record<string, unknown>): void {
  const sampleArrayCount = countRawBenchmarkSampleArrays(artifactRecord)
  if (sampleArrayCount === 0) {
    return
  }

  throw new Error(
    `WorkPaper competitive benchmark artifact stores ${String(
      sampleArrayCount,
    )} raw sample arrays. Run: bun scripts/gen-workpaper-vs-hyperformula-benchmark.ts`,
  )
}

function compactRawBenchmarkSampleArrays(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(compactRawBenchmarkSampleArrays)
  }
  if (!isRecord(value)) {
    return value
  }

  const compacted: Record<string, unknown> = {}
  for (const [key, child] of Object.entries(value)) {
    if (key === 'samples' && isRawNumberArray(child)) {
      continue
    }
    compacted[key] = compactRawBenchmarkSampleArrays(child)
  }
  return compacted
}

function countRawBenchmarkSampleArrays(value: unknown): number {
  if (Array.isArray(value)) {
    return value.reduce((count, child) => count + countRawBenchmarkSampleArrays(child), 0)
  }
  if (!isRecord(value)) {
    return 0
  }

  let count = 0
  for (const [key, child] of Object.entries(value)) {
    if (key === 'samples' && isRawNumberArray(child)) {
      count += 1
      continue
    }
    count += countRawBenchmarkSampleArrays(child)
  }
  return count
}

function isRawNumberArray(value: unknown): value is number[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'number')
}

function recordField(value: Record<string, unknown>, field: string): Record<string, unknown> {
  const child = value[field]
  if (!isRecord(child)) {
    throw new Error(`Expected ${field} to be an object`)
  }
  return child
}

function reportsMatch(actual: unknown, expected: unknown): boolean {
  if (typeof actual === 'number' && typeof expected === 'number') {
    return numbersMatch(actual, expected)
  }
  if (actual === expected) {
    return true
  }
  if (Array.isArray(actual) || Array.isArray(expected)) {
    if (!Array.isArray(actual) || !Array.isArray(expected) || actual.length !== expected.length) {
      return false
    }
    return actual.every((actualItem, index) => reportsMatch(actualItem, expected[index]))
  }
  if (isRecord(actual) || isRecord(expected)) {
    if (!isRecord(actual) || !isRecord(expected)) {
      return false
    }
    const actualKeys = Object.keys(actual)
    const expectedKeys = Object.keys(expected)
    if (actualKeys.length !== expectedKeys.length) {
      return false
    }
    return actualKeys.every((key) => Object.prototype.hasOwnProperty.call(expected, key) && reportsMatch(actual[key], expected[key]))
  }
  return false
}

function numbersMatch(actual: number, expected: number): boolean {
  if (Object.is(actual, expected)) {
    return true
  }
  const scale = Math.max(1, Math.abs(actual), Math.abs(expected))
  return Math.abs(actual - expected) <= Number.EPSILON * 64 * scale
}

function readPackageVersion(packagePath: string): string {
  const pkg = parseJsonRecord(readFileSync(packagePath, 'utf8'))
  if (typeof pkg.version !== 'string' || pkg.version.length === 0) {
    throw new Error(`Unable to read package version from ${packagePath}`)
  }
  return pkg.version
}

function readHyperFormulaMetadata(localRoot: string): ExpandedCompetitiveBenchmarkArtifact['engines']['hyperformula'] {
  const fallback = {
    packageName: 'hyperformula' as const,
    version: '3.2.0',
    sourcePath: localRoot,
    licenseKey: 'gpl-v3',
    metadataSource: 'fallback' as const,
    commit: 'unknown',
  }

  if (!existsSync(localRoot)) {
    return fallback
  }

  const packageJsonPath = join(localRoot, 'package.json')
  if (!existsSync(packageJsonPath)) {
    return fallback
  }

  const pkg = parseJsonRecord(readFileSync(packageJsonPath, 'utf8'))
  const version = typeof pkg.version === 'string' && pkg.version.length > 0 ? pkg.version : fallback.version
  const commit = readGitCommit(localRoot)
  return {
    ...fallback,
    version,
    commit,
    metadataSource: 'local-checkout',
  }
}

function readGitCommit(cwd: string): string {
  try {
    return execFileSync('git', ['rev-parse', 'HEAD'], {
      cwd,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim()
  } catch {
    return 'unknown'
  }
}

function parseJsonRecord(json: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(json)
  if (!isRecord(parsed)) {
    throw new Error('Expected JSON object')
  }
  return parsed
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

function isArtifactShapeInput(value: Record<string, unknown>): value is ArtifactShapeInput {
  return value.schemaVersion === 1 && value.suite === 'workpaper-vs-hyperformula' && Array.isArray(value.results)
}

function isArtifactReportInput(value: Record<string, unknown>): value is ArtifactReportInput {
  return isArtifactShapeInput(value) && Array.isArray(value.families) && isRecord(value.scorecard)
}
