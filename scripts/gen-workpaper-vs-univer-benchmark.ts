import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula.ts'
import {
  WORKPAPER_UNIVER_WORKLOADS,
  buildWorkPaperVsUniverBenchmarkReport,
  runWorkPaperVsUniverBenchmarkSuite,
  type WorkPaperUniverBenchmarkResult,
  type WorkPaperUniverScorecard,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-univer.ts'
import { objectField, readJsonObject, stringField } from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import {
  deriveWorkPaperUniverScorecard,
  parseWorkPaperUniverArtifact,
  type ParsedWorkPaperUniverScorecard,
} from './workpaper-vs-univer-artifact.ts'

interface WorkPaperVsUniverBenchmarkArtifact {
  readonly schemaVersion: 1
  readonly suite: 'workpaper-vs-univer'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly nodeVersion: string
    readonly platform: string
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly warmupCount: number
  }
  readonly engines: {
    readonly workpaper: {
      readonly packageName: '@bilig/headless'
      readonly sourcePath: string
      readonly version: string
    }
    readonly univer: {
      readonly coverageTier: 'workbook-wide'
      readonly packageName: '@univerjs/preset-sheets-node-core'
      readonly sourcePath: string
      readonly version: string
    }
  }
  readonly scorecard: WorkPaperUniverScorecard
  readonly results: readonly WorkPaperUniverBenchmarkResult[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-univer.json')
const isCheckMode = process.argv.slice(2).includes('--check')
const sampleCount = DEFAULT_COMPETITIVE_SAMPLE_COUNT
const warmupCount = DEFAULT_COMPETITIVE_WARMUP_COUNT
const workpaperSourcePath = 'packages/headless'
const univerSourcePath = 'packages/benchmarks/node_modules/@univerjs/preset-sheets-node-core'

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error('WorkPaper vs Univer benchmark artifact is missing. Run: pnpm workpaper:bench:univer:generate')
  }

  const rawArtifact = readJsonObject(outputPath)
  assertEngineSourcePath(rawArtifact, 'workpaper', workpaperSourcePath)
  assertEngineSourcePath(rawArtifact, 'univer', univerSourcePath)
  const artifact = parseWorkPaperUniverArtifact(rawArtifact)
  const actualWorkloads = artifact.results.map((result) => result.workload)
  if (JSON.stringify(actualWorkloads) !== JSON.stringify([...WORKPAPER_UNIVER_WORKLOADS])) {
    throw new Error('WorkPaper vs Univer benchmark workload coverage is out of date. Run: pnpm workpaper:bench:univer:generate')
  }

  const derivedScorecard = deriveWorkPaperUniverScorecard(artifact.results, artifact.scorecard.coverageNote)
  if (!scorecardsMatch(artifact.scorecard, derivedScorecard)) {
    throw new Error('WorkPaper vs Univer scorecard does not match benchmark results. Run: pnpm workpaper:bench:univer:generate')
  }

  console.log(
    JSON.stringify(
      {
        mode: 'check',
        outputPath,
        workloads: actualWorkloads,
      },
      null,
      2,
    ),
  )
  process.exit(0)
}

const report = buildWorkPaperVsUniverBenchmarkReport(
  await runWorkPaperVsUniverBenchmarkSuite({
    sampleCount,
    warmupCount,
  }),
)
const artifact: WorkPaperVsUniverBenchmarkArtifact = {
  schemaVersion: 1,
  suite: 'workpaper-vs-univer',
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
      version: readPackageVersion(join(rootDir, 'packages', 'headless', 'package.json')),
    },
    univer: {
      coverageTier: 'workbook-wide',
      packageName: '@univerjs/preset-sheets-node-core',
      sourcePath: univerSourcePath,
      version: readPackageVersion(
        join(rootDir, 'packages', 'benchmarks', 'node_modules', '@univerjs', 'preset-sheets-node-core', 'package.json'),
      ),
    },
  },
  scorecard: report.scorecard,
  results: report.results,
}

mkdirSync(dirname(outputPath), { recursive: true })
writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(artifact, null, 2)}\n`))
console.log(
  JSON.stringify(
    {
      mode: 'write',
      outputPath,
      workloads: artifact.results.map((result) => result.workload),
      meanAndP95WinCount: artifact.scorecard.meanAndP95WinCount,
      comparableWorkloadCount: artifact.scorecard.comparableWorkloadCount,
    },
    null,
    2,
  ),
)

function assertEngineSourcePath(artifactRecord: Record<string, unknown>, engineName: string, expectedSourcePath: string): void {
  const engine = objectField(objectField(artifactRecord, 'engines'), engineName)
  const actualSourcePath = stringField(engine, 'sourcePath')
  if (actualSourcePath !== expectedSourcePath) {
    throw new Error(
      `WorkPaper vs Univer ${engineName} sourcePath is stale. Expected ${expectedSourcePath}, got ${actualSourcePath}. Run: pnpm workpaper:bench:univer:generate`,
    )
  }
}

function readPackageVersion(packagePath: string): string {
  const parsed: unknown = JSON.parse(readFileSync(packagePath, 'utf8'))
  if (!isRecord(parsed) || typeof parsed.version !== 'string' || parsed.version.length === 0) {
    throw new Error(`Unable to read package version from ${packagePath}`)
  }
  return parsed.version
}

function scorecardsMatch(actual: ParsedWorkPaperUniverScorecard, expected: ParsedWorkPaperUniverScorecard): boolean {
  return (
    actual.comparableWorkloadCount === expected.comparableWorkloadCount &&
    actual.coverageNote === expected.coverageNote &&
    actual.coverageTier === expected.coverageTier &&
    nearlyEqual(actual.directionalMeanRatioGeomean, expected.directionalMeanRatioGeomean) &&
    nearlyEqual(actual.directionalP95RatioGeomean, expected.directionalP95RatioGeomean) &&
    actual.meanAndP95WinCount === expected.meanAndP95WinCount &&
    actual.meanWinCount === expected.meanWinCount &&
    actual.p95WinCount === expected.p95WinCount &&
    actual.univerMeanWinCount === expected.univerMeanWinCount &&
    actual.univerP95WinCount === expected.univerP95WinCount &&
    JSON.stringify(actual.workloadFamilies) === JSON.stringify(expected.workloadFamilies) &&
    actual.worstMeanRatioWorkload === expected.worstMeanRatioWorkload &&
    actual.worstP95RatioWorkload === expected.worstP95RatioWorkload &&
    actual.worstWorkpaperToUniverMeanRatio === expected.worstWorkpaperToUniverMeanRatio &&
    actual.worstWorkpaperToUniverP95Ratio === expected.worstWorkpaperToUniverP95Ratio
  )
}

function nearlyEqual(actual: number, expected: number): boolean {
  const tolerance = Number.EPSILON * Math.max(1, Math.abs(actual), Math.abs(expected)) * 64
  return Math.abs(actual - expected) <= tolerance
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}
