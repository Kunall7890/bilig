import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { execFileSync } from 'node:child_process'
import {
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-hyperformula.ts'
import {
  buildIronCalcRustRunnerInput,
  buildWorkPaperVsIronCalcRustBenchmarkReport,
  IRONCALC_RUST_CRATE_NAME,
  IRONCALC_RUST_CRATE_VERSION,
  runWorkPaperVsIronCalcRustBenchmarkSuite,
  WORKPAPER_IRONCALC_RUST_UNSUPPORTED_WORKLOADS,
  WORKPAPER_IRONCALC_RUST_WORKLOADS,
  type IronCalcRustApiPath,
  type IronCalcRustRunnerOutput,
  type WorkPaperIronCalcRustWorkload,
  type WorkPaperIronCalcRustBenchmarkResult,
  type WorkPaperIronCalcRustScorecard,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-ironcalc-rust.ts'
import { arrayField, asObject, numberArrayField, numberField, objectField, readJsonObject, stringField } from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import { ironCalcRustSidecarCargoToml, ironCalcRustSidecarMainRs } from './ironcalc-rust-sidecar-template.ts'
import {
  deriveWorkPaperIronCalcRustScorecard,
  parseWorkPaperIronCalcRustArtifact,
  type ParsedWorkPaperIronCalcRustScorecard,
} from './workpaper-vs-ironcalc-rust-artifact.ts'

interface WorkPaperVsIronCalcRustBenchmarkArtifact {
  readonly schemaVersion: 1
  readonly suite: 'workpaper-vs-ironcalc-rust'
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
    readonly ironCalcRust: {
      readonly coverageTier: 'workbook-wide-limited'
      readonly packageName: 'ironcalc_base'
      readonly sourcePath: string
      readonly version: string
    }
  }
  readonly scorecard: WorkPaperIronCalcRustScorecard
  readonly results: readonly WorkPaperIronCalcRustBenchmarkResult[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-ironcalc-rust.json')
const cacheDir = join(rootDir, '.cache', 'ironcalc-rust-bench')
const sidecarSrcDir = join(cacheDir, 'src')
const sidecarInputPath = join(cacheDir, 'input.json')
const sidecarOutputPath = join(cacheDir, 'output.json')
const sidecarManifestPath = join(cacheDir, 'Cargo.toml')
const sidecarMainPath = join(sidecarSrcDir, 'main.rs')
const cargoTargetDir = join(rootDir, '.cache', 'ironcalc-rust-target')
const isCheckMode = process.argv.slice(2).includes('--check')
const sampleCount = Number.parseInt(process.env['BILIG_IRONCALC_RUST_SAMPLE_COUNT'] ?? '', 10) || DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT
const warmupCount = DEFAULT_COMPETITIVE_WARMUP_COUNT
const workpaperSourcePath = 'packages/headless'
const ironCalcRustSourcePath = '.cache/ironcalc-rust-bench'
const ironCalcRustWorkloadNames: readonly string[] = WORKPAPER_IRONCALC_RUST_WORKLOADS

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error('WorkPaper vs IronCalc Rust benchmark artifact is missing. Run: pnpm workpaper:bench:ironcalc-rust:generate')
  }

  const rawArtifact = readJsonObject(outputPath)
  assertEngineSourcePath(rawArtifact, 'workpaper', workpaperSourcePath)
  assertEngineSourcePath(rawArtifact, 'ironCalcRust', ironCalcRustSourcePath)
  assertEngineVersion(rawArtifact, 'workpaper', readPackageVersion(join(rootDir, 'packages', 'headless', 'package.json')))
  assertEngineVersion(rawArtifact, 'ironCalcRust', IRONCALC_RUST_CRATE_VERSION)
  const artifact = parseWorkPaperIronCalcRustArtifact(rawArtifact)
  const benchmark = objectField(rawArtifact, 'benchmark')
  const artifactSampleCount = numberField(benchmark, 'sampleCount')
  if (artifactSampleCount !== DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT) {
    throw new Error(
      `WorkPaper vs IronCalc Rust claim artifact must use ${DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT} samples, got ${String(
        benchmark['sampleCount'],
      )}. Run: pnpm workpaper:bench:ironcalc-rust:generate`,
    )
  }
  const actualWorkloads = artifact.results.map((result) => result.workload)
  if (JSON.stringify(actualWorkloads) !== JSON.stringify([...WORKPAPER_IRONCALC_RUST_WORKLOADS])) {
    throw new Error(
      'WorkPaper vs IronCalc Rust benchmark workload coverage is out of date. Run: pnpm workpaper:bench:ironcalc-rust:generate',
    )
  }

  const derivedScorecard = deriveWorkPaperIronCalcRustScorecard(
    artifact.results,
    artifact.scorecard.coverageNote,
    artifact.scorecard.unsupportedWorkloads,
  )
  if (!scorecardsMatch(artifact.scorecard, derivedScorecard)) {
    throw new Error(
      'WorkPaper vs IronCalc Rust scorecard does not match benchmark results. Run: pnpm workpaper:bench:ironcalc-rust:generate',
    )
  }
  if (JSON.stringify(artifact.scorecard.unsupportedWorkloads) !== JSON.stringify(WORKPAPER_IRONCALC_RUST_UNSUPPORTED_WORKLOADS)) {
    throw new Error('WorkPaper vs IronCalc Rust unsupported workload list is out of date. Run: pnpm workpaper:bench:ironcalc-rust:generate')
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

writeSidecarProject()
writeFileSync(sidecarInputPath, `${JSON.stringify(buildIronCalcRustRunnerInput({ sampleCount, warmupCount }), null, 2)}\n`)
execFileSync('cargo', ['run', '--release', '--quiet', '--manifest-path', sidecarManifestPath, '--', sidecarInputPath, sidecarOutputPath], {
  cwd: rootDir,
  env: {
    ...process.env,
    CARGO_TARGET_DIR: cargoTargetDir,
  },
  stdio: 'inherit',
})
const runnerOutput = parseIronCalcRustRunnerOutput(readJsonObject(sidecarOutputPath))
if (runnerOutput.engine.crate !== IRONCALC_RUST_CRATE_NAME || runnerOutput.engine.version !== IRONCALC_RUST_CRATE_VERSION) {
  throw new Error(
    `IronCalc Rust sidecar version mismatch: expected ${IRONCALC_RUST_CRATE_NAME}@${IRONCALC_RUST_CRATE_VERSION}, got ${runnerOutput.engine.crate}@${runnerOutput.engine.version}`,
  )
}
const report = buildWorkPaperVsIronCalcRustBenchmarkReport(
  runWorkPaperVsIronCalcRustBenchmarkSuite(runnerOutput, {
    sampleCount,
    warmupCount,
  }),
)
const artifact: WorkPaperVsIronCalcRustBenchmarkArtifact = {
  schemaVersion: 1,
  suite: 'workpaper-vs-ironcalc-rust',
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
    ironCalcRust: {
      coverageTier: 'workbook-wide-limited',
      packageName: IRONCALC_RUST_CRATE_NAME,
      sourcePath: ironCalcRustSourcePath,
      version: IRONCALC_RUST_CRATE_VERSION,
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

function writeSidecarProject(): void {
  mkdirSync(sidecarSrcDir, { recursive: true })
  writeFileSync(sidecarManifestPath, ironCalcRustSidecarCargoToml())
  writeFileSync(sidecarMainPath, ironCalcRustSidecarMainRs())
}

function assertEngineSourcePath(artifactRecord: Record<string, unknown>, engineName: string, expectedSourcePath: string): void {
  const engine = objectField(objectField(artifactRecord, 'engines'), engineName)
  const actualSourcePath = stringField(engine, 'sourcePath')
  if (actualSourcePath !== expectedSourcePath) {
    throw new Error(
      `WorkPaper vs IronCalc Rust ${engineName} sourcePath is stale. Expected ${expectedSourcePath}, got ${actualSourcePath}. Run: pnpm workpaper:bench:ironcalc-rust:generate`,
    )
  }
}

function assertEngineVersion(artifactRecord: Record<string, unknown>, engineName: string, expectedVersion: string): void {
  const engine = objectField(objectField(artifactRecord, 'engines'), engineName)
  const actualVersion = stringField(engine, 'version')
  if (actualVersion !== expectedVersion) {
    throw new Error(
      `WorkPaper vs IronCalc Rust ${engineName} version is stale. Expected ${expectedVersion}, got ${actualVersion}. Run: pnpm workpaper:bench:ironcalc-rust:generate`,
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

function parseIronCalcRustRunnerOutput(value: Record<string, unknown>): IronCalcRustRunnerOutput {
  const engine = objectField(value, 'engine')
  const crateName = stringField(engine, 'crate')
  const version = stringField(engine, 'version')
  if (crateName !== IRONCALC_RUST_CRATE_NAME || version !== IRONCALC_RUST_CRATE_VERSION) {
    throw new Error(
      `IronCalc Rust sidecar version mismatch: expected ${IRONCALC_RUST_CRATE_NAME}@${IRONCALC_RUST_CRATE_VERSION}, got ${crateName}@${version}`,
    )
  }
  const results = arrayField(value, 'results').map((entry, index) => {
    const result = asObject(entry, `results[${String(index)}]`)
    return {
      apiPath: parseIronCalcRustApiPath(stringField(result, 'apiPath')),
      elapsedMs: numberArrayField(result, 'elapsedMs'),
      verification: objectField(result, 'verification'),
      workload: parseIronCalcRustWorkload(stringField(result, 'workload')),
    }
  })
  return {
    engine: {
      crate: IRONCALC_RUST_CRATE_NAME,
      version,
    },
    results,
  }
}

function parseIronCalcRustApiPath(apiPath: string): IronCalcRustApiPath {
  if (apiPath === 'Model' || apiPath === 'UserModel') {
    return apiPath
  }
  throw new Error(`Unknown IronCalc Rust apiPath: ${apiPath}`)
}

function parseIronCalcRustWorkload(workload: string): WorkPaperIronCalcRustWorkload {
  if (isIronCalcRustWorkload(workload)) {
    return workload
  }
  throw new Error(`Unknown IronCalc Rust workload: ${workload}`)
}

function isIronCalcRustWorkload(workload: string): workload is WorkPaperIronCalcRustWorkload {
  return ironCalcRustWorkloadNames.includes(workload)
}

function scorecardsMatch(actual: ParsedWorkPaperIronCalcRustScorecard, expected: ParsedWorkPaperIronCalcRustScorecard): boolean {
  return (
    actual.comparableWorkloadCount === expected.comparableWorkloadCount &&
    actual.coverageNote === expected.coverageNote &&
    actual.coverageTier === expected.coverageTier &&
    nearlyEqual(actual.directionalMeanRatioGeomean, expected.directionalMeanRatioGeomean) &&
    nearlyEqual(actual.directionalP95RatioGeomean, expected.directionalP95RatioGeomean) &&
    actual.ironCalcRustMeanWinCount === expected.ironCalcRustMeanWinCount &&
    actual.ironCalcRustP95WinCount === expected.ironCalcRustP95WinCount &&
    actual.meanAndP95WinCount === expected.meanAndP95WinCount &&
    actual.meanWinCount === expected.meanWinCount &&
    actual.p95WinCount === expected.p95WinCount &&
    JSON.stringify(actual.unsupportedWorkloads) === JSON.stringify(expected.unsupportedWorkloads) &&
    JSON.stringify(actual.workloadFamilies) === JSON.stringify(expected.workloadFamilies) &&
    actual.worstMeanRatioWorkload === expected.worstMeanRatioWorkload &&
    actual.worstP95RatioWorkload === expected.worstP95RatioWorkload &&
    actual.worstWorkpaperToIronCalcRustMeanRatio === expected.worstWorkpaperToIronCalcRustMeanRatio &&
    actual.worstWorkpaperToIronCalcRustP95Ratio === expected.worstWorkpaperToIronCalcRustP95Ratio
  )
}

function nearlyEqual(actual: number, expected: number): boolean {
  const tolerance = Number.EPSILON * Math.max(1, Math.abs(actual), Math.abs(expected)) * 64
  return Math.abs(actual - expected) <= tolerance
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}
