import { closeSync, existsSync, mkdirSync, openSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { spawn } from 'node:child_process'
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
const lockPath = join(cacheDir, 'artifact.lock')
const isCheckMode = process.argv.slice(2).includes('--check')
const sampleCount = Number.parseInt(process.env['BILIG_IRONCALC_RUST_SAMPLE_COUNT'] ?? '', 10) || DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT
const warmupCount = DEFAULT_COMPETITIVE_WARMUP_COUNT
const workpaperSourcePath = 'packages/headless'
const ironCalcRustSourcePath = '.cache/ironcalc-rust-bench'
const ironCalcRustWorkloadNames: readonly string[] = WORKPAPER_IRONCALC_RUST_WORKLOADS
const rustSidecarChunkSize = Number.parseInt(process.env['BILIG_IRONCALC_RUST_CHUNK_SIZE'] ?? '', 10) || 1

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

const releaseLock = acquireArtifactLock()
process.once('exit', releaseLock)

writeSidecarProject()
writeFileSync(sidecarInputPath, `${JSON.stringify(buildIronCalcRustRunnerInput({ sampleCount, warmupCount }), null, 2)}\n`)
const runnerOutput = await runIronCalcRustSidecarChunks()
if (runnerOutput.engine.crate !== IRONCALC_RUST_CRATE_NAME || runnerOutput.engine.version !== IRONCALC_RUST_CRATE_VERSION) {
  throw new Error(
    `IronCalc Rust sidecar version mismatch: expected ${IRONCALC_RUST_CRATE_NAME}@${IRONCALC_RUST_CRATE_VERSION}, got ${runnerOutput.engine.crate}@${runnerOutput.engine.version}`,
  )
}
const report = buildWorkPaperVsIronCalcRustBenchmarkReport(
  runWorkPaperVsIronCalcRustBenchmarkSuite(runnerOutput, {
    onWorkloadStart: (workload, index, total) => {
      console.log(`WorkPaper IronCalc workload ${String(index + 1)}/${String(total)}: ${workload}`)
    },
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

function acquireArtifactLock(): () => void {
  mkdirSync(cacheDir, { recursive: true })
  const lockBody = `${String(process.pid)}\n${new Date().toISOString()}\n`
  try {
    const lockFd = openSync(lockPath, 'wx')
    writeFileSync(lockFd, lockBody)
    closeSync(lockFd)
    return () => {
      rmSync(lockPath, { force: true })
    }
  } catch (error) {
    if (!isFileExistsError(error)) {
      throw error
    }
  }

  const existingLock = readExistingArtifactLock()
  if (existingLock?.pid !== undefined && processIsAlive(existingLock.pid)) {
    throw new Error(
      `WorkPaper vs IronCalc Rust benchmark generation is already running under pid ${String(
        existingLock.pid,
      )}. Wait for it to finish before regenerating the timing artifact.`,
    )
  }
  rmSync(lockPath, { force: true })
  const lockFd = openSync(lockPath, 'wx')
  writeFileSync(lockFd, lockBody)
  closeSync(lockFd)
  return () => {
    rmSync(lockPath, { force: true })
  }
}

function readExistingArtifactLock(): { pid: number | undefined } | undefined {
  if (!existsSync(lockPath)) {
    return undefined
  }
  const [pidText] = readFileSync(lockPath, 'utf8').split('\n')
  const pid = Number.parseInt(pidText ?? '', 10)
  return { pid: Number.isFinite(pid) && pid > 0 ? pid : undefined }
}

function processIsAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}

function isFileExistsError(error: unknown): boolean {
  return typeof error === 'object' && error !== null && 'code' in error && (error as { readonly code?: unknown }).code === 'EEXIST'
}

function writeSidecarProject(): void {
  mkdirSync(sidecarSrcDir, { recursive: true })
  writeFileSync(sidecarManifestPath, ironCalcRustSidecarCargoToml())
  writeFileSync(sidecarMainPath, ironCalcRustSidecarMainRs())
}

async function runIronCalcRustSidecarChunks(): Promise<IronCalcRustRunnerOutput> {
  const chunkSize = Number.isFinite(rustSidecarChunkSize) && rustSidecarChunkSize > 0 ? Math.trunc(rustSidecarChunkSize) : 1
  const chunks = chunkArray(WORKPAPER_IRONCALC_RUST_WORKLOADS, chunkSize)
  const results: IronCalcRustRunnerOutput['results'] = []
  for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex += 1) {
    const workloads = chunks[chunkIndex]
    const inputPath = join(cacheDir, `input-${String(chunkIndex + 1)}.json`)
    const chunkOutputPath = join(cacheDir, `output-${String(chunkIndex + 1)}.json`)
    writeFileSync(inputPath, `${JSON.stringify(buildIronCalcRustRunnerInput({ sampleCount, warmupCount, workloads }), null, 2)}\n`)
    console.log(
      `IronCalc Rust chunk ${String(chunkIndex + 1)}/${String(chunks.length)} (${String(workloads.length)} workloads): ${workloads.join(
        ', ',
      )}`,
    )
    const reusableOutput = readReusableIronCalcRustChunkOutput(chunkOutputPath, workloads)
    if (reusableOutput !== undefined) {
      console.log(`IronCalc Rust chunk ${String(chunkIndex + 1)}/${String(chunks.length)} reused`)
      results.push(...reusableOutput.results)
      continue
    }
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Chunks run sequentially so timing samples stay isolated and artifact ordering remains deterministic.
    await runIronCalcRustSidecar(inputPath, chunkOutputPath, `chunk ${String(chunkIndex + 1)}/${String(chunks.length)}`)
    const output = parseIronCalcRustRunnerOutput(readJsonObject(chunkOutputPath))
    results.push(...output.results)
  }
  const byWorkload = new Map(results.map((result) => [result.workload, result]))
  const orderedResults = WORKPAPER_IRONCALC_RUST_WORKLOADS.map((workload) => {
    const result = byWorkload.get(workload)
    if (result === undefined) {
      throw new Error(`IronCalc Rust sidecar chunking did not return workload ${workload}`)
    }
    return result
  })
  const output: IronCalcRustRunnerOutput = {
    engine: {
      crate: IRONCALC_RUST_CRATE_NAME,
      version: IRONCALC_RUST_CRATE_VERSION,
    },
    results: orderedResults,
  }
  writeFileSync(sidecarOutputPath, `${JSON.stringify(output, null, 2)}\n`)
  return output
}

function readReusableIronCalcRustChunkOutput(
  chunkOutputPath: string,
  workloads: readonly WorkPaperIronCalcRustWorkload[],
): IronCalcRustRunnerOutput | undefined {
  if (!existsSync(chunkOutputPath)) {
    return undefined
  }
  const output = parseIronCalcRustRunnerOutput(readJsonObject(chunkOutputPath))
  const actualWorkloads = output.results.map((result) => result.workload)
  if (JSON.stringify(actualWorkloads) !== JSON.stringify(workloads)) {
    return undefined
  }
  if (output.results.some((result) => result.elapsedMs.length !== sampleCount)) {
    return undefined
  }
  return output
}

function runIronCalcRustSidecar(inputPath: string, chunkOutputPath: string, label: string): Promise<void> {
  return new Promise((resolveRun, rejectRun) => {
    const child = spawn(
      'cargo',
      ['run', '--release', '--quiet', '--manifest-path', sidecarManifestPath, '--', inputPath, chunkOutputPath],
      {
        cwd: rootDir,
        env: {
          ...process.env,
          CARGO_TARGET_DIR: cargoTargetDir,
        },
        stdio: ['ignore', 'inherit', 'inherit'],
      },
    )
    const heartbeat = setInterval(() => {
      console.log(`IronCalc Rust ${label} still running`)
    }, 15_000)
    child.once('error', (error) => {
      clearInterval(heartbeat)
      rejectRun(error)
    })
    child.once('exit', (code, signal) => {
      clearInterval(heartbeat)
      if (code === 0) {
        resolveRun()
        return
      }
      rejectRun(new Error(`IronCalc Rust ${label} failed with code ${String(code)} signal ${String(signal)}`))
    })
  })
}

function chunkArray<T>(values: readonly T[], chunkSize: number): T[][] {
  const chunks: T[][] = []
  for (let start = 0; start < values.length; start += chunkSize) {
    chunks.push(values.slice(start, start + chunkSize))
  }
  return chunks
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
