import { closeSync, existsSync, mkdirSync, openSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { execFileSync, spawn } from 'node:child_process'
import { createHash } from 'node:crypto'
import {
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT,
  type ComparativeMeasuredEngineResult,
  type ComparativeMemorySummary,
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
  type WorkPaperIronCalcRustFixture,
  type WorkPaperIronCalcRustScorecard,
} from '../packages/benchmarks/src/benchmark-workpaper-vs-ironcalc-rust.ts'
import { univerScenario } from '../packages/benchmarks/src/benchmark-workpaper-vs-univer-workload-resolver.ts'
import type { NumericSummary } from '../packages/benchmarks/src/stats.ts'
import {
  arrayField,
  asObject,
  literalField,
  numberArrayField,
  numberField,
  objectField,
  readJsonObject,
  stringField,
} from './json-scorecard-helpers.ts'
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
const canonicalOutputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-ironcalc-rust.json')
const cacheDir = join(rootDir, '.cache', 'ironcalc-rust-bench')
const sidecarSrcDir = join(cacheDir, 'src')
const sidecarInputPath = join(cacheDir, 'input.json')
const sidecarOutputPath = join(cacheDir, 'output.json')
const sidecarManifestPath = join(cacheDir, 'Cargo.toml')
const sidecarMainPath = join(sidecarSrcDir, 'main.rs')
const cargoTargetDir = join(rootDir, '.cache', 'ironcalc-rust-target')
const lockPath = join(cacheDir, 'artifact.lock')
const workPaperResultCacheDir = join(cacheDir, 'workpaper-results')
const isCheckMode = process.argv.slice(2).includes('--check')
if (isCheckMode && process.env['BILIG_IRONCALC_RUST_WORKLOADS']) {
  throw new Error('BILIG_IRONCALC_RUST_WORKLOADS is only supported in generate mode, not --check')
}
if (isCheckMode && process.env['BILIG_IRONCALC_RUST_OUTPUT_PATH']) {
  throw new Error('BILIG_IRONCALC_RUST_OUTPUT_PATH is only supported in generate mode, not --check')
}
const sampleCount = Number.parseInt(process.env['BILIG_IRONCALC_RUST_SAMPLE_COUNT'] ?? '', 10) || DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT
const warmupCount = DEFAULT_COMPETITIVE_WARMUP_COUNT
const workpaperSourcePath = 'packages/headless'
const ironCalcRustSourcePath = '.cache/ironcalc-rust-bench'
const allIronCalcRustWorkloadNames = new Set<string>(WORKPAPER_IRONCALC_RUST_WORKLOADS)
const selectedIronCalcRustWorkloads = parseSelectedIronCalcRustWorkloads(process.env['BILIG_IRONCALC_RUST_WORKLOADS'])
const ironCalcRustWorkloadNames: readonly string[] = selectedIronCalcRustWorkloads
const outputPath = process.env['BILIG_IRONCALC_RUST_OUTPUT_PATH']
  ? resolve(rootDir, process.env['BILIG_IRONCALC_RUST_OUTPUT_PATH'])
  : canonicalOutputPath
const rustSidecarChunkSize = Number.parseInt(process.env['BILIG_IRONCALC_RUST_CHUNK_SIZE'] ?? '', 10) || 1

if (isCheckMode) {
  if (!existsSync(canonicalOutputPath)) {
    throw new Error('WorkPaper vs IronCalc Rust benchmark artifact is missing. Run: pnpm workpaper:bench:ironcalc-rust:generate')
  }

  const rawArtifact = readJsonObject(canonicalOutputPath)
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
  if (artifact.scorecard.meanAndP95WinCount !== artifact.scorecard.comparableWorkloadCount) {
    const losingWorkloads = artifact.results
      .filter((result) => result.comparison.workpaperToIronCalcRustMeanRatio >= 1 || result.comparison.workpaperToIronCalcRustP95Ratio >= 1)
      .map((result) => result.workload)
    throw new Error(
      `WorkPaper vs IronCalc Rust proof is not achieved: ${String(artifact.scorecard.meanAndP95WinCount)}/${String(
        artifact.scorecard.comparableWorkloadCount,
      )} workloads win mean+p95. Losing workloads: ${losingWorkloads.join(', ')}`,
    )
  }
  assertNoRawBenchmarkSampleArrays(rawArtifact)

  console.log(
    JSON.stringify(
      {
        mode: 'check',
        outputPath: canonicalOutputPath,
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
writeFileSync(
  sidecarInputPath,
  `${JSON.stringify(buildIronCalcRustRunnerInput({ sampleCount, warmupCount, workloads: selectedIronCalcRustWorkloads }), null, 2)}\n`,
)
const runnerOutput = await runIronCalcRustSidecarChunks()
if (runnerOutput.engine.crate !== IRONCALC_RUST_CRATE_NAME || runnerOutput.engine.version !== IRONCALC_RUST_CRATE_VERSION) {
  throw new Error(
    `IronCalc Rust sidecar version mismatch: expected ${IRONCALC_RUST_CRATE_NAME}@${IRONCALC_RUST_CRATE_VERSION}, got ${runnerOutput.engine.crate}@${runnerOutput.engine.version}`,
  )
}
const report = buildWorkPaperVsIronCalcRustBenchmarkReport(
  runWorkPaperVsIronCalcRustBenchmarkSuiteWithCache(runnerOutput, {
    onWorkloadStart: (workload, index, total) => {
      console.log(`WorkPaper IronCalc workload ${String(index + 1)}/${String(total)}: ${workload}`)
    },
    sampleCount,
    workloads: selectedIronCalcRustWorkloads,
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
writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(compactRawBenchmarkSampleArrays(artifact), null, 2)}\n`))
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
  const chunks = chunkArray(selectedIronCalcRustWorkloads, chunkSize)
  const results: IronCalcRustRunnerOutput['results'] = []
  for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex += 1) {
    const workloads = chunks[chunkIndex]
    const inputPath = join(cacheDir, `input-${String(chunkIndex + 1)}.json`)
    const chunkOutputPath = join(cacheDir, `output-${String(chunkIndex + 1)}.json`)
    const chunkMetadataPath = join(cacheDir, `output-${String(chunkIndex + 1)}.meta.json`)
    const inputText = `${JSON.stringify(buildIronCalcRustRunnerInput({ sampleCount, warmupCount, workloads }), null, 2)}\n`
    writeFileSync(inputPath, inputText)
    const expectedChunkMetadata = buildIronCalcRustChunkCacheMetadata(workloads, inputText)
    console.log(
      `IronCalc Rust chunk ${String(chunkIndex + 1)}/${String(chunks.length)} (${String(workloads.length)} workloads): ${workloads.join(
        ', ',
      )}`,
    )
    const reusableOutput = readReusableIronCalcRustChunkOutput(chunkOutputPath, chunkMetadataPath, expectedChunkMetadata)
    if (reusableOutput !== undefined) {
      console.log(`IronCalc Rust chunk ${String(chunkIndex + 1)}/${String(chunks.length)} reused`)
      results.push(...reusableOutput.results)
      continue
    }
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Chunks run sequentially so timing samples stay isolated and artifact ordering remains deterministic.
    await runIronCalcRustSidecar(inputPath, chunkOutputPath, `chunk ${String(chunkIndex + 1)}/${String(chunks.length)}`)
    const output = parseIronCalcRustRunnerOutput(readJsonObject(chunkOutputPath))
    writeIronCalcRustChunkCacheMetadata(chunkMetadataPath, buildIronCalcRustChunkCacheMetadata(workloads, inputText))
    results.push(...output.results)
  }
  const byWorkload = new Map(results.map((result) => [result.workload, result]))
  const orderedResults = selectedIronCalcRustWorkloads.map((workload) => {
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

function runWorkPaperVsIronCalcRustBenchmarkSuiteWithCache(
  ironCalcRunnerOutput: IronCalcRustRunnerOutput,
  options: {
    readonly onWorkloadStart?: (workload: WorkPaperIronCalcRustWorkload, index: number, total: number) => void
    readonly sampleCount: number
    readonly workloads: readonly WorkPaperIronCalcRustWorkload[]
    readonly warmupCount: number
  },
): WorkPaperIronCalcRustBenchmarkResult[] {
  const sourceHash = workPaperBenchmarkSourceHash()
  const ironCalcResultsByWorkload = new Map(ironCalcRunnerOutput.results.map((result) => [result.workload, result]))
  const results: WorkPaperIronCalcRustBenchmarkResult[] = []
  mkdirSync(workPaperResultCacheDir, { recursive: true })
  for (let index = 0; index < options.workloads.length; index += 1) {
    const workload = options.workloads[index]
    const ironCalcResult = ironCalcResultsByWorkload.get(workload)
    if (ironCalcResult === undefined) {
      throw new Error(`IronCalc Rust runner did not return workload ${workload}`)
    }
    options.onWorkloadStart?.(workload, index, options.workloads.length)
    const expectedMetadata = buildWorkPaperResultCacheMetadata(workload, ironCalcResult, sourceHash)
    const resultPath = join(workPaperResultCacheDir, `${workload}.json`)
    const metadataPath = join(workPaperResultCacheDir, `${workload}.meta.json`)
    const reusableResult = readReusableWorkPaperResult(resultPath, metadataPath, expectedMetadata)
    if (reusableResult !== undefined) {
      console.log(`WorkPaper IronCalc workload ${String(index + 1)}/${String(options.workloads.length)} reused: ${workload}`)
      results.push(reusableResult)
      continue
    }
    const [result] = runWorkPaperVsIronCalcRustBenchmarkSuite(
      {
        engine: ironCalcRunnerOutput.engine,
        results: [ironCalcResult],
      },
      {
        sampleCount: options.sampleCount,
        workloads: [workload],
        warmupCount: options.warmupCount,
      },
    )
    if (result === undefined) {
      throw new Error(`WorkPaper vs IronCalc Rust benchmark did not produce workload ${workload}`)
    }
    writeFileSync(resultPath, `${JSON.stringify(result, null, 2)}\n`)
    writeWorkPaperResultCacheMetadata(metadataPath, expectedMetadata)
    results.push(result)
  }
  return results
}

interface WorkPaperResultCacheMetadata {
  readonly schemaVersion: 1
  readonly cacheKey: string
  readonly sampleCount: number
  readonly warmupCount: number
  readonly workload: WorkPaperIronCalcRustWorkload
  readonly host: {
    readonly arch: string
    readonly nodeVersion: string
    readonly platform: string
  }
  readonly hashes: {
    readonly ironCalcRunnerResult: string
    readonly source: string
  }
}

function buildWorkPaperResultCacheMetadata(
  workload: WorkPaperIronCalcRustWorkload,
  ironCalcResult: IronCalcRustRunnerOutput['results'][number],
  sourceHash: string,
): WorkPaperResultCacheMetadata {
  const hashes = {
    ironCalcRunnerResult: sha256(JSON.stringify(ironCalcResult)),
    source: sourceHash,
  }
  const keyPayload = {
    schemaVersion: 1,
    sampleCount,
    warmupCount,
    workload,
    host: {
      arch: process.arch,
      nodeVersion: process.version,
      platform: process.platform,
    },
    hashes,
  }
  return {
    ...keyPayload,
    cacheKey: sha256(JSON.stringify(keyPayload)),
  }
}

function readReusableWorkPaperResult(
  resultPath: string,
  metadataPath: string,
  expected: WorkPaperResultCacheMetadata,
): WorkPaperIronCalcRustBenchmarkResult | undefined {
  if (!existsSync(resultPath) || !workPaperResultCacheMetadataMatches(metadataPath, expected)) {
    return undefined
  }
  try {
    const result = parseCachedWorkPaperResult(readJsonObject(resultPath))
    if (
      result.workload !== expected.workload ||
      result.engines.workpaper.status !== 'supported' ||
      result.engines.ironCalcRust.status !== 'supported' ||
      result.engines.workpaper.elapsedMs.samples.length !== expected.sampleCount ||
      result.engines.ironCalcRust.elapsedMs.samples.length !== expected.sampleCount
    ) {
      return undefined
    }
    return result
  } catch {
    return undefined
  }
}

function parseCachedWorkPaperResult(value: Record<string, unknown>): WorkPaperIronCalcRustBenchmarkResult {
  const workload = parseIronCalcRustWorkload(stringField(value, 'workload'))
  return {
    workload,
    category: literalField(value, 'category', 'workbook-wide-limited'),
    comparable: literalField(value, 'comparable', true),
    fixture: parseCachedWorkPaperFixture(objectField(value, 'fixture'), workload),
    comparison: parseCachedWorkPaperComparison(objectField(value, 'comparison')),
    engines: parseCachedWorkPaperEngines(objectField(value, 'engines')),
  }
}

function parseCachedWorkPaperFixture(
  value: Record<string, unknown>,
  workload: WorkPaperIronCalcRustWorkload,
): WorkPaperIronCalcRustFixture {
  const expectedFamily = univerScenario(workload).fixture.family
  const actualFamily = stringField(value, 'family')
  if (actualFamily !== expectedFamily) {
    throw new Error(`Cached WorkPaper result family mismatch for ${workload}: ${actualFamily}`)
  }
  const result = objectField(value, 'result')
  const editValue = value['edit']
  const edit = editValue === undefined ? undefined : parseCachedWorkPaperFixtureEdit(asObject(editValue, 'fixture.edit'))
  return {
    ...(edit !== undefined ? { edit } : {}),
    family: expectedFamily,
    formula: stringField(value, 'formula'),
    result: {
      address: stringField(result, 'address'),
      col: numberField(result, 'col'),
      row: numberField(result, 'row'),
      sheetName: stringField(result, 'sheetName'),
    },
    rowCount: numberField(value, 'rowCount'),
  }
}

function parseCachedWorkPaperFixtureEdit(value: Record<string, unknown>): NonNullable<WorkPaperIronCalcRustFixture['edit']> {
  return {
    address: stringField(value, 'address'),
    col: numberField(value, 'col'),
    row: numberField(value, 'row'),
    sheetName: stringField(value, 'sheetName'),
    value: parseCachedEditableValue(value['value']),
  }
}

function parseCachedEditableValue(value: unknown): boolean | number | string {
  if (typeof value === 'boolean' || typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }
  throw new Error('Expected cached fixture edit value to be a finite primitive')
}

function parseCachedWorkPaperComparison(value: Record<string, unknown>): WorkPaperIronCalcRustBenchmarkResult['comparison'] {
  const fasterEngine = stringField(value, 'fasterEngine')
  if (fasterEngine !== 'workpaper' && fasterEngine !== 'ironcalc-rust') {
    throw new Error(`Unexpected cached fasterEngine: ${fasterEngine}`)
  }
  return {
    confidenceIntervalOverlaps: literalBooleanField(value, 'confidenceIntervalOverlaps'),
    fasterEngine,
    maxRelativeNoise: numberField(value, 'maxRelativeNoise'),
    meanSpeedup: numberField(value, 'meanSpeedup'),
    verificationEquivalent: literalField(value, 'verificationEquivalent', true),
    workpaperToIronCalcRustMeanRatio: numberField(value, 'workpaperToIronCalcRustMeanRatio'),
    workpaperToIronCalcRustMedianRatio: numberField(value, 'workpaperToIronCalcRustMedianRatio'),
    workpaperToIronCalcRustP95Ratio: numberField(value, 'workpaperToIronCalcRustP95Ratio'),
  }
}

function parseCachedWorkPaperEngines(value: Record<string, unknown>): WorkPaperIronCalcRustBenchmarkResult['engines'] {
  const ironCalcRust = objectField(value, 'ironCalcRust')
  return {
    ironCalcRust: {
      status: literalField(ironCalcRust, 'status', 'supported'),
      apiPath: parseIronCalcRustApiPath(stringField(ironCalcRust, 'apiPath')),
      elapsedMs: parseNumericSummary(objectField(ironCalcRust, 'elapsedMs')),
      verification: objectField(ironCalcRust, 'verification'),
    },
    workpaper: parseCachedMeasuredEngineResult(objectField(value, 'workpaper')),
  }
}

function parseCachedMeasuredEngineResult(value: Record<string, unknown>): ComparativeMeasuredEngineResult {
  const engineCountersValue = value['engineCounters']
  const engineCounters =
    engineCountersValue === undefined ? undefined : parseCachedEngineCounters(asObject(engineCountersValue, 'engineCounters'))
  return {
    status: literalField(value, 'status', 'supported'),
    elapsedMs: parseNumericSummary(objectField(value, 'elapsedMs')),
    memoryDeltaBytes: parseMemorySummary(objectField(value, 'memoryDeltaBytes')),
    ...(engineCounters !== undefined ? { engineCounters } : {}),
    verification: objectField(value, 'verification'),
  }
}

function parseCachedEngineCounters(value: Record<string, unknown>): Record<string, NumericSummary> {
  return Object.fromEntries(
    Object.entries(value).map(([key, entry]) => [key, parseNumericSummary(asObject(entry, `engineCounters.${key}`))]),
  )
}

function parseMemorySummary(value: Record<string, unknown>): ComparativeMemorySummary {
  return {
    rssBytes: parseNumericSummary(objectField(value, 'rssBytes')),
    heapUsedBytes: parseNumericSummary(objectField(value, 'heapUsedBytes')),
    heapTotalBytes: parseNumericSummary(objectField(value, 'heapTotalBytes')),
    externalBytes: parseNumericSummary(objectField(value, 'externalBytes')),
    arrayBuffersBytes: parseNumericSummary(objectField(value, 'arrayBuffersBytes')),
  }
}

function parseNumericSummary(value: Record<string, unknown>): NumericSummary {
  return {
    samples: numberArrayField(value, 'samples'),
    min: numberField(value, 'min'),
    median: numberField(value, 'median'),
    p95: numberField(value, 'p95'),
    max: numberField(value, 'max'),
    mean: numberField(value, 'mean'),
    standardDeviation: numberField(value, 'standardDeviation'),
    relativeStandardDeviation: numberField(value, 'relativeStandardDeviation'),
    standardError: numberField(value, 'standardError'),
    confidence95: {
      low: numberField(objectField(value, 'confidence95'), 'low'),
      high: numberField(objectField(value, 'confidence95'), 'high'),
    },
  }
}

function literalBooleanField(value: Record<string, unknown>, field: string): boolean {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${field} to be a boolean`)
  }
  return fieldValue
}

function workPaperResultCacheMetadataMatches(metadataPath: string, expected: WorkPaperResultCacheMetadata): boolean {
  if (!existsSync(metadataPath)) {
    return false
  }
  const actual = readJsonObject(metadataPath)
  return stringField(actual, 'cacheKey') === expected.cacheKey
}

function writeWorkPaperResultCacheMetadata(metadataPath: string, metadata: WorkPaperResultCacheMetadata): void {
  writeFileSync(metadataPath, `${JSON.stringify(metadata, null, 2)}\n`)
}

function workPaperBenchmarkSourceHash(): string {
  const trackedFiles = execFileSync(
    'git',
    ['ls-files', 'packages/core/src', 'packages/headless/src', 'packages/benchmarks/src', 'scripts'],
    { cwd: rootDir, encoding: 'utf8' },
  )
    .split('\n')
    .filter((path) => path.length > 0)
    .toSorted()
  const hash = createHash('sha256')
  for (const filePath of trackedFiles) {
    hash.update(filePath)
    hash.update('\0')
    hash.update(readFileSync(join(rootDir, filePath)))
    hash.update('\0')
  }
  const diff = execFileSync(
    'git',
    ['diff', '--no-ext-diff', '--binary', '--', 'packages/core/src', 'packages/headless/src', 'packages/benchmarks/src', 'scripts'],
    { cwd: rootDir },
  )
  hash.update(diff)
  return hash.digest('hex')
}

function readReusableIronCalcRustChunkOutput(
  chunkOutputPath: string,
  metadataPath: string,
  expectedMetadata: IronCalcRustChunkCacheMetadata,
): IronCalcRustRunnerOutput | undefined {
  if (!existsSync(chunkOutputPath) || !ironCalcRustChunkCacheMetadataMatches(metadataPath, expectedMetadata)) {
    return undefined
  }
  const output = parseIronCalcRustRunnerOutput(readJsonObject(chunkOutputPath))
  const actualWorkloads = output.results.map((result) => result.workload)
  if (JSON.stringify(actualWorkloads) !== JSON.stringify(expectedMetadata.workloads)) {
    return undefined
  }
  if (output.results.some((result) => result.elapsedMs.length !== sampleCount)) {
    return undefined
  }
  return output
}

interface IronCalcRustChunkCacheMetadata {
  readonly schemaVersion: 1
  readonly cacheKey: string
  readonly crate: string
  readonly version: string
  readonly sampleCount: number
  readonly warmupCount: number
  readonly workloads: readonly WorkPaperIronCalcRustWorkload[]
  readonly host: {
    readonly arch: string
    readonly nodeVersion: string
    readonly platform: string
  }
  readonly hashes: {
    readonly cargoLock: string | null
    readonly cargoToml: string
    readonly inputJson: string
    readonly sidecarMain: string
  }
}

function buildIronCalcRustChunkCacheMetadata(
  workloads: readonly WorkPaperIronCalcRustWorkload[],
  inputText: string,
): IronCalcRustChunkCacheMetadata {
  const cargoToml = ironCalcRustSidecarCargoToml()
  const sidecarMain = ironCalcRustSidecarMainRs()
  const cargoLockPath = join(cacheDir, 'Cargo.lock')
  const hashes = {
    cargoLock: existsSync(cargoLockPath) ? sha256(readFileSync(cargoLockPath, 'utf8')) : null,
    cargoToml: sha256(cargoToml),
    inputJson: sha256(inputText),
    sidecarMain: sha256(sidecarMain),
  }
  const keyPayload = {
    schemaVersion: 1,
    crate: IRONCALC_RUST_CRATE_NAME,
    version: IRONCALC_RUST_CRATE_VERSION,
    sampleCount,
    warmupCount,
    workloads,
    host: {
      arch: process.arch,
      nodeVersion: process.version,
      platform: process.platform,
    },
    hashes,
  }
  return {
    ...keyPayload,
    cacheKey: sha256(JSON.stringify(keyPayload)),
  }
}

function ironCalcRustChunkCacheMetadataMatches(metadataPath: string, expected: IronCalcRustChunkCacheMetadata): boolean {
  if (!existsSync(metadataPath)) {
    return false
  }
  const actual = readJsonObject(metadataPath)
  return stringField(actual, 'cacheKey') === expected.cacheKey
}

function writeIronCalcRustChunkCacheMetadata(metadataPath: string, metadata: IronCalcRustChunkCacheMetadata): void {
  writeFileSync(metadataPath, `${JSON.stringify(metadata, null, 2)}\n`)
}

function sha256(content: string): string {
  return createHash('sha256').update(content).digest('hex')
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

function parseSelectedIronCalcRustWorkloads(raw: string | undefined): readonly WorkPaperIronCalcRustWorkload[] {
  if (raw === undefined || raw.trim().length === 0) {
    return WORKPAPER_IRONCALC_RUST_WORKLOADS
  }
  const selectedNames = raw
    .split(/[\s,]+/u)
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
  if (selectedNames.length === 0) {
    throw new Error('BILIG_IRONCALC_RUST_WORKLOADS did not contain any workload names')
  }
  const selected: WorkPaperIronCalcRustWorkload[] = []
  for (const workload of selectedNames) {
    if (!isKnownIronCalcRustWorkload(workload)) {
      throw new Error(`Unknown IronCalc Rust workload in BILIG_IRONCALC_RUST_WORKLOADS: ${workload}`)
    }
    selected.push(workload)
  }
  return selected
}

function isKnownIronCalcRustWorkload(workload: string): workload is WorkPaperIronCalcRustWorkload {
  return allIronCalcRustWorkloadNames.has(workload)
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

function assertNoRawBenchmarkSampleArrays(artifactRecord: Record<string, unknown>): void {
  const sampleArrayCount = countRawBenchmarkSampleArrays(artifactRecord)
  if (sampleArrayCount === 0) {
    return
  }

  throw new Error(
    `WorkPaper vs IronCalc Rust benchmark artifact stores ${String(
      sampleArrayCount,
    )} raw sample arrays. Run: pnpm workpaper:bench:ironcalc-rust:generate`,
  )
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
