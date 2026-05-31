import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { parsePublicWorkbookCorpusCase } from './public-workbook-corpus-json.ts'
import { compactRepoLocalPaths } from './public-workbook-corpus-output.ts'
import { formatByteSize, startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'
import { inspectWorkbookFootprintIsolatedWithMetrics } from './public-workbook-corpus-footprint.ts'
import {
  formulaOracleFormulaCountResourceLimitPreflight,
  importResourceLimitPreflight,
  roundTripResourceLimitPreflight,
  structuralSmokeResourceLimitPreflight,
  unsupportedPreflightResourceLimitCase,
  unsupportedPreflightResourceLimitCaseForLimits,
  unsupportedResourceLimitCase,
  unsupportedRssLimitCase,
  type ResourceLimitPreflight,
} from './public-workbook-corpus-resource-limits.ts'
import {
  startVerificationRuntimeMetrics,
  withPeakRssBytes,
  withVerificationRuntimeMetrics,
} from './public-workbook-corpus-verification-metrics.ts'
import { artifactBaseEvidence, failedCase } from './public-workbook-corpus-verify-cases.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookFeatureCounts } from './public-workbook-corpus-types.ts'
import type { WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const publicWorkbookCorpusFootprintWorkerScriptPath = fileURLToPath(
  new URL('./public-workbook-corpus-footprint-worker.ts', import.meta.url),
)
const publicWorkbookCorpusVerifyWorkerScriptPath = fileURLToPath(new URL('./public-workbook-corpus-verify-worker.ts', import.meta.url))
const memorySensitiveFootprintPreflightMaxRssBytes = 256 * 1024 * 1024
const noop = (): void => undefined

export const verificationWorkerPhasePrefix = 'bilig-public-workbook-verify-phase='
export const disableBunSmolVerificationWorkerEnvVar = 'BILIG_PUBLIC_WORKBOOK_VERIFY_DISABLE_BUN_SMOL'

type RuntimeVersions = Readonly<Record<string, string | undefined>>

interface VerifyCachedWorkbookArtifactIsolatedArgs {
  readonly artifact: PublicWorkbookArtifact
  readonly cacheDir: string
  readonly manifestPath: string
  readonly runStructuralSmoke: boolean
  readonly timeoutMs: number
  readonly maxRssBytes: number
  readonly maxCellCount: number
  readonly rssCheckIntervalMs?: number
}

export function shouldUseBunSmolForVerificationWorker(
  args: {
    readonly versions?: RuntimeVersions
    readonly env?: Readonly<Record<string, string | undefined>>
  } = {},
): boolean {
  const versions = args.versions ?? readProcessRuntimeVersions()
  if (!readRuntimeVersion(versions, 'bun')) {
    return false
  }
  const env = args.env ?? process.env
  const disabled = env[disableBunSmolVerificationWorkerEnvVar]
  return disabled !== '1' && disabled?.toLowerCase() !== 'true'
}

function readProcessRuntimeVersions(): RuntimeVersions {
  const versions: Record<string, string | undefined> = {}
  for (const [key, value] of Object.entries(process.versions)) {
    if (typeof value === 'string') {
      versions[key] = value
    }
  }
  return versions
}

function readRuntimeVersion(versions: RuntimeVersions, key: string): string | undefined {
  for (const [name, value] of Object.entries(versions)) {
    if (name === key && typeof value === 'string') {
      return value
    }
  }
  return undefined
}

export function buildVerificationWorkerProcessArgs(
  workerArgs: readonly string[],
  args: {
    readonly versions?: RuntimeVersions
    readonly env?: Readonly<Record<string, string | undefined>>
  } = {},
): string[] {
  return shouldUseBunSmolForVerificationWorker(args) ? ['--smol', ...workerArgs] : [...workerArgs]
}

export function verifyCachedWorkbookArtifactIsolated(args: VerifyCachedWorkbookArtifactIsolatedArgs): Promise<PublicWorkbookCorpusCase> {
  const baseEvidence = artifactBaseEvidence(args.artifact)
  const runtimeMetrics = startVerificationRuntimeMetrics()
  if (shouldUseResourceLimitedFootprintPreflight(args.maxRssBytes) && existsSync(join(args.cacheDir, args.artifact.cachePath))) {
    return tryVerifyResourceLimitedFootprintIsolated(args, baseEvidence, runtimeMetrics).then(
      (resourceLimitedFootprintCase) =>
        resourceLimitedFootprintCase ?? verifyCachedWorkbookArtifactInWorker(args, baseEvidence, runtimeMetrics),
    )
  }
  return verifyCachedWorkbookArtifactInWorker(args, baseEvidence, runtimeMetrics)
}

function verifyCachedWorkbookArtifactInWorker(
  args: VerifyCachedWorkbookArtifactIsolatedArgs,
  baseEvidence: readonly string[],
  runtimeMetrics: ReturnType<typeof startVerificationRuntimeMetrics>,
): Promise<PublicWorkbookCorpusCase> {
  return new Promise<PublicWorkbookCorpusCase>((resolvePromise) => {
    const workerArgs = [
      publicWorkbookCorpusVerifyWorkerScriptPath,
      'verify-artifact-worker',
      '--manifest',
      args.manifestPath,
      '--cache-dir',
      args.cacheDir,
      '--artifact-id',
      args.artifact.id,
      '--artifact-json-base64',
      Buffer.from(JSON.stringify(args.artifact), 'utf8').toString('base64'),
      '--verify-max-rss-mb',
      String(Math.ceil(args.maxRssBytes / 1024 / 1024)),
      '--verify-max-cells',
      String(args.maxCellCount),
      ...(args.runStructuralSmoke ? ['--structural-smoke'] : []),
    ]
    const childArgs = buildVerificationWorkerProcessArgs(workerArgs)
    const child = spawn(process.execPath, childArgs, {
      detached: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let stderrRemainder = ''
    let latestWorkerPhase = 'startup'
    let peakRssBytes = 0
    let timer: ReturnType<typeof setTimeout>
    let stopRssWatchdog = noop
    const finish = createOneShotResolver(resolvePromise, () => {
      clearTimeout(timer)
      stopRssWatchdog()
    })
    const terminateChild = (signal: 'SIGTERM' | 'SIGKILL'): void => {
      terminateChildProcess(child, signal, { processGroup: true })
    }
    stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes: args.maxRssBytes,
      intervalMs: args.rssCheckIntervalMs,
      onSample: (rssBytes) => {
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
      },
      onLimitExceeded: (rssBytes) => {
        terminateChild('SIGTERM')
        const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
        forceKillTimer.unref()
        finish(
          withVerificationRuntimeMetrics(
            unsupportedRssLimitCase(args.artifact, baseEvidence, rssBytes, args.maxRssBytes, [
              `rss-limit-phase=${latestWorkerPhase}`,
              `peak-rss=${formatByteSize(Math.max(peakRssBytes, rssBytes))}`,
              'The workbook was isolated in a subprocess so the corpus verification run could continue.',
            ]),
            runtimeMetrics,
            Math.max(peakRssBytes, rssBytes),
          ),
        )
      },
    })
    timer = setTimeout(() => {
      terminateChild('SIGTERM')
      const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
      forceKillTimer.unref()
      finish(
        withVerificationRuntimeMetrics(
          failedCase(args.artifact, 'error', baseEvidence, [
            `Verification timed out after ${String(args.timeoutMs)}ms`,
            'The workbook was isolated in a subprocess so the corpus verification run could continue.',
          ]),
          runtimeMetrics,
          peakRssBytes,
        ),
      )
    }, args.timeoutMs)
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
      const lines = `${stderrRemainder}${chunk}`.split(/\r?\n/u)
      stderrRemainder = lines.pop() ?? ''
      for (const line of lines) {
        if (line.startsWith(verificationWorkerPhasePrefix)) {
          latestWorkerPhase = line.slice(verificationWorkerPhasePrefix.length)
        }
      }
    })
    child.on('error', (error) => {
      finish(
        withVerificationRuntimeMetrics(
          failedCase(args.artifact, 'error', baseEvidence, [`Verification subprocess failed to start: ${error.message}`]),
          runtimeMetrics,
          peakRssBytes,
        ),
      )
    })
    child.on('close', (code, signal) => {
      if (code !== 0) {
        const failureDetails = compactVerificationWorkerOutput(stderr || stdout)
        finish(
          withVerificationRuntimeMetrics(
            failedCase(args.artifact, 'error', baseEvidence, [
              `Verification subprocess exited with ${signal ? `signal ${signal}` : `code ${String(code)}`}`,
              ...(failureDetails ? [failureDetails] : []),
            ]),
            runtimeMetrics,
            peakRssBytes,
          ),
        )
        return
      }
      try {
        const parsed: unknown = JSON.parse(stdout)
        finish(withPeakRssBytes(parsePublicWorkbookCorpusCase(parsed), peakRssBytes))
      } catch (error) {
        const details = compactVerificationWorkerOutput(stderr || stdout)
        finish(
          withVerificationRuntimeMetrics(
            failedCase(args.artifact, 'error', baseEvidence, [
              `Verification subprocess returned invalid JSON: ${error instanceof Error ? error.message : String(error)}`,
              ...(details ? [details] : []),
            ]),
            runtimeMetrics,
            peakRssBytes,
          ),
        )
      }
    })
  })
}

export function shouldUseResourceLimitedFootprintPreflight(maxRssBytes: number): boolean {
  return maxRssBytes <= memorySensitiveFootprintPreflightMaxRssBytes
}

export function buildResourceLimitedFootprintVerificationCase(args: {
  readonly artifact: PublicWorkbookArtifact
  readonly footprint: WorkbookFootprint
  readonly baseEvidence: readonly string[]
  readonly runStructuralSmoke: boolean
  readonly maxCellCount: number
}): PublicWorkbookCorpusCase | null {
  if (args.footprint.featureCounts.cellCount > args.maxCellCount) {
    return unsupportedResourceLimitCase(args.artifact, args.baseEvidence, args.footprint, args.maxCellCount)
  }
  const importResourceLimit = importResourceLimitPreflight(args.artifact, args.footprint)
  if (importResourceLimit) {
    return unsupportedPreflightResourceLimitCase(args.artifact, args.baseEvidence, args.footprint, importResourceLimit)
  }
  if (args.footprint.largeSimpleXlsxImport?.eligible !== true) {
    return null
  }
  const compactResourceLimits = largeSimpleFootprintCompactResourceLimits(
    args.artifact,
    args.footprint.featureCounts,
    args.runStructuralSmoke,
  )
  if (!compactResourceLimits) {
    return null
  }
  return unsupportedPreflightResourceLimitCaseForLimits(args.artifact, args.baseEvidence, args.footprint, compactResourceLimits)
}

async function tryVerifyResourceLimitedFootprintIsolated(
  args: {
    readonly artifact: PublicWorkbookArtifact
    readonly cacheDir: string
    readonly runStructuralSmoke: boolean
    readonly timeoutMs: number
    readonly maxRssBytes: number
    readonly maxCellCount: number
    readonly rssCheckIntervalMs?: number
  },
  baseEvidence: readonly string[],
  runtimeMetrics: ReturnType<typeof startVerificationRuntimeMetrics>,
): Promise<PublicWorkbookCorpusCase | null> {
  const cachePath = join(args.cacheDir, args.artifact.cachePath)
  if (!existsSync(cachePath)) {
    return null
  }
  const startedAt = performance.now()
  const footprintResult = await inspectWorkbookFootprintIsolatedWithMetrics({
    bytes: new Uint8Array(0),
    filePath: cachePath,
    fileName: args.artifact.fileName,
    scriptPath: publicWorkbookCorpusFootprintWorkerScriptPath,
    options: {
      timeoutMs: args.timeoutMs,
      maxRssBytes: args.maxRssBytes,
      rssCheckIntervalMs: args.rssCheckIntervalMs ?? 500,
    },
  })
  runtimeMetrics.phaseTimings.push({ phase: 'inspect-footprint', elapsedMs: roundElapsedMs(performance.now() - startedAt) })
  if (!footprintResult.footprint) {
    if (footprintResult.peakRssBytes > args.maxRssBytes) {
      return withVerificationRuntimeMetrics(
        unsupportedRssLimitCase(args.artifact, baseEvidence, footprintResult.peakRssBytes, args.maxRssBytes, [
          'rss-limit-phase=inspect-footprint',
          `peak-rss=${formatByteSize(footprintResult.peakRssBytes)}`,
          'The workbook was isolated in a subprocess so the corpus verification run could continue.',
        ]),
        runtimeMetrics,
        footprintResult.peakRssBytes,
      )
    }
    return null
  }
  const corpusCase = buildResourceLimitedFootprintVerificationCase({
    artifact: args.artifact,
    footprint: footprintResult.footprint,
    baseEvidence,
    runStructuralSmoke: args.runStructuralSmoke,
    maxCellCount: args.maxCellCount,
  })
  return corpusCase ? withVerificationRuntimeMetrics(corpusCase, runtimeMetrics, footprintResult.peakRssBytes) : null
}

function largeSimpleFootprintCompactResourceLimits(
  artifact: PublicWorkbookArtifact,
  featureCounts: PublicWorkbookFeatureCounts,
  runStructuralSmoke: boolean,
): readonly ResourceLimitPreflight[] | null {
  const formulaOracleResourceLimit = formulaOracleFormulaCountResourceLimitPreflight(featureCounts)
  const roundTripResourceLimit = roundTripResourceLimitPreflight(artifact, featureCounts)
  const structuralSmokeResourceLimit = runStructuralSmoke ? structuralSmokeResourceLimitPreflight(featureCounts) : null
  if (
    featureCounts.cellCount === 0 ||
    (featureCounts.formulaCellCount > 0 && formulaOracleResourceLimit === null) ||
    roundTripResourceLimit === null ||
    (runStructuralSmoke && structuralSmokeResourceLimit === null)
  ) {
    return null
  }
  return [
    ...(formulaOracleResourceLimit ? [formulaOracleResourceLimit] : []),
    roundTripResourceLimit,
    ...(structuralSmokeResourceLimit ? [structuralSmokeResourceLimit] : []),
  ]
}

export function compactVerificationWorkerOutput(value: string): string | null {
  const withoutPhaseMarkers = value
    .split(/\r?\n/u)
    .filter((line) => !line.startsWith(verificationWorkerPhasePrefix))
    .join('\n')
  const compacted = compactRepoLocalPaths(withoutPhaseMarkers, rootDir).replace(/\s+/gu, ' ').trim()
  return compacted.length > 0 ? compacted.slice(0, 1_000) : null
}

function roundElapsedMs(value: number): number {
  return Math.max(0, Math.round(value))
}

function createOneShotResolver<T>(resolveValue: (value: T) => void, cleanup: () => void): (value: T) => void {
  let settled = false
  return (value) => {
    if (settled) {
      return
    }
    settled = true
    cleanup()
    resolveValue(value)
  }
}
