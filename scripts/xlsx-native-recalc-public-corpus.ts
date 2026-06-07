#!/usr/bin/env bun

import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

import { ErrorCode, ValueTag, type ErrorCode as ProtocolErrorCode } from '../packages/protocol/src/enums.js'
import type { CellValue } from '../packages/protocol/src/types.js'
import {
  readXlsxFormulaCacheCellsFromFile,
  type XlsxFormulaCacheCell,
  type XlsxFormulaCacheLiteral,
} from '../packages/xlsx/src/formula-cache-reader.js'
import { cellValuesMatchOracle } from './public-workbook-corpus-workbook.ts'
import { parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { readFlagArg, readMegabytesArg, readNumberArg, readStringArg } from './public-workbook-corpus-cli.ts'
import type { FormulaOracle, PublicWorkbookArtifact, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'
import { formatByteSize, startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const mib = 1024 * 1024
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultCliPath = join(rootDir, 'packages', 'xlsx-formula-recalc', 'dist', 'cli.js')
const defaultOutputDir = join(rootDir, '.cache', 'xlsx-native-recalc-public-corpus', 'outputs')
const defaultMaxRssBytes = 350 * mib
const defaultLimit = 50
const defaultMaxFormulaReadsPerWorkbook = 1_000
const defaultTimeoutMs = 180_000
const rssCheckIntervalMs = 10
const noop = (): void => undefined

export interface NativeRecalcPublicCorpusTarget {
  readonly id: string
  readonly label: string
  readonly sourceUrl: string
  readonly sha256: string
  readonly inputPath: string
  readonly outputPath: string
  readonly maxRssBytes: number
  readonly formulaCellCount: number
  readonly selectedFormulaCellCount: number
  readonly reads: readonly string[]
  readonly expectedReads: Readonly<Record<string, CellValue>>
}

export interface NativeRecalcPublicCorpusResult {
  readonly id: string
  readonly label: string
  readonly status: 'passed' | 'unsupported' | 'failed' | 'skipped'
  readonly sourceUrl: string
  readonly sha256: string
  readonly formulaCellCount: number
  readonly selectedFormulaCellCount: number
  readonly maxRssBytes: number
  readonly peakRssBytes: number | null
  readonly diagnosticsMaxObservedRssBytes?: number
  readonly patchedCacheCount?: number
  readonly nativeKernelFormulaCellCount?: number
  readonly unsupportedReason?: string
  readonly reason?: string
}

interface ChildRunResult {
  readonly exitCode: number | null
  readonly signal: string | null
  readonly stdout: string
  readonly stderr: string
  readonly peakRssBytes: number | null
  readonly rssLimitExceededBytes?: number
  readonly timedOut: boolean
}

interface NativeRecalcPublicCorpusDiagnostics {
  readonly engineMode?: string
  readonly maxObservedRssBytes?: number
  readonly patchedCacheCount?: number
  readonly nativeKernelFormulaCellCount?: number
  readonly unsupportedReason?: string
}

async function main(): Promise<void> {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const outputDir = resolve(readStringArg('--output-dir', defaultOutputDir))
  const cliPath = resolve(readStringArg('--cli', defaultCliPath))
  const nodeBin = readStringArg('--node-bin', process.env['BILIG_XLSX_NATIVE_RECALC_NODE'] ?? 'node')
  const limit = readNumberArg('--limit', defaultLimit)
  const maxFormulaReadsPerWorkbook = readNumberArg('--max-formula-reads-per-workbook', defaultMaxFormulaReadsPerWorkbook)
  const maxRssBytes = readMegabytesArg('--max-rss-mb', defaultMaxRssBytes)
  const timeoutMs = readNumberArg('--timeout-ms', defaultTimeoutMs)
  const requireFormulaWorkbookCount = readNumberArg('--require-formula-workbook-count', 1)
  const requirePassed = readFlagArg('--require-passed')
  const dryRun = readFlagArg('--dry-run')
  mkdirSync(outputDir, { recursive: true })

  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8')))
  const discoveredTargets = discoverNativeRecalcPublicCorpusTargets({
    cacheDir,
    manifest,
    outputDir,
    maxRssBytes,
    limit,
    maxFormulaReadsPerWorkbook,
  })
  const results: NativeRecalcPublicCorpusResult[] = []
  if (!existsSync(cliPath) && !dryRun) {
    results.push(failedSetupResult('xlsx-native-recalc-cli', 'xlsx-recalc dist CLI', maxRssBytes, `CLI not found: ${cliPath}`))
  } else {
    for (const target of discoveredTargets.targets) {
      results.push(
        dryRun
          ? skippedResult(target, 'dry run')
          : // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential Node workers keep host RSS bounded and attribution clear.
            await runNativeRecalcPublicCorpusTarget(target, { cliPath, nodeBin, timeoutMs }),
      )
    }
  }

  const summary = summarizeNativeRecalcPublicCorpusResults({
    discoveredFormulaWorkbookCount: discoveredTargets.discoveredFormulaWorkbookCount,
    attemptedFormulaWorkbookCount: discoveredTargets.targets.length,
    requireFormulaWorkbookCount,
    results,
  })
  process.stdout.write(
    `${JSON.stringify(
      {
        mode: 'xlsx-native-recalc-public-corpus',
        targets: {
          manifestPath,
          cacheDir,
          maxRssBytes,
          limit,
          maxFormulaReadsPerWorkbook,
          requireFormulaWorkbookCount,
          requirePassed,
          dryRun,
        },
        summary,
        results,
      },
      null,
      2,
    )}\n`,
  )

  const failed =
    summary.discoveredFormulaWorkbookCount < requireFormulaWorkbookCount ||
    (requirePassed && results.some((result) => result.status !== 'passed'))
  if (failed) {
    throw new Error('Native XLSX public corpus recalc gate failed')
  }
}

export function discoverNativeRecalcPublicCorpusTargets(args: {
  readonly manifest: PublicWorkbookManifest
  readonly cacheDir: string
  readonly outputDir: string
  readonly maxRssBytes: number
  readonly limit: number
  readonly maxFormulaReadsPerWorkbook: number
}): {
  readonly discoveredFormulaWorkbookCount: number
  readonly targets: readonly NativeRecalcPublicCorpusTarget[]
} {
  const targets: NativeRecalcPublicCorpusTarget[] = []
  let discoveredFormulaWorkbookCount = 0
  for (const artifact of args.manifest.artifacts) {
    const inputPath = join(args.cacheDir, artifact.cachePath)
    if (!existsSync(inputPath)) {
      continue
    }
    const oracles = readFormulaOraclesForArtifact(artifact, inputPath)
    if (oracles.length === 0) {
      continue
    }
    discoveredFormulaWorkbookCount += 1
    if (targets.length >= args.limit) {
      continue
    }
    targets.push(
      buildNativeRecalcPublicCorpusTarget({
        artifact,
        inputPath,
        outputDir: args.outputDir,
        maxRssBytes: args.maxRssBytes,
        oracles,
        maxFormulaReadsPerWorkbook: args.maxFormulaReadsPerWorkbook,
      }),
    )
  }
  return { discoveredFormulaWorkbookCount, targets }
}

function readFormulaOraclesForArtifact(artifact: PublicWorkbookArtifact, inputPath: string): readonly FormulaOracle[] {
  try {
    if (!isOpenXmlWorkbookFileName(artifact.fileName)) {
      return []
    }
    return formulaOraclesFromNativeFormulaCacheCells(readXlsxFormulaCacheCellsFromFile(inputPath).cells)
  } catch {
    return []
  }
}

export function formulaOraclesFromNativeFormulaCacheCells(cells: readonly XlsxFormulaCacheCell[]): readonly FormulaOracle[] {
  const oracles: FormulaOracle[] = []
  for (const cell of cells) {
    if (cell.cachedValue === undefined) {
      continue
    }
    const target = splitFormulaCacheTarget(cell.target)
    const expected = cellValueFromFormulaCacheLiteral(cell.cachedValue)
    if (!target || !expected) {
      continue
    }
    oracles.push({
      sheetName: target.sheetName,
      address: target.address,
      expected,
    })
  }
  return oracles
}

export function cellValueFromFormulaCacheLiteral(value: XlsxFormulaCacheLiteral): CellValue | null {
  if (value === null) {
    return { tag: ValueTag.Empty }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { tag: ValueTag.Number, value } : null
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  if (typeof value === 'string') {
    return errorCodeFromFormulaCacheText(value) ? null : { tag: ValueTag.String, value, stringId: 0 }
  }
  return null
}

function splitFormulaCacheTarget(target: string): { readonly sheetName: string; readonly address: string } | null {
  const separatorIndex = findFormulaCacheTargetSeparator(target)
  if (separatorIndex <= 0 || separatorIndex === target.length - 1) {
    return null
  }
  const sheetName = unquoteFormulaCacheSheetName(target.slice(0, separatorIndex))
  const address = target.slice(separatorIndex + 1)
  return sheetName && /^[A-Z]+[1-9][0-9]*$/u.test(address) ? { sheetName, address } : null
}

function findFormulaCacheTargetSeparator(target: string): number {
  let inQuotedSheetName = false
  for (let index = 0; index < target.length; index += 1) {
    const char = target[index]
    if (char === "'") {
      if (inQuotedSheetName && target[index + 1] === "'") {
        index += 1
        continue
      }
      inQuotedSheetName = !inQuotedSheetName
      continue
    }
    if (char === '!' && !inQuotedSheetName) {
      return index
    }
  }
  return -1
}

function unquoteFormulaCacheSheetName(value: string): string | null {
  if (!value.startsWith("'")) {
    return value.length > 0 ? value : null
  }
  if (!value.endsWith("'") || value.length < 2) {
    return null
  }
  return value.slice(1, -1).replaceAll("''", "'")
}

function errorCodeFromFormulaCacheText(value: string): ErrorCode | null {
  switch (value.toUpperCase()) {
    case '#DIV/0!':
      return ErrorCode.Div0
    case '#REF!':
      return ErrorCode.Ref
    case '#VALUE!':
      return ErrorCode.Value
    case '#NAME?':
      return ErrorCode.Name
    case '#N/A':
      return ErrorCode.NA
    case '#NUM!':
      return ErrorCode.Num
    case '#FIELD!':
      return ErrorCode.Field
    case '#NULL!':
      return ErrorCode.Null
    default:
      return null
  }
}

function isOpenXmlWorkbookFileName(fileName: string): boolean {
  const normalized = fileName.toLowerCase()
  return normalized.endsWith('.xlsx') || normalized.endsWith('.xlsm') || normalized.endsWith('.xltx') || normalized.endsWith('.xltm')
}

export function buildNativeRecalcPublicCorpusTarget(args: {
  readonly artifact: PublicWorkbookArtifact
  readonly inputPath: string
  readonly outputDir: string
  readonly maxRssBytes: number
  readonly oracles: readonly FormulaOracle[]
  readonly maxFormulaReadsPerWorkbook: number
}): NativeRecalcPublicCorpusTarget {
  const selectedOracles = args.oracles.slice(0, args.maxFormulaReadsPerWorkbook)
  const expectedReads: Record<string, CellValue> = {}
  const reads = selectedOracles.map((oracle) => {
    const target = formulaOracleReadTarget(oracle)
    expectedReads[target] = oracle.expected
    return target
  })
  return {
    id: args.artifact.id,
    label: args.artifact.fileName,
    sourceUrl: args.artifact.sourceUrl,
    sha256: args.artifact.sha256,
    inputPath: args.inputPath,
    outputPath: join(args.outputDir, `${args.artifact.id}-${args.artifact.sha256.slice(0, 12)}.native.xlsx`),
    maxRssBytes: args.maxRssBytes,
    formulaCellCount: args.oracles.length,
    selectedFormulaCellCount: selectedOracles.length,
    reads,
    expectedReads,
  }
}

export function formulaOracleReadTarget(oracle: FormulaOracle): string {
  return `${quoteSheetNameForTarget(oracle.sheetName)}!${oracle.address}`
}

export function quoteSheetNameForTarget(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/u.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

export function buildNativeRecalcPublicCorpusCliArgs(target: NativeRecalcPublicCorpusTarget): string[] {
  const args = [
    target.inputPath,
    '--out',
    target.outputPath,
    '--engine',
    'streaming-native',
    '--fallback-policy',
    'error',
    '--max-rss-bytes',
    String(target.maxRssBytes),
    '--json',
  ]
  for (const read of target.reads) {
    args.push('--read', read)
  }
  return args
}

export async function runNativeRecalcPublicCorpusTarget(
  target: NativeRecalcPublicCorpusTarget,
  options: { readonly cliPath: string; readonly nodeBin: string; readonly timeoutMs: number },
): Promise<NativeRecalcPublicCorpusResult> {
  const child = await runChildProcess(options.nodeBin, [options.cliPath, ...buildNativeRecalcPublicCorpusCliArgs(target)], {
    maxRssBytes: target.maxRssBytes,
    timeoutMs: options.timeoutMs,
  })
  const peakRssBytes = child.peakRssBytes
  if (child.rssLimitExceededBytes !== undefined) {
    return failedRuntimeResult(
      target,
      peakRssBytes,
      `peak RSS ${formatByteSize(child.rssLimitExceededBytes)} exceeded ${formatByteSize(target.maxRssBytes)}`,
    )
  }
  if (child.timedOut) {
    return failedRuntimeResult(target, peakRssBytes, `worker timed out after ${String(options.timeoutMs)}ms`)
  }
  const summary = parseCliSummary(child.stdout)
  if (child.exitCode !== 0) {
    const diagnostics = summary ? readDiagnostics(summary) : null
    const unsupportedReason = diagnostics?.unsupportedReason
    return unsupportedReason
      ? unsupportedResult(target, peakRssBytes, unsupportedReason, diagnostics)
      : failedRuntimeResult(
          target,
          peakRssBytes,
          `worker exited with ${child.signal ? `signal ${child.signal}` : `code ${String(child.exitCode)}`}: ${compactOutput(child.stderr || child.stdout)}`,
        )
  }
  if (!summary) {
    return failedRuntimeResult(target, peakRssBytes, 'worker returned invalid JSON')
  }
  const diagnostics = readDiagnostics(summary)
  if (diagnostics.engineMode !== 'streaming-native') {
    return failedRuntimeResult(
      target,
      peakRssBytes,
      `expected streaming-native diagnostics, received ${diagnostics.engineMode ?? 'missing'}`,
    )
  }
  const mismatch = firstMismatchedRead(summary, target)
  if (mismatch) {
    return failedRuntimeResult(target, peakRssBytes, mismatch)
  }
  return {
    ...baseResult(target, 'passed', peakRssBytes),
    ...(diagnostics.maxObservedRssBytes === undefined ? {} : { diagnosticsMaxObservedRssBytes: diagnostics.maxObservedRssBytes }),
    ...(diagnostics.patchedCacheCount === undefined ? {} : { patchedCacheCount: diagnostics.patchedCacheCount }),
    ...(diagnostics.nativeKernelFormulaCellCount === undefined
      ? {}
      : { nativeKernelFormulaCellCount: diagnostics.nativeKernelFormulaCellCount }),
  }
}

export function summarizeNativeRecalcPublicCorpusResults(args: {
  readonly discoveredFormulaWorkbookCount: number
  readonly attemptedFormulaWorkbookCount: number
  readonly requireFormulaWorkbookCount: number
  readonly results: readonly NativeRecalcPublicCorpusResult[]
}): {
  readonly discoveredFormulaWorkbookCount: number
  readonly attemptedFormulaWorkbookCount: number
  readonly requiredFormulaWorkbookCount: number
  readonly passedWorkbookCount: number
  readonly unsupportedWorkbookCount: number
  readonly failedWorkbookCount: number
  readonly selectedFormulaCellCount: number
  readonly nativeKernelFormulaCellCount: number
  readonly patchedCacheCount: number
  readonly maxPeakRssBytes: number | null
} {
  return {
    discoveredFormulaWorkbookCount: args.discoveredFormulaWorkbookCount,
    attemptedFormulaWorkbookCount: args.attemptedFormulaWorkbookCount,
    requiredFormulaWorkbookCount: args.requireFormulaWorkbookCount,
    passedWorkbookCount: args.results.filter((result) => result.status === 'passed').length,
    unsupportedWorkbookCount: args.results.filter((result) => result.status === 'unsupported').length,
    failedWorkbookCount: args.results.filter((result) => result.status === 'failed').length,
    selectedFormulaCellCount: args.results.reduce((sum, result) => sum + result.selectedFormulaCellCount, 0),
    nativeKernelFormulaCellCount: args.results.reduce((sum, result) => sum + (result.nativeKernelFormulaCellCount ?? 0), 0),
    patchedCacheCount: args.results.reduce((sum, result) => sum + (result.patchedCacheCount ?? 0), 0),
    maxPeakRssBytes: maxNullable(args.results.map((result) => result.peakRssBytes)),
  }
}

function firstMismatchedRead(summary: Readonly<Record<string, unknown>>, target: NativeRecalcPublicCorpusTarget): string | null {
  const reads = asRecord(summary['reads'])
  for (const read of target.reads) {
    const actual = cellValueFromCliRead(reads[read])
    const expected = target.expectedReads[read]
    if (actual && expected && cellValuesMatchOracle(actual, expected)) {
      continue
    }
    return `read ${read} did not match cached formula oracle`
  }
  return null
}

export function cellValueFromCliRead(value: unknown): CellValue | null {
  const record = asRecord(value)
  const tag = record['tag']
  switch (tag) {
    case 0:
      return { tag: ValueTag.Empty }
    case 1: {
      const number = record['value']
      return typeof number === 'number' && Number.isFinite(number) ? { tag: ValueTag.Number, value: number } : null
    }
    case 2: {
      const boolean = record['value']
      return typeof boolean === 'boolean' ? { tag: ValueTag.Boolean, value: boolean } : null
    }
    case 3: {
      const string = record['value']
      return typeof string === 'string' ? { tag: ValueTag.String, value: string, stringId: 0 } : null
    }
    case 4: {
      const code = record['code']
      return typeof code === 'number' && isErrorCode(code) ? { tag: ValueTag.Error, code } : null
    }
    default:
      return null
  }
}

function isErrorCode(value: number): value is ProtocolErrorCode {
  return Number.isInteger(value) && value >= 0 && value <= 11
}

function readDiagnostics(summary: Readonly<Record<string, unknown>>): NativeRecalcPublicCorpusDiagnostics {
  const diagnostics = asRecord(summary['diagnostics'])
  const formulaCounts = asRecord(diagnostics['formulaCounts'])
  return {
    ...(typeof diagnostics['engineMode'] === 'string' ? { engineMode: diagnostics['engineMode'] } : {}),
    ...(typeof diagnostics['maxObservedRssBytes'] === 'number' ? { maxObservedRssBytes: diagnostics['maxObservedRssBytes'] } : {}),
    ...(typeof diagnostics['patchedCacheCount'] === 'number' ? { patchedCacheCount: diagnostics['patchedCacheCount'] } : {}),
    ...(typeof diagnostics['unsupportedReason'] === 'string' ? { unsupportedReason: diagnostics['unsupportedReason'] } : {}),
    ...(typeof formulaCounts['nativeKernelFormulaCellCount'] === 'number'
      ? { nativeKernelFormulaCellCount: formulaCounts['nativeKernelFormulaCellCount'] }
      : {}),
  }
}

function parseCliSummary(stdout: string): Readonly<Record<string, unknown>> | null {
  try {
    return asRecord(JSON.parse(stdout.trim()))
  } catch {
    return null
  }
}

function runChildProcess(
  command: string,
  args: readonly string[],
  options: { readonly maxRssBytes: number; readonly timeoutMs: number },
): Promise<ChildRunResult> {
  return new Promise((resolvePromise) => {
    const child = spawn(command, args, {
      cwd: rootDir,
      detached: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let peakRssBytes = 0
    let settled = false
    let rssLimitExceededBytes: number | undefined
    let timedOut = false
    let stopRssWatchdog = noop
    const finish = (result: Omit<ChildRunResult, 'stdout' | 'stderr' | 'peakRssBytes' | 'rssLimitExceededBytes' | 'timedOut'>): void => {
      if (settled) {
        return
      }
      settled = true
      clearTimeout(timer)
      stopRssWatchdog()
      // oxlint-disable-next-line eslint-plugin-promise(no-multiple-resolved) -- `settled` gates close/error/watchdog races before resolving.
      resolvePromise({
        ...result,
        stdout,
        stderr,
        peakRssBytes: peakRssBytes || null,
        ...(rssLimitExceededBytes === undefined ? {} : { rssLimitExceededBytes }),
        timedOut,
      })
    }
    stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes: options.maxRssBytes,
      intervalMs: rssCheckIntervalMs,
      onSample: (rssBytes) => {
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
      },
      onLimitExceeded: (rssBytes) => {
        rssLimitExceededBytes = rssBytes
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
        terminateChildProcess(child, 'SIGTERM', { processGroup: true })
      },
    })
    const timer = setTimeout(() => {
      timedOut = true
      terminateChildProcess(child, 'SIGTERM', { processGroup: true })
    }, options.timeoutMs)
    timer.unref()
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
    })
    child.on('error', (error) => {
      stderr += error.message
      finish({ exitCode: 1, signal: null })
    })
    child.on('close', (exitCode, signal) => finish({ exitCode, signal }))
  })
}

function baseResult(
  target: NativeRecalcPublicCorpusTarget,
  status: NativeRecalcPublicCorpusResult['status'],
  peakRssBytes: number | null,
): Omit<NativeRecalcPublicCorpusResult, 'reason' | 'unsupportedReason'> {
  return {
    id: target.id,
    label: target.label,
    status,
    sourceUrl: target.sourceUrl,
    sha256: target.sha256,
    formulaCellCount: target.formulaCellCount,
    selectedFormulaCellCount: target.selectedFormulaCellCount,
    maxRssBytes: target.maxRssBytes,
    peakRssBytes,
  }
}

function unsupportedResult(
  target: NativeRecalcPublicCorpusTarget,
  peakRssBytes: number | null,
  unsupportedReason: string,
  diagnostics: NativeRecalcPublicCorpusDiagnostics,
): NativeRecalcPublicCorpusResult {
  return {
    ...baseResult(target, 'unsupported', peakRssBytes),
    unsupportedReason,
    ...(diagnostics.maxObservedRssBytes === undefined ? {} : { diagnosticsMaxObservedRssBytes: diagnostics.maxObservedRssBytes }),
    ...(diagnostics.patchedCacheCount === undefined ? {} : { patchedCacheCount: diagnostics.patchedCacheCount }),
    ...(diagnostics.nativeKernelFormulaCellCount === undefined
      ? {}
      : { nativeKernelFormulaCellCount: diagnostics.nativeKernelFormulaCellCount }),
  }
}

function failedRuntimeResult(
  target: NativeRecalcPublicCorpusTarget,
  peakRssBytes: number | null,
  reason: string,
): NativeRecalcPublicCorpusResult {
  return {
    ...baseResult(target, 'failed', peakRssBytes),
    reason,
  }
}

function failedSetupResult(id: string, label: string, maxRssBytes: number, reason: string): NativeRecalcPublicCorpusResult {
  return {
    id,
    label,
    status: 'failed',
    sourceUrl: '',
    sha256: '',
    formulaCellCount: 0,
    selectedFormulaCellCount: 0,
    maxRssBytes,
    peakRssBytes: null,
    reason,
  }
}

function skippedResult(target: NativeRecalcPublicCorpusTarget, reason: string): NativeRecalcPublicCorpusResult {
  return {
    ...baseResult(target, 'skipped', null),
    reason,
  }
}

function maxNullable(values: readonly (number | null)[]): number | null {
  let max: number | null = null
  for (const value of values) {
    if (value !== null) {
      max = max === null ? value : Math.max(max, value)
    }
  }
  return max
}

function compactOutput(value: string): string {
  const compacted = value.replace(/\s+/gu, ' ').trim()
  return compacted.length > 500 ? `${compacted.slice(0, 500)}...` : compacted
}

function asRecord(value: unknown): Readonly<Record<string, unknown>> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error('expected object')
  }
  const record: Record<string, unknown> = {}
  for (const [key, entry] of Object.entries(value)) {
    record[key] = entry
  }
  return record
}

if (import.meta.url === `file://${process.argv[1]}`) {
  await main()
}
