#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { buildWorkbookBenchmarkCorpus } from '../packages/benchmarks/src/workbook-corpus.js'
import { exportXlsx } from '../packages/excel-import/src/index.js'
import { assertLocalCiResourceGuardAllowsRun } from './ci-local-resource-guard.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import {
  parseCaptureArgs,
  parseEmitXlsxArgs,
  parsePreflightArgs,
  parseSaveStorageStateArgs,
  productionBiligSameCorpusUrl,
  type CaptureArgs,
  type EmitXlsxArgs,
  type PreflightArgs,
} from './ui-responsiveness-same-corpus-args.ts'
import type { SameCorpusCapture } from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import {
  captureSameCorpusUiResponsiveness,
  preflightSameCorpusIncumbentAccess,
  saveStorageState,
} from './ui-responsiveness-same-corpus-page.ts'
import type { SameCorpusPreflight } from './ui-responsiveness-same-corpus-preflight.ts'
import { sameCorpusPreflightProductInvalidReasons } from './ui-responsiveness-same-corpus-preflight.ts'

export { parseCaptureArgs, parseEmitXlsxArgs, parsePreflightArgs, parseSaveStorageStateArgs } from './ui-responsiveness-same-corpus-args.ts'
export { buildSameCorpusFingerprint, verifyXlsxCorpusFingerprint } from './ui-responsiveness-same-corpus-fingerprint.ts'
export {
  buildSameCorpusCaptureArtifact,
  captureSameCorpusUiResponsiveness,
  collectSameCorpusProductMeasurements,
  preflightSameCorpusIncumbentAccess,
  saveStorageState,
} from './ui-responsiveness-same-corpus-page.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)

async function main(): Promise<void> {
  const saveStorageStateArgs = parseSaveStorageStateArgs(process.argv.slice(2))
  if (saveStorageStateArgs) {
    assertSameCorpusBrowserRunAllowed()
    await saveStorageState(saveStorageStateArgs)
    return
  }

  const emitXlsxArgs = parseEmitXlsxArgs(process.argv.slice(2))
  if (emitXlsxArgs) {
    emitSameCorpusXlsx(emitXlsxArgs)
    return
  }

  const preflightArgs = parsePreflightArgs(process.argv.slice(2))
  if (preflightArgs) {
    assertSameCorpusBrowserRunAllowed()
    const preflight = await preflightSameCorpusIncumbentAccess(preflightArgs)
    const serializedJson = `${JSON.stringify(preflight, null, 2)}\n`
    if (preflightArgs.outputPath) {
      mkdirSync(dirname(preflightArgs.outputPath), { recursive: true })
      writeFileSync(
        preflightArgs.outputPath,
        formatJsonForRepo({
          rootDir,
          serializedJson,
          tempPrefix: 'ui-responsiveness-same-corpus-preflight',
        }),
      )
    }
    console.log(serializedJson.trim())
    assertSameCorpusPreflightReady(preflight)
    return
  }

  const args = parseCaptureArgs(process.argv.slice(2))
  assertProductionBiligEvidenceSource(args)
  assertSameCorpusBrowserRunAllowed()
  await assertClaimGradeCaptureIncumbentsReady(args)
  const servedBilig = args.biligUrlSource === 'served-production' ? await startServedBiligProductionRuntime(args) : null
  try {
    const capture = await captureSameCorpusUiResponsiveness(args)
    mkdirSync(dirname(args.outputPath), { recursive: true })
    writeFileSync(
      args.outputPath,
      formatJsonForRepo({
        rootDir,
        serializedJson: `${JSON.stringify(capture, null, 2)}\n`,
        tempPrefix: 'ui-responsiveness-same-corpus-capture',
      }),
    )
    console.log(
      JSON.stringify(
        {
          outputPath: args.outputPath,
          corpusCaseId: args.corpusId,
          sampleCount: args.sampleCount,
          workloads: capture.cases.map((entry) => entry.workload),
          currentContractEvidenceComplete: capture.runManifest.currentContractEvidenceComplete,
          googleSheetsTenXRequirementSatisfied: capture.runManifest.googleSheetsTenXRequirementSatisfied,
        },
        null,
        2,
      ),
    )
    if (!args.allowIncompleteEvidence) {
      assertSameCorpusCaptureEvidenceReady(capture)
    }
  } finally {
    if (servedBilig) {
      await stopServedBiligProductionRuntime(servedBilig)
    }
  }
}

export function assertSameCorpusBrowserRunAllowed(
  rootDirForGuard: string = rootDir,
  env: Readonly<Record<string, string | undefined>> = process.env,
): void {
  assertLocalCiResourceGuardAllowsRun(rootDirForGuard, env, { runLabel: 'same-corpus UI browser capture' })
}

export function assertSameCorpusCaptureEvidenceReady(capture: SameCorpusCapture): void {
  assertSameCorpusCaptureClaimReady(capture)
}

export function assertSameCorpusPreflightReady(preflight: SameCorpusPreflight): void {
  const invalidLines = preflight.products.flatMap((product) =>
    sameCorpusPreflightProductInvalidReasons(product).map((reason) => `- ${product.product}: ${reason}`),
  )
  if (preflight.allCheckedProductsReady && invalidLines.length === 0) {
    return
  }
  throw new Error(['Same-corpus UI incumbent preflight is not ready for claim-grade capture.', ...invalidLines].join('\n'))
}

type SameCorpusPreflightRunner = (args: PreflightArgs) => Promise<SameCorpusPreflight>

export function preflightArgsForClaimGradeCapture(args: CaptureArgs): PreflightArgs | null {
  if (args.allowIncompleteEvidence) {
    return null
  }
  return {
    corpusId: args.corpusId,
    googleSheetsUrl: args.googleSheetsUrl,
    googleSheetsStorageStatePath: args.googleSheetsStorageStatePath,
    headless: args.headless,
    microsoftExcelWebUrl: args.microsoftExcelWebUrl,
    microsoftExcelWebStorageStatePath: args.microsoftExcelWebStorageStatePath,
    outputPath: null,
    readyTimeoutMs: args.readyTimeoutMs,
    storageStatePath: args.storageStatePath,
  }
}

export async function assertClaimGradeCaptureIncumbentsReady(
  args: CaptureArgs,
  runPreflight: SameCorpusPreflightRunner = preflightSameCorpusIncumbentAccess,
): Promise<void> {
  const preflightArgs = preflightArgsForClaimGradeCapture(args)
  if (!preflightArgs) {
    return
  }
  const preflight = await runPreflight(preflightArgs)
  assertSameCorpusPreflightReady(preflight)
}

export function assertSameCorpusCaptureCurrentContractEvidenceReady(capture: SameCorpusCapture): void {
  const currentContractBlockingReasons = capture.runManifest.invalidReasons.filter(
    (reason) => reason !== 'not every required workload is 10x against Google Sheets',
  )
  if (currentContractBlockingReasons.length === 0) {
    return
  }
  throw new Error(
    [
      'Same-corpus UI capture artifact is not valid evidence for the dominance scorecard.',
      'The artifact was written for diagnosis, but the capture command exits non-zero until browser-visible proof satisfies the current contract.',
      'Use --allow-incomplete-evidence only for exploratory captures that must not be fed into a public 10x claim.',
      ...currentContractBlockingReasons.map((reason) => `- ${reason}`),
    ].join('\n'),
  )
}

export function assertSameCorpusCaptureClaimReady(capture: SameCorpusCapture): void {
  if (capture.runManifest.invalidReasons.length === 0) {
    return
  }
  throw new Error(
    [
      'Same-corpus UI capture artifact is not valid claim-grade Google Sheets 10x evidence.',
      'The artifact was written for diagnosis, but the capture command exits non-zero until every claimed workload has current proof and 10x mean+p95 against live Google Sheets.',
      'Use --allow-incomplete-evidence only for exploratory captures that must not be fed into a public 10x claim.',
      ...capture.runManifest.invalidReasons.map((reason) => `- ${reason}`),
    ].join('\n'),
  )
}

export function assertProductionBiligEvidenceSource(args: CaptureArgs): void {
  if (args.allowIncompleteEvidence || args.biligUrlSource !== 'default-dev') {
    return
  }
  throw new Error(
    [
      'Same-corpus UI capture needs production Bilig runtime proof.',
      'The default Bilig URL is a localhost dev server and cannot satisfy the dominance scorecard.',
      'Use --serve-bilig-production to build and serve the production web bundle for this capture, or pass --bilig-url <production-bilig-url>.',
      'Use --allow-incomplete-evidence only for diagnostic captures that must not be fed into a public 10x claim.',
    ].join('\n'),
  )
}

interface ServedBiligProductionRuntime {
  readonly process: ReturnType<typeof Bun.spawn>
  readonly url: string
}

async function startServedBiligProductionRuntime(args: CaptureArgs): Promise<ServedBiligProductionRuntime> {
  const url = productionBiligSameCorpusUrl(args.biligProductionHost, args.biligProductionPort, args.corpusId)
  console.log(`Building @bilig/web production bundle for same-corpus capture...`)
  const build = Bun.spawnSync(['pnpm', '--filter', '@bilig/web', 'build'], {
    cwd: rootDir,
    stdin: 'ignore',
    stdout: 'inherit',
    stderr: 'inherit',
  })
  if (build.exitCode !== 0) {
    throw new Error(`Failed to build @bilig/web production bundle for same-corpus capture: exit ${String(build.exitCode)}`)
  }

  console.log(`Serving @bilig/web production bundle for same-corpus capture at ${url}...`)
  const process = Bun.spawn(
    [
      'pnpm',
      '--dir',
      join(rootDir, 'apps/web'),
      'exec',
      'vite',
      'preview',
      '--host',
      args.biligProductionHost,
      '--port',
      String(args.biligProductionPort),
      '--strictPort',
    ],
    {
      cwd: rootDir,
      stdin: 'ignore',
      stdout: 'inherit',
      stderr: 'inherit',
    },
  )
  await waitForServedBiligProductionRuntime(process, url, args.readyTimeoutMs)
  return { process, url }
}

async function waitForServedBiligProductionRuntime(process: ReturnType<typeof Bun.spawn>, url: string, timeoutMs: number): Promise<void> {
  await pollServedBiligProductionRuntime(process, url, Date.now() + timeoutMs)
}

async function pollServedBiligProductionRuntime(process: ReturnType<typeof Bun.spawn>, url: string, deadline: number): Promise<void> {
  if (Date.now() >= deadline) {
    throw new Error(`Timed out waiting for production Bilig preview at ${url}`)
  }
  const exit = await Promise.race([process.exited.then((exitCode) => ({ exitCode })), sleep(250).then(() => null)])
  if (exit) {
    throw new Error(`Production Bilig preview exited before it was ready: exit ${String(exit.exitCode)}`)
  }
  try {
    const response = await fetch(url, { signal: AbortSignal.timeout(1_000) })
    if (response.ok) {
      return
    }
  } catch {
    // Keep polling until the preview server is ready or the timeout expires.
  }
  return pollServedBiligProductionRuntime(process, url, deadline)
}

async function stopServedBiligProductionRuntime(runtime: ServedBiligProductionRuntime): Promise<void> {
  try {
    runtime.process.kill('SIGTERM')
  } catch {
    return
  }
  const stopped = await Promise.race([runtime.process.exited.then(() => true), sleep(3_000).then(() => false)])
  if (!stopped) {
    runtime.process.kill('SIGKILL')
    await runtime.process.exited
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((done) => setTimeout(done, ms))
}

export function emitSameCorpusXlsx(args: EmitXlsxArgs): void {
  mkdirSync(args.targetDirectory, { recursive: true })
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const outputFile = join(args.targetDirectory, `${args.corpusId}.xlsx`)
  const workbookBytes = Buffer.from(exportXlsx(corpus.snapshot))
  if (args.check) {
    if (!existsSync(outputFile)) {
      throw new Error(`Same-corpus XLSX fixture is missing: ${outputFile}`)
    }
    const existingBytes = readFileSync(outputFile)
    if (!existingBytes.equals(workbookBytes)) {
      throw new Error(`Same-corpus XLSX fixture is stale: ${outputFile}`)
    }
  } else {
    writeFileSync(outputFile, workbookBytes)
  }

  const publicGithubRawUrl = `https://raw.githubusercontent.com/proompteng/bilig/main/packages/benchmarks/baselines/ui-same-corpus/${corpus.id}.xlsx`
  console.log(
    JSON.stringify(
      {
        mode: args.check ? 'check-xlsx' : 'emit-xlsx',
        outputFile,
        corpusCaseId: corpus.id,
        materializedCells: corpus.materializedCellCount,
        googleSheetsUploadMode: 'native_google_sheets',
        publicGithubRawUrl,
        publicForgejoRawUrl: `https://code.proompteng.ai/kalmyk/bilig/raw/branch/main/packages/benchmarks/baselines/ui-same-corpus/${corpus.id}.xlsx`,
        microsoftExcelWebUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(publicGithubRawUrl)}`,
        googleSheetsAuthStateCommand:
          'pnpm ui:same-corpus:capture -- --save-storage-state <state.json> --auth-product google-sheets --google-sheets-url <url> [--corpus wide-mixed-250k]',
        microsoftExcelWebAuthStateCommand:
          'pnpm ui:same-corpus:capture -- --save-storage-state <state.json> --auth-product microsoft-excel-web --microsoft-excel-web-url <url> [--corpus wide-mixed-250k]',
        biligProductionRuntimeRequirement:
          'Use --serve-bilig-production or --bilig-url <production-bilig-url> for dominance evidence. The default localhost dev URL is only valid with --allow-incomplete-evidence.',
        preflightCommand:
          'pnpm ui:same-corpus:capture -- --preflight --google-sheets-url <url> --microsoft-excel-web-url <url> [--google-sheets-storage-state <state.json>] [--microsoft-excel-web-storage-state <state.json>]',
        captureCommand:
          'pnpm ui:same-corpus:capture -- --serve-bilig-production --output <capture.json> --google-sheets-url <url> --microsoft-excel-web-url <url> [--google-sheets-storage-state <state.json>] [--microsoft-excel-web-storage-state <state.json>]',
        externalProductionCaptureCommand:
          'pnpm ui:same-corpus:capture -- --output <capture.json> --bilig-url <production-bilig-url> --google-sheets-url <url> --microsoft-excel-web-url <url> [--google-sheets-storage-state <state.json>] [--microsoft-excel-web-storage-state <state.json>]',
        diagnosticCaptureCommand:
          'pnpm ui:same-corpus:capture -- --output <capture.json> --google-sheets-url <url> --allow-incomplete-evidence',
      },
      null,
      2,
    ),
  )
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    await main()
  } catch (error: unknown) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exitCode = 1
  }
}
