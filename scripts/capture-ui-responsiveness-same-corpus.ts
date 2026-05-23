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
  type EmitXlsxArgs,
} from './ui-responsiveness-same-corpus-args.ts'
import type { SameCorpusCapture } from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import {
  captureSameCorpusUiResponsiveness,
  preflightSameCorpusIncumbentAccess,
  saveStorageState,
} from './ui-responsiveness-same-corpus-page.ts'

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
    return
  }

  const args = parseCaptureArgs(process.argv.slice(2))
  assertSameCorpusBrowserRunAllowed()
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
}

export function assertSameCorpusBrowserRunAllowed(
  rootDirForGuard: string = rootDir,
  env: Readonly<Record<string, string | undefined>> = process.env,
): void {
  assertLocalCiResourceGuardAllowsRun(rootDirForGuard, env, { runLabel: 'same-corpus UI browser capture' })
}

export function assertSameCorpusCaptureEvidenceReady(capture: SameCorpusCapture): void {
  const blockingReasons = capture.runManifest.invalidReasons.filter(
    (reason) => reason !== 'not every required workload is 10x against Google Sheets',
  )
  if (blockingReasons.length === 0) {
    return
  }
  throw new Error(
    [
      'Same-corpus UI capture artifact is not valid evidence for the dominance scorecard.',
      'The artifact was written for diagnosis, but the capture command exits non-zero until browser-visible proof satisfies the current contract.',
      'Use --allow-incomplete-evidence only for exploratory captures that must not be fed into a public 10x claim.',
      ...blockingReasons.map((reason) => `- ${reason}`),
    ].join('\n'),
  )
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
          'Use --bilig-url <production-bilig-url> for dominance evidence. The default localhost dev URL is only valid with --allow-incomplete-evidence.',
        preflightCommand:
          'pnpm ui:same-corpus:capture -- --preflight --google-sheets-url <url> --microsoft-excel-web-url <url> [--google-sheets-storage-state <state.json>] [--microsoft-excel-web-storage-state <state.json>]',
        captureCommand:
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
