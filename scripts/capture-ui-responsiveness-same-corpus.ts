#!/usr/bin/env bun

import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'
import { pathToFileURL } from 'node:url'

import { chromium, type Browser, type Page } from '@playwright/test'
import {
  getWorkbookBenchmarkCorpusDefinition,
  isWorkbookBenchmarkCorpusId,
  type WorkbookBenchmarkCorpusId,
} from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCapture,
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

interface CaptureArgs {
  readonly biligUrl: string
  readonly corpusId: WorkbookBenchmarkCorpusId
  readonly deltaX: number
  readonly deltaY: number
  readonly googleSheetsUrl: string
  readonly headless: boolean
  readonly microsoftExcelWebUrl: string
  readonly outputPath: string
  readonly sampleCount: number
}

interface ScrollSample {
  readonly operationResponseMs: number
  readonly postOperationFrameMs: number
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCorpusId: WorkbookBenchmarkCorpusId = 'wide-mixed-250k'
const defaultViewport = { width: 1440, height: 900 } as const

async function main(): Promise<void> {
  const args = parseCaptureArgs(process.argv.slice(2))
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
        workload: 'visible-scroll-response',
      },
      null,
      2,
    ),
  )
}

export function parseCaptureArgs(argv: readonly string[]): CaptureArgs {
  const corpusId = parseCorpusId(argumentValue(argv, '--corpus') ?? defaultCorpusId)
  const outputPath = argumentValue(argv, '--output')
  const googleSheetsUrl = argumentValue(argv, '--google-sheets-url')
  const microsoftExcelWebUrl = argumentValue(argv, '--microsoft-excel-web-url')
  if (!outputPath || !googleSheetsUrl || !microsoftExcelWebUrl) {
    throw new Error(
      [
        'Missing required arguments.',
        'Usage: bun scripts/capture-ui-responsiveness-same-corpus.ts',
        '  --output <capture.json>',
        '  --google-sheets-url <same-corpus-google-sheets-url>',
        '  --microsoft-excel-web-url <same-corpus-excel-web-url>',
        '  [--bilig-url <local-bilig-url>] [--corpus wide-mixed-250k] [--samples 3] [--delta-x 0] [--delta-y 720] [--headed]',
      ].join('\n'),
    )
  }
  const sampleCount = parsePositiveInteger(argumentValue(argv, '--samples') ?? '3', '--samples')
  return {
    biligUrl: argumentValue(argv, '--bilig-url') ?? `http://127.0.0.1:5173/?benchmarkCorpus=${encodeURIComponent(corpusId)}`,
    corpusId,
    deltaX: parseNonNegativeNumber(argumentValue(argv, '--delta-x') ?? '0', '--delta-x'),
    deltaY: parseNonNegativeNumber(argumentValue(argv, '--delta-y') ?? '720', '--delta-y'),
    googleSheetsUrl,
    headless: !argv.includes('--headed'),
    microsoftExcelWebUrl,
    outputPath: resolve(outputPath),
    sampleCount,
  }
}

export async function captureSameCorpusUiResponsiveness(args: CaptureArgs): Promise<SameCorpusCapture> {
  const corpus = getWorkbookBenchmarkCorpusDefinition(args.corpusId)
  const browser = await chromium.launch({ headless: args.headless })
  try {
    const [bilig, googleSheets, microsoftExcelWeb] = await Promise.all([
      measureProduct(browser, 'bilig', args.biligUrl, args),
      measureProduct(browser, 'google-sheets', args.googleSheetsUrl, args),
      measureProduct(browser, 'microsoft-excel-web', args.microsoftExcelWebUrl, args),
    ])
    return {
      schemaVersion: 1,
      suite: 'ui-responsiveness-same-corpus-capture',
      sampleCount: args.sampleCount,
      limitations: [
        'Caller must supply Google Sheets and Microsoft Excel Web URLs for the same exported Bilig benchmark corpus.',
        'This capture measures browser-visible scroll response; edit latency must be captured by a separate same-corpus workload.',
      ],
      cases: [
        {
          id: `same-corpus-${args.corpusId}-visible-scroll-response`,
          corpusCaseId: args.corpusId,
          materializedCells: corpus.materializedCellCount,
          workload: 'visible-scroll-response',
          bilig,
          googleSheets,
          microsoftExcelWeb,
        },
      ],
    }
  } finally {
    await browser.close()
  }
}

async function measureProduct(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  args: CaptureArgs,
): Promise<SameCorpusCaptureMeasurement> {
  const samples = await measureProductSamples(browser, product, url, args)

  return {
    product,
    source: url,
    operationResponseMsSamples: samples.map((entry) => entry.operationResponseMs),
    postOperationFrameMsSamples: samples.map((entry) => entry.postOperationFrameMs),
    limitations: productLimitations(product),
  }
}

async function measureProductSamples(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  args: CaptureArgs,
  sampleIndex = 0,
  samples: ScrollSample[] = [],
): Promise<ScrollSample[]> {
  if (sampleIndex >= args.sampleCount) {
    return samples
  }
  const page = await browser.newPage({ viewport: defaultViewport })
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, product, args.corpusId)
    samples.push(await measureVisibleScrollResponse(page, args.deltaX, args.deltaY))
  } finally {
    await page.close()
  }
  return measureProductSamples(browser, product, url, args, sampleIndex + 1, samples)
}

async function waitForProductReady(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  corpusId: WorkbookBenchmarkCorpusId,
): Promise<void> {
  if (product === 'bilig') {
    await page.waitForSelector('[data-testid="sheet-grid"]', { state: 'visible', timeout: 30_000 })
    await page.waitForFunction(
      (expectedCorpusId) => {
        const collector = (
          window as Window & {
            __biligScrollPerf?: {
              getBenchmarkState?: () => {
                state: string
                error: string | null
                fixture: { id: string; materializedCellCount: number; sheetName: string } | null
              }
            }
          }
        ).__biligScrollPerf
        const state = collector?.getBenchmarkState?.()
        return state?.state === 'ready' && state.fixture?.id === expectedCorpusId
      },
      corpusId,
      { timeout: 60_000 },
    )
    await settleFrames(page, 12)
    return
  }

  if (product === 'google-sheets') {
    await page.waitForFunction(
      () =>
        !window.location.href.includes('accounts.google.com') &&
        document.title.includes('Google Sheets') &&
        !document.body.innerText.includes('Sign in\nto continue to Google Sheets'),
      { timeout: 60_000 },
    )
    await settleFrames(page, 120)
    return
  }

  await page.waitForFunction(
    () => document.title.toLowerCase().includes('.xlsx') || document.body.innerText.toLowerCase().includes('excel'),
    { timeout: 60_000 },
  )
  await settleFrames(page, 180)
}

async function measureVisibleScrollResponse(page: Page, deltaX: number, deltaY: number): Promise<ScrollSample> {
  await page.mouse.move(defaultViewport.width / 2, defaultViewport.height / 2)
  const startedAt = performance.now()
  await page.mouse.wheel(deltaX, deltaY)
  const frameIntervals = await page.evaluate(async () => {
    const intervals: number[] = []
    let previous = performance.now()
    await new Promise<void>((finish) => {
      const step = (now: number): void => {
        intervals.push(now - previous)
        previous = now
        if (intervals.length >= 12) {
          finish()
          return
        }
        requestAnimationFrame(step)
      }
      requestAnimationFrame(step)
    })
    return intervals
  })
  const operationResponseMs = performance.now() - startedAt
  return {
    operationResponseMs,
    postOperationFrameMs: percentile(frameIntervals, 0.95),
  }
}

async function settleFrames(page: Page, frames: number): Promise<void> {
  await page.evaluate(async (frameCount) => {
    await Array.from({ length: frameCount }).reduce<Promise<void>>(async (previous) => {
      await previous
      await new Promise<void>((resolveFrame) => requestAnimationFrame(() => resolveFrame()))
    }, Promise.resolve())
  }, frames)
}

function productLimitations(product: UiResponsivenessSameCorpusProduct): string[] {
  if (product === 'bilig') {
    return ['Bilig timing is captured from the supplied local app URL and benchmarkCorpus route.']
  }
  if (product === 'google-sheets') {
    return ['Google Sheets timing requires the supplied URL to be browser-accessible and loaded with the same benchmark corpus.']
  }
  return ['Microsoft Excel Web timing requires the supplied URL to be browser-accessible and loaded with the same benchmark corpus.']
}

function percentile(values: readonly number[], percentileValue: number): number {
  if (values.length === 0) {
    throw new Error('Cannot compute percentile for an empty sample set')
  }
  const sorted = [...values].toSorted((left, right) => left - right)
  const index = Math.ceil(percentileValue * sorted.length) - 1
  return sorted[Math.max(0, Math.min(sorted.length - 1, index))]!
}

function parseCorpusId(value: string): WorkbookBenchmarkCorpusId {
  if (!isWorkbookBenchmarkCorpusId(value)) {
    throw new Error(`Unexpected workbook benchmark corpus id: ${value}`)
  }
  return value
}

function parsePositiveInteger(value: string, flag: string): number {
  const parsed = Number(value)
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new Error(`${flag} must be a positive integer`)
  }
  return parsed
}

function parseNonNegativeNumber(value: string, flag: string): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed < 0) {
    throw new Error(`${flag} must be a non-negative number`)
  }
  return parsed
}

function argumentValue(argv: readonly string[], name: string): string | null {
  const index = argv.indexOf(name)
  if (index === -1) {
    return null
  }
  const value = argv[index + 1]
  if (!value) {
    throw new Error(`Missing value after ${name}`)
  }
  return value
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    await main()
  } catch (error: unknown) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exitCode = 1
  }
}
