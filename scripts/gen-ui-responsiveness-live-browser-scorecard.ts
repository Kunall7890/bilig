#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import { chromium, type Browser, type Page } from '@playwright/test'
import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import { assertLocalCiResourceGuardAllowsRun } from './ci-local-resource-guard.ts'
import { readJsonObject } from './json-scorecard-helpers.ts'
import { parseSameCorpusCapture, parseUiResponsivenessLiveBrowserScorecard } from './ui-responsiveness-live-browser-scorecard-parse.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import {
  buildMissingSameCorpusProof,
  buildSameCorpusProof,
  validateSameCorpusProof,
  type UiResponsivenessSameCorpusProof,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'

export { parseSameCorpusCapture, parseUiResponsivenessLiveBrowserScorecard } from './ui-responsiveness-live-browser-scorecard-parse.ts'
export { buildMissingSameCorpusProof, buildSameCorpusProof } from './ui-responsiveness-same-corpus-scorecard-proof.ts'
export { buildSameCorpusCaptureRunManifest, sameCorpusScenarioCaseFields } from './ui-responsiveness-same-corpus-scorecard-proof.ts'
export type {
  SameCorpusCapture,
  SameCorpusBiligRuntimeProof,
  SameCorpusBiligRuntimeProofSample,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusFingerprint,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureRunManifest,
  SameCorpusCaptureMeasurement,
  SameCorpusCaptureVerifiedCell,
  SameCorpusOperationResponseProof,
  SameCorpusScenarioCaseFields,
  SameCorpusProductSourceWorkbookFingerprint,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
  UiResponsivenessSameCorpusProof,
  UiResponsivenessSameCorpusRunManifest,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'
export type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export type UiResponsivenessLiveBrowserVendor = 'google-sheets' | 'microsoft-excel-web'

export interface UiResponsivenessLiveBrowserCase {
  readonly id: string
  readonly vendor: UiResponsivenessLiveBrowserVendor
  readonly product: string
  readonly sourceUrl: string
  readonly finalUrl: string
  readonly title: string
  readonly accessMode: 'public-comment-only' | 'public-view-only' | 'public-office-web-viewer'
  readonly workload: 'open-public-workbook-and-scroll-viewport'
  readonly sampleCount: number
  readonly loadToReadyMs: NumericSummary
  readonly scrollResponseMs: NumericSummary
  readonly postScrollFrameMs: NumericSummary
  readonly passed: boolean
  readonly limitations: string[]
}

export interface UiResponsivenessLiveBrowserScorecard {
  readonly schemaVersion: 1
  readonly suite: 'ui-responsiveness-live-browser-timing'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts'
    readonly evidenceKind: 'live-public-browser-playwright'
    readonly browserEngine: 'chromium'
    readonly measuredOperation: 'public-workbook-load-and-viewport-scroll'
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly viewport: {
      readonly width: number
      readonly height: number
    }
    readonly samplingOrder: 'google-sheets-then-microsoft-excel-web'
  }
  readonly summary: {
    readonly directBrowserTimingCaptured: boolean
    readonly allRequiredCasesPassed: boolean
    readonly requiredVendorCount: number
    readonly capturedVendors: UiResponsivenessLiveBrowserVendor[]
    readonly limitations: string[]
  }
  readonly cases: UiResponsivenessLiveBrowserCase[]
  readonly sameCorpusProof: UiResponsivenessSameCorpusProof
}

export interface UiResponsivenessLiveBrowserCliArgs {
  readonly isCheckMode: boolean
  readonly capturePath: string | null
}

export interface SameCorpusScreenshotArtifactValidationOptions {
  readonly requireGitTracked?: boolean | undefined
  readonly rootDir?: string | undefined
  readonly trackedArtifactPaths?: readonly string[] | undefined
}

export interface SameCorpusCaptureArtifactValidationOptions {
  readonly capturePath?: string | undefined
  readonly rootDir?: string | undefined
}

interface BrowserCaseSpec {
  readonly id: string
  readonly vendor: UiResponsivenessLiveBrowserVendor
  readonly product: string
  readonly sourceUrl: string
  readonly expectedTitleIncludes: string
}

interface BrowserCaseSample {
  readonly finalUrl: string
  readonly title: string
  readonly accessMode: UiResponsivenessLiveBrowserCase['accessMode']
  readonly loadToReadyMs: number
  readonly scrollResponseMs: number
  readonly postScrollFrameMs: number
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'ui-responsiveness-live-browser-scorecard.json')
const defaultSameCorpusCapturePath = join(rootDir, '.cache', 'ui-responsiveness', 'same-corpus-capture.json')
const sampleCount = 3
const viewport = { width: 1440, height: 900 } as const
const microsoftExcelSourceWorkbook =
  'https://github.com/fileformat-blog-gists/SampleFiles/raw/main/Spreadsheet-File-Formats/XLSX/Pivot-Tables-and-Charts.xlsx'
const caseSpecs = [
  {
    id: 'google-sheets-public-grid-scroll',
    vendor: 'google-sheets',
    product: 'Google Sheets public spreadsheet',
    sourceUrl: 'https://docs.google.com/spreadsheets/d/1Awcx961Qm_cJw7X-7hsKEGCzS-0yMw0TqwbniaxNkeU/edit',
    expectedTitleIncludes: 'Google Sheets',
  },
  {
    id: 'microsoft-excel-web-public-xlsx-scroll',
    vendor: 'microsoft-excel-web',
    product: 'Microsoft Office Web Viewer public XLSX',
    sourceUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(microsoftExcelSourceWorkbook)}`,
    expectedTitleIncludes: '.xlsx',
  },
] as const satisfies readonly BrowserCaseSpec[]

async function main(): Promise<void> {
  const args = parseUiResponsivenessLiveBrowserCliArgs(process.argv.slice(2))
  if (args.isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `UI responsiveness live browser scorecard is missing. Run: bun scripts/gen-ui-responsiveness-live-browser-scorecard.ts`,
      )
    }
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(readJsonObject(outputPath))
    validateUiResponsivenessLiveBrowserScorecard(scorecard)
    validateSameCorpusCaptureArtifactMatchesScorecard(scorecard)
    validateSameCorpusScreenshotArtifacts(scorecard.sameCorpusProof, { requireGitTracked: true })
    logResult('check', scorecard)
    return
  }

  assertUiResponsivenessLiveBrowserRunAllowed()
  const sameCorpusProof = args.capturePath
    ? buildSameCorpusProof(parseSameCorpusCapture(readJsonObject(resolve(args.capturePath))))
    : buildMissingSameCorpusProof()
  const scorecard = await buildUiResponsivenessLiveBrowserScorecard(new Date().toISOString(), sameCorpusProof)
  validateUiResponsivenessLiveBrowserScorecard(scorecard)
  validateSameCorpusScreenshotArtifacts(scorecard.sameCorpusProof)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildUiResponsivenessLiveBrowserScorecard(
  generatedAt: string,
  sameCorpusProof = buildMissingSameCorpusProof(),
): Promise<UiResponsivenessLiveBrowserScorecard> {
  const browser = await chromium.launch({ headless: true })
  try {
    const cases = await measureBrowserCases(browser)

    return {
      schemaVersion: 1,
      suite: 'ui-responsiveness-live-browser-timing',
      generatedAt,
      host: {
        arch: process.arch,
        platform: process.platform,
      },
      source: {
        artifactGenerator: 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts',
        evidenceKind: 'live-public-browser-playwright',
        browserEngine: 'chromium',
        measuredOperation: 'public-workbook-load-and-viewport-scroll',
      },
      benchmark: {
        sampleCount,
        viewport,
        samplingOrder: 'google-sheets-then-microsoft-excel-web',
      },
      summary: {
        directBrowserTimingCaptured: cases.length === caseSpecs.length,
        allRequiredCasesPassed: cases.every((entry) => entry.passed),
        requiredVendorCount: caseSpecs.length,
        capturedVendors: caseSpecs.map((entry) => entry.vendor),
        limitations: [
          'Public unauthenticated browser timing covers load and viewport scroll only; it does not cover authenticated edit latency.',
          'The incumbent workbooks are public representative workbooks, not bilig-generated benchmark corpuses.',
          'Network, tenant, CDN, and browser-cache conditions can move live public-web measurements between runs.',
        ],
      },
      cases,
      sameCorpusProof,
    }
  } finally {
    await browser.close()
  }
}

export function assertUiResponsivenessLiveBrowserRunAllowed(
  rootDirForGuard: string = rootDir,
  env: Readonly<Record<string, string | undefined>> = process.env,
): void {
  assertLocalCiResourceGuardAllowsRun(rootDirForGuard, env, { runLabel: 'UI responsiveness live browser scorecard generation' })
}

export function parseUiResponsivenessLiveBrowserCliArgs(argv: readonly string[]): UiResponsivenessLiveBrowserCliArgs {
  return {
    isCheckMode: argv.includes('--check'),
    capturePath: argumentValue(argv, '--capture'),
  }
}

async function measureBrowserCases(
  browser: Browser,
  specIndex = 0,
  cases: UiResponsivenessLiveBrowserCase[] = [],
): Promise<UiResponsivenessLiveBrowserCase[]> {
  const spec = caseSpecs[specIndex]
  if (!spec) {
    return cases
  }
  cases.push(buildBrowserCase(spec, await measureBrowserCaseSamples(browser, spec)))
  return measureBrowserCases(browser, specIndex + 1, cases)
}

async function measureBrowserCaseSamples(
  browser: Browser,
  spec: BrowserCaseSpec,
  sampleIndex = 0,
  samples: BrowserCaseSample[] = [],
): Promise<BrowserCaseSample[]> {
  if (sampleIndex >= sampleCount) {
    return samples
  }
  const page = await browser.newPage({ viewport })
  try {
    samples.push(await measureBrowserCase(page, spec))
  } finally {
    await page.close()
  }
  return measureBrowserCaseSamples(browser, spec, sampleIndex + 1, samples)
}

export function validateUiResponsivenessLiveBrowserScorecard(scorecard: UiResponsivenessLiveBrowserScorecard): void {
  if (scorecard.benchmark.sampleCount < 3) {
    throw new Error('UI responsiveness live browser scorecard must contain at least 3 samples per case')
  }
  if (!scorecard.summary.directBrowserTimingCaptured || !scorecard.summary.allRequiredCasesPassed) {
    throw new Error('UI responsiveness live browser scorecard summary reports missing or failed browser timing evidence')
  }
  for (const spec of caseSpecs) {
    const entry = scorecard.cases.find((candidate) => candidate.id === spec.id)
    if (!entry) {
      throw new Error(`UI responsiveness live browser scorecard is missing required case: ${spec.id}`)
    }
    if (entry.vendor !== spec.vendor) {
      throw new Error(`UI responsiveness live browser scorecard vendor mismatch for case: ${spec.id}`)
    }
    if (!entry.passed) {
      throw new Error(`UI responsiveness live browser scorecard contains a failed case: ${spec.id}`)
    }
    if (entry.sampleCount < scorecard.benchmark.sampleCount) {
      throw new Error(`UI responsiveness live browser scorecard has too few samples for case: ${spec.id}`)
    }
    if (!entry.title.includes(spec.expectedTitleIncludes)) {
      throw new Error(`UI responsiveness live browser scorecard title does not match ${spec.vendor}: ${entry.title}`)
    }
    validateSummary(entry.loadToReadyMs, `${spec.id} loadToReadyMs`)
    validateSummary(entry.scrollResponseMs, `${spec.id} scrollResponseMs`)
    validateSummary(entry.postScrollFrameMs, `${spec.id} postScrollFrameMs`)
  }
  for (const vendor of caseSpecs.map((entry) => entry.vendor)) {
    if (!scorecard.summary.capturedVendors.includes(vendor)) {
      throw new Error(`UI responsiveness live browser scorecard is missing vendor: ${vendor}`)
    }
  }
  if (scorecard.summary.limitations.length === 0 || !scorecard.cases.every((entry) => entry.limitations.length > 0)) {
    throw new Error('UI responsiveness live browser scorecard must disclose benchmark limitations')
  }
  validateSameCorpusProof(scorecard.sameCorpusProof)
}

export function validateSameCorpusScreenshotArtifacts(
  proof: UiResponsivenessSameCorpusProof,
  options: SameCorpusScreenshotArtifactValidationOptions = {},
): void {
  if (!proof.captured) {
    return
  }
  const validationRootDir = options.rootDir ?? rootDir
  const artifactPaths = uniqueScreenshotArtifactPaths(proof)
  if (artifactPaths.length === 0) {
    throw new Error('UI responsiveness same-corpus proof is missing screenshot artifact paths')
  }

  const repoRelativePaths = artifactPaths.map((artifactPath) => validateScreenshotArtifactPath(validationRootDir, artifactPath))
  validateSameCorpusMutationTargetScreenshotArtifactHashes(validationRootDir, proof)
  if (options.requireGitTracked !== true) {
    return
  }

  const trackedPaths = new Set(options.trackedArtifactPaths ?? gitTrackedPaths(validationRootDir, repoRelativePaths))
  for (const artifactPath of repoRelativePaths) {
    if (!trackedPaths.has(artifactPath)) {
      throw new Error(`UI responsiveness same-corpus screenshot artifact is not tracked by git: ${artifactPath}`)
    }
  }
}

export function validateSameCorpusCaptureArtifactMatchesScorecard(
  scorecard: UiResponsivenessLiveBrowserScorecard,
  options: SameCorpusCaptureArtifactValidationOptions = {},
): void {
  if (!scorecard.sameCorpusProof.captured) {
    return
  }
  const validationRootDir = options.rootDir ?? rootDir
  const capturePath = options.capturePath ?? defaultSameCorpusCapturePath
  const absolutePath = resolve(validationRootDir, capturePath)
  const repoRelativePath = relative(validationRootDir, absolutePath)
  if (repoRelativePath.length === 0 || repoRelativePath.startsWith('..') || repoRelativePath.startsWith('/')) {
    throw new Error(`UI responsiveness same-corpus capture artifact escapes the repository: ${capturePath}`)
  }
  if (!existsSync(absolutePath)) {
    throw new Error(`UI responsiveness same-corpus capture artifact is missing: ${repoRelativePath}`)
  }
  const capture = parseSameCorpusCapture(readJsonObject(absolutePath))
  const expectedProof = buildSameCorpusProof(capture)
  if (stableJsonString(scorecard.sameCorpusProof) !== stableJsonString(expectedProof)) {
    throw new Error(`UI responsiveness same-corpus scorecard proof does not match capture artifact: ${repoRelativePath}`)
  }
}

function validateSameCorpusMutationTargetScreenshotArtifactHashes(validationRootDir: string, proof: UiResponsivenessSameCorpusProof): void {
  const expectedHashesByPath = new Map<string, string>()
  for (const entry of proof.cases) {
    for (const productProof of entry.scenarioProof.semanticUiProof.products) {
      for (const mutationProof of productProof.mutationTargetProofs) {
        for (const artifact of sameCorpusMutationTargetScreenshotArtifacts(mutationProof)) {
          const repoRelativePath = validateScreenshotArtifactPath(validationRootDir, artifact.screenshotPath)
          const expectedHash = artifact.screenshotSha256.trim().toLowerCase()
          if (!/^[a-f0-9]{64}$/u.test(expectedHash)) {
            throw new Error(`UI responsiveness same-corpus mutation screenshot artifact is missing SHA256: ${repoRelativePath}`)
          }
          const previousHash = expectedHashesByPath.get(repoRelativePath)
          if (previousHash && previousHash !== expectedHash) {
            throw new Error(`UI responsiveness same-corpus mutation screenshot artifact has conflicting SHA256 values: ${repoRelativePath}`)
          }
          expectedHashesByPath.set(repoRelativePath, expectedHash)
          const actualHash = createHash('sha256')
            .update(readFileSync(resolve(validationRootDir, repoRelativePath)))
            .digest('hex')
          if (actualHash !== expectedHash) {
            throw new Error(`UI responsiveness same-corpus mutation screenshot artifact SHA256 mismatch: ${repoRelativePath}`)
          }
        }
      }
    }
  }
}

function sameCorpusMutationTargetScreenshotArtifacts(mutationProof: {
  readonly screenshotPath: string | null
  readonly screenshotSha256: string | null
  readonly targetScreenshots?: {
    readonly before: { readonly screenshotPath: string | null; readonly screenshotSha256: string | null }
    readonly after: { readonly screenshotPath: string | null; readonly screenshotSha256: string | null }
    readonly restored: { readonly screenshotPath: string | null; readonly screenshotSha256: string | null }
  } | null
}): readonly { readonly screenshotPath: string; readonly screenshotSha256: string }[] {
  const artifacts = [
    mutationProof.targetScreenshots?.before,
    mutationProof.targetScreenshots?.after,
    mutationProof.targetScreenshots?.restored,
  ]
    .filter((entry): entry is { readonly screenshotPath: string | null; readonly screenshotSha256: string | null } => Boolean(entry))
    .flatMap((entry) =>
      entry.screenshotPath && entry.screenshotSha256
        ? [{ screenshotPath: entry.screenshotPath, screenshotSha256: entry.screenshotSha256 }]
        : [],
    )
  if (artifacts.length > 0) {
    return artifacts
  }
  return mutationProof.screenshotPath && mutationProof.screenshotSha256
    ? [{ screenshotPath: mutationProof.screenshotPath, screenshotSha256: mutationProof.screenshotSha256 }]
    : []
}

function uniqueScreenshotArtifactPaths(proof: UiResponsivenessSameCorpusProof): string[] {
  return [
    ...new Set(
      proof.cases.flatMap((entry) => [
        ...(entry.scenarioProof.screenshotProof.captured ? [...entry.scenarioProof.screenshotProof.artifactPaths] : []),
        ...entry.scenarioProof.semanticUiProof.products.flatMap((productProof) =>
          productProof.mutationTargetProofs.flatMap((mutationProof) =>
            sameCorpusMutationTargetScreenshotArtifacts(mutationProof).map((artifact) => artifact.screenshotPath),
          ),
        ),
      ]),
    ),
  ].toSorted()
}

function stableJsonString(value: unknown): string {
  return JSON.stringify(stableJsonValue(value))
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, stableJsonValue(entry)]),
    )
  }
  return value
}

function validateScreenshotArtifactPath(validationRootDir: string, artifactPath: string): string {
  if (artifactPath.length === 0) {
    throw new Error('UI responsiveness same-corpus screenshot artifact path is empty')
  }
  const absolutePath = resolve(validationRootDir, artifactPath)
  const repoRelativePath = relative(validationRootDir, absolutePath)
  if (repoRelativePath.length === 0 || repoRelativePath.startsWith('..') || repoRelativePath.startsWith('/')) {
    throw new Error(`UI responsiveness same-corpus screenshot artifact escapes the repository: ${artifactPath}`)
  }
  if (!existsSync(absolutePath)) {
    throw new Error(`UI responsiveness same-corpus screenshot artifact is missing: ${repoRelativePath}`)
  }
  const stats = statSync(absolutePath)
  if (!stats.isFile() || stats.size <= 0) {
    throw new Error(`UI responsiveness same-corpus screenshot artifact is empty or not a file: ${repoRelativePath}`)
  }
  return repoRelativePath
}

function gitTrackedPaths(validationRootDir: string, repoRelativePaths: readonly string[]): readonly string[] {
  if (repoRelativePaths.length === 0) {
    return []
  }
  const output = execFileSync('git', ['-C', validationRootDir, 'ls-files', '--', ...repoRelativePaths], {
    encoding: 'utf8',
  })
  return output
    .split(/\r?\n/u)
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
}

async function measureBrowserCase(page: Page, spec: BrowserCaseSpec): Promise<BrowserCaseSample> {
  const startedAt = performance.now()
  await page.goto(spec.sourceUrl, { waitUntil: 'domcontentloaded', timeout: 45_000 })
  await waitForCaseReady(page, spec)
  const loadToReadyMs = performance.now() - startedAt
  const title = await page.title()
  const accessMode = await detectAccessMode(page, spec.vendor)

  await page.mouse.move(viewport.width / 2, viewport.height / 2)
  const scrollStartedAt = performance.now()
  await page.mouse.wheel(0, 720)
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
  const scrollResponseMs = performance.now() - scrollStartedAt
  return {
    finalUrl: page.url(),
    title,
    accessMode,
    loadToReadyMs,
    scrollResponseMs,
    postScrollFrameMs: summarizeNumbers(frameIntervals).p95,
  }
}

async function waitForCaseReady(page: Page, spec: BrowserCaseSpec): Promise<void> {
  if (spec.vendor === 'google-sheets') {
    await page.waitForFunction(
      () =>
        !window.location.href.includes('accounts.google.com') &&
        document.title.includes('Google Sheets') &&
        (document.body.innerText.includes('Comment only') || document.body.innerText.includes('View only')),
      { timeout: 45_000 },
    )
    await page.waitForTimeout(2_000)
    return
  }

  await page.waitForFunction(() => document.title.endsWith('.xlsx'), { timeout: 45_000 })
  await page.waitForTimeout(5_000)
}

async function detectAccessMode(
  page: Page,
  vendor: UiResponsivenessLiveBrowserVendor,
): Promise<UiResponsivenessLiveBrowserCase['accessMode']> {
  if (vendor === 'microsoft-excel-web') {
    return 'public-office-web-viewer'
  }
  const bodyText = await page.locator('body').innerText({ timeout: 5_000 })
  return bodyText.includes('Comment only') ? 'public-comment-only' : 'public-view-only'
}

function buildBrowserCase(spec: BrowserCaseSpec, samples: readonly BrowserCaseSample[]): UiResponsivenessLiveBrowserCase {
  const title = samples[0]?.title ?? ''
  const finalUrl = samples[0]?.finalUrl ?? ''
  const accessMode = samples[0]?.accessMode ?? 'public-view-only'
  const loadToReadyMs = summarizeNumbers(samples.map((entry) => entry.loadToReadyMs))
  const scrollResponseMs = summarizeNumbers(samples.map((entry) => entry.scrollResponseMs))
  const postScrollFrameMs = summarizeNumbers(samples.map((entry) => entry.postScrollFrameMs))
  return {
    id: spec.id,
    vendor: spec.vendor,
    product: spec.product,
    sourceUrl: spec.sourceUrl,
    finalUrl,
    title,
    accessMode,
    workload: 'open-public-workbook-and-scroll-viewport',
    sampleCount: samples.length,
    loadToReadyMs,
    scrollResponseMs,
    postScrollFrameMs,
    passed:
      samples.length === sampleCount &&
      title.includes(spec.expectedTitleIncludes) &&
      samples.every(
        (entry) => Number.isFinite(entry.loadToReadyMs) && Number.isFinite(entry.scrollResponseMs) && entry.postScrollFrameMs > 0,
      ),
    limitations: [
      'Public browser timing cannot exercise authenticated workbook editing or tenant-local collaboration paths.',
      'This timing is direct incumbent browser evidence, but it is not a same-corpus 10x proof by itself.',
    ],
  }
}

function validateSummary(summary: NumericSummary, label: string): void {
  if (summary.samples.length < sampleCount) {
    throw new Error(`UI responsiveness live browser scorecard has too few samples for ${label}`)
  }
  for (const value of [summary.min, summary.median, summary.p95, summary.max, summary.mean, ...summary.samples]) {
    if (!Number.isFinite(value) || value < 0) {
      throw new Error(`UI responsiveness live browser scorecard has invalid numeric summary for ${label}`)
    }
  }
}

function argumentValue(argv: readonly string[], name: string): string | null {
  const index = argv.indexOf(name)
  if (index === -1) {
    return null
  }
  const value = argv[index + 1]
  if (value === undefined || value.trim().length === 0 || value.startsWith('-')) {
    throw new Error(`Missing value after ${name}`)
  }
  return value
}

function logResult(mode: 'check' | 'write', scorecard: UiResponsivenessLiveBrowserScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        capturedVendors: scorecard.summary.capturedVendors,
        caseCount: scorecard.cases.length,
        sameCorpusProofCaptured: scorecard.sameCorpusProof.captured,
        sameCorpusTenXMeanAndP95CaseCount: scorecard.sameCorpusProof.tenXMeanAndP95CaseCount,
      },
      null,
      2,
    ),
  )
}

if (import.meta.url === `file://${process.argv[1]}`) {
  await main()
}
