import { mkdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import { chromium, type Browser, type BrowserContextOptions, type Page } from '@playwright/test'

import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusCase } from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCapture,
  SameCorpusBiligRuntimeProof,
  SameCorpusBiligRuntimeProofSample,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { CaptureArgs, PreflightArgs, SaveStorageStateArgs } from './ui-responsiveness-same-corpus-args.ts'
import { defaultViewport } from './ui-responsiveness-same-corpus-args.ts'
import { buildSameCorpusCaptureRunManifest, sameCorpusScenarioCaseFields } from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadMutatesWorkbook,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusMutatingWorkload,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'
import { verifyProductCorpus, waitForVerifiedBiligRenderedSurface } from './ui-responsiveness-same-corpus-verification.ts'
import {
  buildCaptureScenarioProof,
  captureSameCorpusProductVisualProof,
  type SameCorpusMutationTargetProof,
  type SameCorpusMutationTargetReadback,
  type SameCorpusProductVisualProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import {
  captureSameCorpusMutationTargetScreenshotProof,
  readSameCorpusMutationTargetReadback,
  readSameCorpusMutationTargetRevisionProof,
  readSameCorpusMutationTargetSelection,
  readSameCorpusVisibleMutationTargetReadback,
  selectSameCorpusMutationTargetRange,
  type SameCorpusMutationTargetSelection,
} from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import { productLimitations, sameCorpusChromiumLaunchOptions, settleFrames } from './ui-responsiveness-same-corpus-page-utils.ts'
import {
  measureVisibleScrollResponseWithRetries,
  movePointerToProductViewport,
  resetProductScrollPosition,
} from './ui-responsiveness-same-corpus-scroll-page.ts'
import { measureVisibleNonScrollResponse } from './ui-responsiveness-same-corpus-visible-response-page.ts'
import {
  incumbentEditableWorkloadBlocker,
  measureProductWorkload,
  restoreProductWorkbookMutation,
  sameCorpusFillColorExpectedColor,
  sameCorpusFillColorSwatchLabel,
  type ProductOperationSample,
} from './ui-responsiveness-same-corpus-workload-runner.ts'
import type { PreflightProductResult, SameCorpusPreflight } from './ui-responsiveness-same-corpus-preflight.ts'

interface ProductSampleCollection {
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly biligRuntimeProof?: SameCorpusBiligRuntimeProof
  readonly mutationTargetProofs: readonly SameCorpusMutationTargetProof[]
  readonly samples: readonly ProductOperationSample[]
}

interface SameCorpusProductMeasurementUrls {
  readonly biligUrl: string
  readonly googleSheetsUrl: string
  readonly microsoftExcelWebUrl: string | null
}

interface SameCorpusProductMeasurements {
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb?: SameCorpusCaptureMeasurement | undefined
}

type SameCorpusProductMeasure = (
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  workload: UiResponsivenessSameCorpusWorkload,
) => Promise<SameCorpusCaptureMeasurement>

export async function saveStorageState(args: SaveStorageStateArgs): Promise<void> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch(sameCorpusChromiumLaunchOptions(args.headless))
  const context = await browser.newContext({ viewport: defaultViewport })
  const page = await context.newPage()
  try {
    await page.goto(args.authUrl, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, args.product, captureArgsForStorageState(args))
    const corpusVerification = await verifyProductCorpus(page, args.product, args.authUrl, corpus)
    mkdirSync(dirname(args.targetPath), { recursive: true })
    await context.storageState({ path: args.targetPath })
    console.log(
      JSON.stringify(
        {
          mode: 'save-storage-state',
          product: args.product,
          corpusCaseId: corpus.id,
          targetPath: args.targetPath,
          finalUrl: page.url(),
          title: await page.title(),
          corpusVerification,
        },
        null,
        2,
      ),
    )
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, args.product, args.authUrl, 0, error), { cause: error })
  } finally {
    await context.close()
    await browser.close()
  }
}

export async function captureSameCorpusUiResponsiveness(args: CaptureArgs): Promise<SameCorpusCapture> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch(sameCorpusChromiumLaunchOptions(args.headless))
  try {
    const cases = await captureSameCorpusWorkloadCases(browser, corpus, args)
    return buildSameCorpusCaptureArtifact({
      sampleCount: args.sampleCount,
      limitations: defaultSameCorpusCaptureLimitations(),
      cases,
    })
  } finally {
    await browser.close()
  }
}

export function buildSameCorpusCaptureArtifact(args: {
  readonly sampleCount: number
  readonly limitations?: readonly string[]
  readonly cases: readonly SameCorpusCaptureCase[]
}): SameCorpusCapture {
  const cases = args.cases.map((entry) => ({
    ...entry,
    ...sameCorpusScenarioCaseFields(entry.scenarioProof),
  }))
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-capture',
    sampleCount: args.sampleCount,
    runManifest: buildSameCorpusCaptureRunManifest(cases, args.sampleCount),
    limitations: [...(args.limitations ?? defaultSameCorpusCaptureLimitations())],
    cases,
  }
}

function defaultSameCorpusCaptureLimitations(): readonly string[] {
  return [
    'Caller must supply a Google Sheets URL for the same exported Bilig benchmark corpus.',
    'Microsoft Excel Web can be supplied as an additional incumbent comparison, but it is not required for the Google Sheets 10x claim.',
    'Edit and format workloads require the supplied incumbent URLs to allow browser-driven editing in the authenticated context.',
    'Bilig non-scroll operationResponseMs measures interaction-visible response first; strict rendered-grid proof is still required after the authoritative frame catches up.',
    'Bilig captures record authoritativeRenderProofMsSamples separately so interaction-visible timing cannot hide slow or missing rendered-proof completion.',
  ]
}

async function captureSameCorpusWorkloadCases(
  browser: Browser,
  corpus: WorkbookBenchmarkCorpusCase,
  args: CaptureArgs,
  workloadIndex = 0,
  cases: SameCorpusCaptureCase[] = [],
): Promise<SameCorpusCaptureCase[]> {
  const workload = requiredUiResponsivenessSameCorpusWorkloads[workloadIndex]
  if (!workload) {
    return cases
  }
  const caseId = `same-corpus-${args.corpusId}-${workload}`
  const visualProofs: SameCorpusProductVisualProof[] = []
  const { bilig, googleSheets, microsoftExcelWeb } = await collectSameCorpusProductMeasurements(
    args,
    (product, url, measuredWorkload) => measureProduct(browser, product, url, corpus, args, measuredWorkload, caseId, visualProofs),
    workload,
  )
  const scenarioProof = buildCaptureScenarioProof({ bilig, googleSheets, microsoftExcelWeb, visualProofs, workload })
  if (!scenarioProof.screenshotProof.captured || !scenarioProof.pixelGridProof.captured) {
    throw new Error(`same-corpus UI capture is missing browser-visible proof for ${caseId}: ${JSON.stringify(scenarioProof)}`)
  }
  cases.push({
    id: caseId,
    corpusCaseId: args.corpusId,
    materializedCells: corpus.materializedCellCount,
    workload,
    ...sameCorpusScenarioCaseFields(scenarioProof),
    scenarioProof,
    bilig,
    googleSheets,
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
  })
  return await captureSameCorpusWorkloadCases(browser, corpus, args, workloadIndex + 1, cases)
}

export async function collectSameCorpusProductMeasurements(
  urls: SameCorpusProductMeasurementUrls,
  measure: SameCorpusProductMeasure,
  workload: UiResponsivenessSameCorpusWorkload = 'scroll-vertical',
): Promise<SameCorpusProductMeasurements> {
  const bilig = await measure('bilig', urls.biligUrl, workload)
  assertSameCorpusProductMeasurement('bilig', urls.biligUrl, bilig, workload)
  const googleSheets = await measure('google-sheets', urls.googleSheetsUrl, workload)
  assertSameCorpusProductMeasurement('google-sheets', urls.googleSheetsUrl, googleSheets, workload)
  if (!urls.microsoftExcelWebUrl) {
    return { bilig, googleSheets }
  }
  const microsoftExcelWeb = await measure('microsoft-excel-web', urls.microsoftExcelWebUrl, workload)
  assertSameCorpusProductMeasurement('microsoft-excel-web', urls.microsoftExcelWebUrl, microsoftExcelWeb, workload)
  return { bilig, googleSheets, microsoftExcelWeb }
}

function assertSameCorpusProductMeasurement(
  product: UiResponsivenessSameCorpusProduct,
  source: string,
  measurement: SameCorpusCaptureMeasurement,
  workload: UiResponsivenessSameCorpusWorkload,
): void {
  if (measurement.product !== product) {
    throw new Error(`same-corpus UI measurement expected ${product} but received ${measurement.product}`)
  }
  if (measurement.source !== source) {
    throw new Error(`same-corpus UI measurement for ${product} used an unexpected source URL`)
  }
  assertSameCorpusSampleArray(product, 'operation response', measurement.operationResponseMsSamples)
  if (product === 'bilig') {
    assertSameCorpusSampleArray(
      product,
      'authoritative render proof',
      measurement.authoritativeRenderProofMsSamples,
      measurement.operationResponseMsSamples.length,
    )
  }
  assertSameCorpusSampleArray(
    product,
    'post-operation frame',
    measurement.postOperationFrameMsSamples,
    measurement.operationResponseMsSamples.length,
  )
  if (uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
    assertSameCorpusSampleArray(
      product,
      'scroll-event response',
      measurement.scrollEventResponseMsSamples,
      measurement.operationResponseMsSamples.length,
    )
    assertSameCorpusSampleArray(
      product,
      'scroll movement',
      measurement.scrollMovementPxSamples,
      measurement.operationResponseMsSamples.length,
    )
  }
}

function assertSameCorpusSampleArray(
  product: UiResponsivenessSameCorpusProduct,
  label: string,
  samples: readonly number[] | undefined,
  expectedLength?: number,
): void {
  if (!samples || samples.length === 0) {
    throw new Error(`same-corpus UI measurement for ${product} is missing ${label} samples`)
  }
  if (expectedLength !== undefined && samples.length !== expectedLength) {
    throw new Error(
      `same-corpus UI measurement for ${product} has ${String(samples.length)} ${label} samples but expected ${String(expectedLength)}`,
    )
  }
  for (const sample of samples) {
    if (!Number.isFinite(sample)) {
      throw new Error(`same-corpus UI measurement for ${product} has a non-finite ${label} sample`)
    }
  }
}

export async function preflightSameCorpusIncumbentAccess(args: PreflightArgs): Promise<SameCorpusPreflight> {
  const corpus = buildWorkbookBenchmarkCorpus(args.corpusId)
  const browser = await chromium.launch(sameCorpusChromiumLaunchOptions(args.headless))
  try {
    const productSpecs = [
      ...(args.googleSheetsUrl ? [{ product: 'google-sheets' as const, url: args.googleSheetsUrl }] : []),
      ...(args.microsoftExcelWebUrl ? [{ product: 'microsoft-excel-web' as const, url: args.microsoftExcelWebUrl }] : []),
    ]
    const products = await Promise.all(productSpecs.map((spec) => preflightIncumbentProduct(browser, spec.product, spec.url, corpus, args)))
    return {
      mode: 'preflight',
      corpusCaseId: corpus.id,
      materializedCells: corpus.materializedCellCount,
      requiredProductCount: 2,
      checkedProductCount: products.length,
      readyProductCount: products.filter((product) => product.status === 'ready').length,
      blockedProductCount: products.filter((product) => product.status === 'blocked').length,
      allCheckedProductsReady: products.length > 0 && products.every((product) => product.status === 'ready'),
      products,
    }
  } finally {
    await browser.close()
  }
}

function captureArgsForStorageState(args: SaveStorageStateArgs): CaptureArgs {
  return {
    biligUrl: args.authUrl,
    allowIncompleteEvidence: true,
    biligStorageStatePath: null,
    corpusId: args.corpusId,
    deltaX: 0,
    deltaY: 720,
    googleSheetsUrl: args.authUrl,
    googleSheetsStorageStatePath: null,
    headless: args.headless,
    microsoftExcelWebUrl: args.authUrl,
    microsoftExcelWebStorageStatePath: null,
    outputPath: args.targetPath,
    readyTimeoutMs: args.readyTimeoutMs,
    sampleCount: 1,
    storageStatePath: null,
  }
}

async function preflightIncumbentProduct(
  browser: Browser,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  url: string,
  corpus: WorkbookBenchmarkCorpusCase,
  args: PreflightArgs,
): Promise<PreflightProductResult> {
  const context = await browser.newContext(browserContextOptionsForPreflightProduct(product, args))
  const page = await context.newPage()
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, product, captureArgsForPreflight(args, product, url))
    await assertIncumbentEditableForPreflight(page, product)
    const corpusVerification = await verifyProductCorpus(page, product, url, corpus)
    return {
      product,
      source: url,
      finalUrl: page.url(),
      title: await page.title(),
      status: 'ready',
      blocker: null,
      corpusVerification,
      limitations: productLimitations(product, storageStatePathForPreflightProduct(product, args)),
    }
  } catch (error: unknown) {
    const diagnostic = await collectPageDiagnostic(page)
    return {
      product,
      source: url,
      finalUrl: diagnostic.finalUrl,
      title: diagnostic.title,
      status: 'blocked',
      blocker: await productReadyFailureMessage(page, product, url, 0, error),
      corpusVerification: null,
      limitations: productLimitations(product, storageStatePathForPreflightProduct(product, args)),
    }
  } finally {
    await context.close()
  }
}

async function assertIncumbentEditableForPreflight(
  page: Page,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
): Promise<void> {
  const bodyText = await page
    .locator('body')
    .innerText({ timeout: 2_000 })
    .catch(() => '')
  const blocker = incumbentEditableWorkloadBlocker(product, page.url(), bodyText)
  if (blocker) {
    throw new Error(`Cannot preflight same-corpus editable workloads on ${product}: ${blocker}`)
  }
}

function browserContextOptionsForPreflightProduct(
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  args: PreflightArgs,
): BrowserContextOptions {
  const storageState = storageStatePathForPreflightProduct(product, args)
  return {
    viewport: defaultViewport,
    ...(storageState ? { storageState } : {}),
  }
}

function storageStatePathForPreflightProduct(
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  args: PreflightArgs,
): string | null {
  if (product === 'google-sheets') {
    return args.googleSheetsStorageStatePath ?? args.storageStatePath
  }
  return args.microsoftExcelWebStorageStatePath ?? args.storageStatePath
}

function captureArgsForPreflight(
  args: PreflightArgs,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  url: string,
): CaptureArgs {
  return {
    biligUrl: url,
    allowIncompleteEvidence: true,
    biligStorageStatePath: null,
    corpusId: args.corpusId,
    deltaX: 0,
    deltaY: 720,
    googleSheetsUrl: product === 'google-sheets' ? url : '',
    googleSheetsStorageStatePath: args.googleSheetsStorageStatePath,
    headless: args.headless,
    microsoftExcelWebUrl: product === 'microsoft-excel-web' ? url : '',
    microsoftExcelWebStorageStatePath: args.microsoftExcelWebStorageStatePath,
    outputPath: args.outputPath ?? '',
    readyTimeoutMs: args.readyTimeoutMs,
    sampleCount: 1,
    storageStatePath: args.storageStatePath,
  }
}

async function measureProduct(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  corpus: WorkbookBenchmarkCorpusCase,
  args: CaptureArgs,
  workload: UiResponsivenessSameCorpusWorkload,
  caseId?: string,
  visualProofs?: SameCorpusProductVisualProof[],
): Promise<SameCorpusCaptureMeasurement> {
  const collection = await measureProductSamples(browser, product, url, corpus, args, workload, caseId, visualProofs)
  if (caseId && visualProofs && uiSameCorpusWorkloadMutatesWorkbook(workload)) {
    visualProofs.push(
      await captureSameCorpusProductVisualProofFromFreshPage({
        args,
        browser,
        caseId,
        corpusVerification: collection.corpusVerification,
        mutationTargetProofs: collection.mutationTargetProofs,
        product,
        url,
        workload,
      }),
    )
  }

  return {
    product,
    source: url,
    operationResponseMsSamples: collection.samples.map((entry) => entry.operationResponseMs),
    operationResponseProofs: collection.samples.map((entry) => entry.operationResponseProof),
    ...(product === 'bilig'
      ? { authoritativeRenderProofMsSamples: collection.samples.map((entry) => entry.authoritativeRenderProofMs ?? Number.NaN) }
      : {}),
    ...(uiSameCorpusWorkloadMutatesWorkbook(workload)
      ? { committedTargetProofMsSamples: collection.samples.map((entry) => entry.committedTargetProofMs ?? Number.NaN) }
      : {}),
    postOperationFrameMsSamples: collection.samples.map((entry) => entry.postOperationFrameMs),
    ...(uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
      ? {
          scrollEventResponseMsSamples: collection.samples.map((entry) => entry.scrollEventResponseMs ?? Number.NaN),
          scrollMovementPxSamples: collection.samples.map((entry) => entry.scrollMovementPx ?? Number.NaN),
        }
      : {}),
    ...(collection.biligRuntimeProof ? { biligRuntimeProof: collection.biligRuntimeProof } : {}),
    corpusVerification: collection.corpusVerification,
    limitations: productLimitations(product, storageStatePathForProduct(product, args)),
  }
}

async function measureProductSamples(
  browser: Browser,
  product: UiResponsivenessSameCorpusProduct,
  url: string,
  corpus: WorkbookBenchmarkCorpusCase,
  args: CaptureArgs,
  workload: UiResponsivenessSameCorpusWorkload,
  caseId: string | undefined = undefined,
  visualProofs: SameCorpusProductVisualProof[] | undefined = undefined,
  sampleIndex = 0,
  samples: ProductOperationSample[] = [],
  corpusVerification: SameCorpusCaptureCorpusVerification | null = null,
  runtimeProofSamples: SameCorpusBiligRuntimeProofSample[] = [],
  mutationTargetProofs: SameCorpusMutationTargetProof[] = [],
): Promise<ProductSampleCollection> {
  if (sampleIndex >= args.sampleCount) {
    if (!corpusVerification) {
      throw new Error(`Missing same-corpus fingerprint verification for ${product}`)
    }
    return {
      corpusVerification,
      ...(product === 'bilig' ? { biligRuntimeProof: buildBiligRuntimeProof(url, runtimeProofSamples) } : {}),
      mutationTargetProofs,
      samples,
    }
  }
  const context = await browser.newContext(browserContextOptionsForProduct(product, args))
  const page = await context.newPage()
  let nextCorpusVerification = corpusVerification
  try {
    const loadStartedAt = performance.now()
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, product, args)
    if (product === 'bilig') {
      runtimeProofSamples.push(await readBiligRuntimeProofSample(page, sampleIndex))
    }
    const loadToReadyMs = performance.now() - loadStartedAt
    nextCorpusVerification ??= await verifyProductCorpus(page, product, url, corpus)
    if (product !== 'microsoft-excel-web' && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
      await resetProductScrollPosition(page, product)
      await settleFrames(page, 3)
    }
    const mutatingWorkload: UiResponsivenessSameCorpusMutatingWorkload | null = uiSameCorpusWorkloadMutatesWorkbook(workload)
      ? workload
      : null
    const mutationTarget = mutatingWorkload
      ? await readSameCorpusMutationTargetSelection({ page, product, sheetName: nextCorpusVerification.sheetName })
      : null
    const mutationTargetBefore =
      mutationTarget === null ? null : await readSameCorpusMutationTargetReadback({ page, product, target: mutationTarget })
    const operationStartedAt = workload === 'open-workbook' ? loadStartedAt : performance.now()
    const sample = await measureProductWorkload({
      page,
      product,
      captureArgs: args,
      workload,
      sampleIndex,
      loadToReadyMs,
      hooks: {
        measureVisibleScrollResponseWithRetries,
        measureVisibleNonScrollResponse,
        movePointerToProductViewport,
      },
    })
    const sampleWithRenderProof = await withAuthoritativeRenderProofTiming(page, product, sample, operationStartedAt, args.readyTimeoutMs)
    if (mutationTarget && mutationTargetBefore && mutatingWorkload) {
      const mutationTargetProof = await captureSameCorpusMutationTargetProofForSample({
        before: mutationTargetBefore,
        caseId,
        operationStartedAt,
        outputPath: args.outputPath,
        page,
        product,
        sampleIndex,
        target: mutationTarget,
        workload: mutatingWorkload,
      })
      mutationTargetProofs.push(mutationTargetProof)
      samples.push({ ...sampleWithRenderProof, committedTargetProofMs: mutationTargetProof.committedTargetProofMs })
    } else {
      samples.push(sampleWithRenderProof)
      await restoreProductWorkbookMutation(page, workload)
    }
    if (caseId && visualProofs && sampleIndex === 0 && !uiSameCorpusWorkloadMutatesWorkbook(workload)) {
      visualProofs.push(
        await captureSameCorpusProductVisualProof({
          caseId,
          corpusVerification: nextCorpusVerification,
          outputPath: args.outputPath,
          page,
          product,
          sampleIndex,
          sampleCount: args.sampleCount,
          workload,
        }),
      )
    }
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, product, url, sampleIndex, error), { cause: error })
  } finally {
    await context.close()
  }
  return measureProductSamples(
    browser,
    product,
    url,
    corpus,
    args,
    workload,
    caseId,
    visualProofs,
    sampleIndex + 1,
    samples,
    nextCorpusVerification,
    runtimeProofSamples,
    mutationTargetProofs,
  )
}

async function captureSameCorpusProductVisualProofFromFreshPage(input: {
  readonly args: CaptureArgs
  readonly browser: Browser
  readonly caseId: string
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly mutationTargetProofs: readonly SameCorpusMutationTargetProof[]
  readonly product: UiResponsivenessSameCorpusProduct
  readonly url: string
  readonly workload: UiResponsivenessSameCorpusWorkload
}): Promise<SameCorpusProductVisualProof> {
  const context = await input.browser.newContext(browserContextOptionsForProduct(input.product, input.args))
  const page = await context.newPage()
  try {
    await page.goto(input.url, { waitUntil: 'domcontentloaded', timeout: 60_000 })
    await waitForProductReady(page, input.product, input.args)
    const target = input.mutationTargetProofs[0] ? mutationTargetSelectionFromProof(input.mutationTargetProofs[0]) : null
    if (target) {
      await selectSameCorpusMutationTargetRange({ page, product: input.product, target })
    }
    return await captureSameCorpusProductVisualProof({
      caseId: input.caseId,
      corpusVerification: input.corpusVerification,
      mutationTargetProofs: input.mutationTargetProofs,
      outputPath: input.args.outputPath,
      page,
      product: input.product,
      sampleCount: input.args.sampleCount,
      sampleIndex: 0,
      workload: input.workload,
    })
  } catch (error: unknown) {
    throw new Error(await productReadyFailureMessage(page, input.product, input.url, 0, error), { cause: error })
  } finally {
    await context.close()
  }
}

function mutationTargetSelectionFromProof(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetSelection {
  const [startAddress, endAddress = startAddress] = proof.targetRange.split(':')
  return {
    endAddress,
    sheetName: proof.sheetName,
    sheetId: proof.sheetId,
    startAddress,
    targetRange: proof.targetRange,
  }
}

async function captureSameCorpusMutationTargetProofForSample(args: {
  readonly before: SameCorpusMutationTargetReadback
  readonly caseId?: string
  readonly operationStartedAt: number
  readonly outputPath: string
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetProof> {
  await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target })
  const after = await readSameCorpusMutationTargetReadback({ page: args.page, product: args.product, target: args.target })
  const visibleAfter = await readSameCorpusVisibleMutationTargetReadback({
    page: args.page,
    product: args.product,
    target: args.target,
    workload: args.workload,
  })
  const screenshotPath = mutationTargetScreenshotArtifactPath({
    caseId: args.caseId,
    outputPath: args.outputPath,
    product: args.product,
    sampleIndex: args.sampleIndex,
    workload: args.workload,
  })
  mkdirSync(dirname(screenshotPath), { recursive: true })
  const screenshotProof = await captureSameCorpusMutationTargetScreenshotProof(
    args.page,
    args.product,
    screenshotPath,
    repoRelativePath(screenshotPath),
  )
  const revisions = await readSameCorpusMutationTargetRevisionProof({
    page: args.page,
    product: args.product,
    readback: after,
    screenshotSha256: screenshotProof.screenshotSha256,
    target: args.target,
  })
  const committedTargetProofMs = Math.max(0, performance.now() - args.operationStartedAt)
  await restoreProductWorkbookMutation(args.page, args.workload)
  await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target })
  const restored = await readSameCorpusMutationTargetReadback({ page: args.page, product: args.product, target: args.target })
  const visibleRestored = await readSameCorpusVisibleMutationTargetReadback({
    page: args.page,
    product: args.product,
    target: args.target,
    workload: args.workload,
  })
  return {
    after,
    authoritativeReadbackRevision: revisions.authoritativeReadbackRevision,
    before: args.before,
    committedTargetProofMs,
    intendedOperation: args.workload,
    intendedPayload: intendedMutationTargetPayload(args.product, args.workload, args.sampleIndex),
    restored,
    sampleIndex: args.sampleIndex,
    screenshotPath: screenshotProof.screenshotPath,
    screenshotSha256: screenshotProof.screenshotSha256,
    sheetId: args.target.sheetId,
    sheetName: args.target.sheetName,
    targetRange: args.target.targetRange,
    undoRestoreStatus: sameCorpusMutationReadbacksEqual(args.before, restored) ? 'verified' : 'failed',
    visibleAfter,
    visibleRenderRevision: revisions.visibleRenderRevision,
    visibleRestored,
    workload: args.workload,
  }
}

function intendedMutationTargetPayload(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['intendedPayload'] {
  if (workload === 'formula-edit') {
    return { kind: 'formula', formula: `=${String(sampleIndex + 1)}+1` }
  }
  if (workload === 'fill-format-change') {
    return {
      kind: 'fill-color',
      expectedFillColor: sameCorpusFillColorExpectedColor(sampleIndex),
      swatchLabel: sameCorpusFillColorSwatchLabel(sampleIndex),
    }
  }
  return { kind: 'cell-value', value: `${product}-same-corpus-${String(sampleIndex + 1)}` }
}

function mutationTargetScreenshotArtifactPath(args: {
  readonly caseId?: string
  readonly outputPath: string
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): string {
  const caseId = args.caseId ?? `same-corpus-${args.workload}`
  return resolve(`${args.outputPath}.proof`, caseId, 'mutation-target', `${args.product}-sample-${String(args.sampleIndex + 1)}-after.png`)
}

function repoRelativePath(path: string): string {
  return relative(process.cwd(), path)
}

function sameCorpusMutationReadbacksEqual(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return (
    left.value === right.value &&
    left.formula === right.formula &&
    left.fillColor === right.fillColor &&
    left.visibleText === right.visibleText &&
    left.source === right.source
  )
}

async function withAuthoritativeRenderProofTiming(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  sample: ProductOperationSample,
  operationStartedAt: number,
  timeoutMs: number,
): Promise<ProductOperationSample> {
  if (product !== 'bilig') {
    return sample
  }
  await waitForVerifiedBiligRenderedSurface(page, timeoutMs)
  return {
    ...sample,
    authoritativeRenderProofMs: Math.max(sample.operationResponseMs, performance.now() - operationStartedAt),
  }
}

function buildBiligRuntimeProof(source: string, samples: readonly SameCorpusBiligRuntimeProofSample[]): SameCorpusBiligRuntimeProof {
  const presentSamples = samples.filter((sample) => sample.present)
  const firstPresent = presentSamples[0]
  const actualBuildKind =
    presentSamples.length === 0
      ? 'unknown'
      : presentSamples.every((sample) => sample.buildKind === 'production')
        ? 'production'
        : (firstPresent?.buildKind ?? 'unknown')
  const prod = presentSamples.length > 0 && presentSamples.every((sample) => sample.prod)
  const dev = presentSamples.some((sample) => sample.dev)
  const remoteSyncEnabled = firstPresent?.remoteSyncEnabled ?? null
  const entryRoute = firstPresent?.entryRoute ?? null
  const verified =
    samples.length > 0 && samples.every((sample) => sample.present && sample.buildKind === 'production' && sample.prod && !sample.dev)
  return {
    product: 'bilig',
    source,
    verificationMethod: 'window.__biligRuntimeBuild',
    requiredBuildKind: 'production',
    actualBuildKind,
    mode: firstPresent?.mode ?? 'unknown',
    dev,
    prod,
    remoteSyncEnabled,
    entryRoute,
    sampleCount: samples.length,
    verified,
    samples: samples.map((sample) => ({ ...sample })),
  }
}

async function readBiligRuntimeProofSample(page: Page, sampleIndex: number): Promise<SameCorpusBiligRuntimeProofSample> {
  const runtimeBuild = await page.evaluate(() => {
    const value = (window as Window & { __biligRuntimeBuild?: unknown }).__biligRuntimeBuild
    if (!value || typeof value !== 'object') {
      return null
    }
    const app = Reflect.get(value, 'app')
    const buildKind = Reflect.get(value, 'buildKind')
    const mode = Reflect.get(value, 'mode')
    const dev = Reflect.get(value, 'dev')
    const prod = Reflect.get(value, 'prod')
    const remoteSyncEnabled = Reflect.get(value, 'remoteSyncEnabled')
    const entryRoute = Reflect.get(value, 'entryRoute')
    return {
      app: typeof app === 'string' ? app : null,
      buildKind: buildKind === 'production' || buildKind === 'development' ? buildKind : 'unknown',
      mode: typeof mode === 'string' ? mode : 'unknown',
      dev: dev === true,
      prod: prod === true,
      remoteSyncEnabled: typeof remoteSyncEnabled === 'boolean' ? remoteSyncEnabled : null,
      entryRoute: typeof entryRoute === 'string' ? entryRoute : null,
    }
  })
  if (!runtimeBuild) {
    return {
      sampleIndex,
      present: false,
      app: null,
      buildKind: 'unknown',
      mode: 'unknown',
      dev: false,
      prod: false,
      remoteSyncEnabled: null,
      entryRoute: null,
    }
  }
  return {
    sampleIndex,
    present: runtimeBuild.app === 'bilig-web',
    app: runtimeBuild.app,
    buildKind: runtimeBuild.buildKind,
    mode: runtimeBuild.mode,
    dev: runtimeBuild.dev,
    prod: runtimeBuild.prod,
    remoteSyncEnabled: runtimeBuild.remoteSyncEnabled,
    entryRoute: runtimeBuild.entryRoute,
  }
}

function browserContextOptionsForProduct(product: UiResponsivenessSameCorpusProduct, args: CaptureArgs): BrowserContextOptions {
  const storageState = storageStatePathForProduct(product, args)
  return {
    viewport: defaultViewport,
    ...(storageState ? { storageState } : {}),
  }
}

function storageStatePathForProduct(product: UiResponsivenessSameCorpusProduct, args: CaptureArgs): string | null {
  if (product === 'bilig') {
    return args.biligStorageStatePath ?? args.storageStatePath
  }
  if (product === 'google-sheets') {
    return args.googleSheetsStorageStatePath ?? args.storageStatePath
  }
  return args.microsoftExcelWebStorageStatePath ?? args.storageStatePath
}

async function waitForProductReady(page: Page, product: UiResponsivenessSameCorpusProduct, args: CaptureArgs): Promise<void> {
  if (product === 'bilig') {
    await page.waitForSelector('[data-testid="sheet-grid"]', { state: 'visible', timeout: args.readyTimeoutMs })
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
      args.corpusId,
      { timeout: args.readyTimeoutMs },
    )
    await settleFrames(page, 180)
    await waitForVerifiedBiligRenderedSurface(page, args.readyTimeoutMs)
    return
  }

  if (product === 'google-sheets') {
    await page.waitForFunction(
      () =>
        !window.location.href.includes('accounts.google.com') &&
        document.title.includes('Google Sheets') &&
        !document.body.innerText.includes('Sign in\nto continue to Google Sheets'),
      { timeout: args.readyTimeoutMs },
    )
    await settleFrames(page, 120)
    return
  }

  await page.waitForFunction(
    () => document.title.toLowerCase().includes('.xlsx') || document.body.innerText.toLowerCase().includes('excel'),
    { timeout: args.readyTimeoutMs },
  )
  await waitForFrameElementReady(page, '.ewr-grdcontarea-grid', args.readyTimeoutMs)
  await settleFrames(page, 180)
}

async function productReadyFailureMessage(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  sourceUrl: string,
  sampleIndex: number,
  cause: unknown,
): Promise<string> {
  const diagnostic = await collectPageDiagnostic(page)
  const causeMessage = cause instanceof Error ? cause.message : String(cause)
  const productHint =
    product === 'google-sheets' && diagnostic.finalUrl.includes('accounts.google.com')
      ? 'Google Sheets redirected to sign-in; provide a public/shareable sheet URL or run with --google-sheets-storage-state from an authenticated Playwright session.'
      : product === 'microsoft-excel-web' && sourceUrl.includes('view.officeapps.live.com')
        ? 'Microsoft Excel Web did not become measurable; confirm the viewer URL wraps a Microsoft-accessible public HTTPS XLSX URL for the same emitted corpus.'
        : 'The same-corpus page did not reach the expected measurable state.'
  return [
    `Failed to prepare ${product} for same-corpus UI capture on sample ${String(sampleIndex + 1)}.`,
    productHint,
    `sourceUrl: ${sourceUrl}`,
    `finalUrl: ${diagnostic.finalUrl}`,
    `title: ${diagnostic.title}`,
    `body: ${diagnostic.bodySnippet}`,
    `cause: ${causeMessage}`,
  ].join('\n')
}

async function collectPageDiagnostic(page: Page): Promise<{ finalUrl: string; title: string; bodySnippet: string }> {
  const [title, bodySnippet] = await Promise.all([
    page.title().catch(() => ''),
    page
      .locator('body')
      .innerText({ timeout: 2_000 })
      .catch(() => ''),
  ])
  return {
    finalUrl: page.url(),
    title,
    bodySnippet: bodySnippet.replace(/\s+/g, ' ').trim().slice(0, 500),
  }
}

async function waitForFrameElementReady(page: Page, selector: string, timeoutMs: number, startedAt = performance.now()): Promise<void> {
  const box = await firstFrameElementBox(page, selector)
  if (box && box.width > 0 && box.height > 0) {
    return
  }
  if (performance.now() - startedAt >= timeoutMs) {
    throw new Error(`Timed out waiting for frame element ${selector}`)
  }
  await page.waitForTimeout(250)
  return waitForFrameElementReady(page, selector, timeoutMs, startedAt)
}

async function firstFrameElementBox(page: Page, selector: string): Promise<{ x: number; y: number; width: number; height: number } | null> {
  const boxes = await Promise.all(
    page.frames().map(async (frame) => {
      const locator = frame.locator(selector).first()
      const count = await locator.count().catch(() => 0)
      if (count === 0) {
        return null
      }
      const box = await locator.boundingBox().catch(() => null)
      if (box && box.width > 0 && box.height > 0) {
        return box
      }
      return null
    }),
  )
  return boxes.find((box): box is { x: number; y: number; width: number; height: number } => box !== null) ?? null
}
