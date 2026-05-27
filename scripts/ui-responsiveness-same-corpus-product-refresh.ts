import { chromium } from '@playwright/test'

import { buildWorkbookBenchmarkCorpus, isWorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'
import type { CaptureArgs, RefreshProductArgs } from './ui-responsiveness-same-corpus-args.ts'
import { buildSameCorpusCaptureArtifact, measureProduct } from './ui-responsiveness-same-corpus-page.ts'
import {
  buildCaptureScenarioProof,
  type SameCorpusProductVisualProof,
  type SameCorpusScenarioProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import {
  sameCorpusScenarioCaseFields,
  type SameCorpusCapture,
  type SameCorpusCaptureCase,
  type SameCorpusCaptureMeasurement,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { sameCorpusChromiumLaunchOptions } from './ui-responsiveness-same-corpus-page-utils.ts'

export async function refreshSameCorpusCaptureProduct(args: {
  readonly capture: SameCorpusCapture
  readonly captureArgs: CaptureArgs
  readonly product: RefreshProductArgs['product']
}): Promise<SameCorpusCapture> {
  const corpus = buildWorkbookBenchmarkCorpus(args.captureArgs.corpusId)
  assertRefreshCaptureMatchesArgs(args.capture, args.captureArgs, args.product)
  const browser = await chromium.launch(sameCorpusChromiumLaunchOptions(args.captureArgs.headless))
  try {
    const cases: SameCorpusCaptureCase[] = []
    for (const entry of args.capture.cases) {
      const visualProofs = sameCorpusVisualProofsFromScenarioProof(entry.scenarioProof, entry.id).filter(
        (proof) => proof.product !== args.product,
      )
      // oxlint-disable-next-line eslint(no-await-in-loop) -- Product refreshes run sequentially to avoid cross-workload browser resource contention.
      const measurement = await measureProduct(
        browser,
        args.product,
        args.captureArgs.biligUrl,
        corpus,
        args.captureArgs,
        entry.workload,
        entry.id,
        visualProofs,
      )
      cases.push(rebuildSameCorpusCaseWithRefreshedProduct({ entry, measurement, product: args.product, visualProofs }))
    }
    return buildSameCorpusCaptureArtifact({
      sampleCount: args.capture.sampleCount,
      limitations: refreshedCaptureLimitations(args.capture.limitations, args.product),
      cases,
    })
  } finally {
    await browser.close()
  }
}

export function captureArgsForProductRefresh(args: RefreshProductArgs, capture: SameCorpusCapture): CaptureArgs {
  const corpusId = sameCorpusCaptureCorpusId(capture)
  return {
    allowIncompleteEvidence: args.allowIncompleteEvidence,
    biligProductionHost: args.biligProductionHost,
    biligProductionPort: args.biligProductionPort,
    biligUrl:
      args.biligUrlSource === 'served-production'
        ? ''
        : (args.biligUrl ?? `http://localhost:5173/?benchmarkCorpus=${encodeURIComponent(corpusId)}`),
    biligUrlSource: args.biligUrlSource,
    biligStorageStatePath: args.biligStorageStatePath,
    corpusId,
    deltaX: 0,
    deltaY: 720,
    googleSheetsUrl: sameCorpusCaptureProductSource(capture, 'google-sheets'),
    googleSheetsStorageStatePath: null,
    headless: args.headless,
    microsoftExcelWebUrl: sameCorpusCaptureOptionalProductSource(capture, 'microsoft-excel-web'),
    microsoftExcelWebStorageStatePath: null,
    outputPath: args.outputPath,
    readyTimeoutMs: args.readyTimeoutMs,
    sampleCount: capture.sampleCount,
    storageStatePath: args.storageStatePath,
  }
}

export function rebuildSameCorpusCaseWithRefreshedProduct(args: {
  readonly entry: SameCorpusCaptureCase
  readonly measurement: SameCorpusCaptureMeasurement
  readonly product: RefreshProductArgs['product']
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
}): SameCorpusCaptureCase {
  if (args.measurement.product !== args.product) {
    throw new Error(`Cannot refresh ${args.product} same-corpus proof with ${args.measurement.product} measurement`)
  }
  const bilig = args.measurement
  const scenarioProof = buildCaptureScenarioProof({
    bilig,
    googleSheets: args.entry.googleSheets,
    microsoftExcelWeb: args.entry.microsoftExcelWeb,
    visualProofs: args.visualProofs,
    workload: args.entry.workload,
  })
  return {
    ...args.entry,
    ...sameCorpusScenarioCaseFields(scenarioProof),
    scenarioProof,
    bilig,
  }
}

export function sameCorpusVisualProofsFromScenarioProof(
  proof: SameCorpusScenarioProof,
  caseId: string,
): readonly SameCorpusProductVisualProof[] {
  return proof.pixelGridProof.products.map((pixelGridProof) => {
    const screenshotPath = sameCorpusScenarioScreenshotPathForProduct(proof, caseId, pixelGridProof.product)
    return {
      product: pixelGridProof.product,
      screenshotPath,
      screenshotCaptured: screenshotPath !== null && !proof.screenshotProof.missingProducts.includes(pixelGridProof.product),
      pixelGridProof,
      semanticUiProof: proof.semanticUiProof.products.find((entry) => entry.product === pixelGridProof.product),
    }
  })
}

function assertRefreshCaptureMatchesArgs(capture: SameCorpusCapture, args: CaptureArgs, product: RefreshProductArgs['product']): void {
  if (product !== 'bilig') {
    throw new Error('Only Bilig same-corpus product refresh is supported.')
  }
  if (capture.sampleCount !== args.sampleCount) {
    throw new Error('Same-corpus product refresh sample count does not match the existing capture.')
  }
  if (capture.cases.some((entry) => entry.corpusCaseId !== args.corpusId)) {
    throw new Error('Same-corpus product refresh corpus id does not match the existing capture.')
  }
}

function sameCorpusCaptureCorpusId(capture: SameCorpusCapture): CaptureArgs['corpusId'] {
  const corpusIds = [...new Set(capture.cases.map((entry) => entry.corpusCaseId))]
  const corpusId = corpusIds[0]
  if (corpusIds.length !== 1 || !corpusId || !isWorkbookBenchmarkCorpusId(corpusId)) {
    throw new Error('Cannot refresh same-corpus product evidence for a capture with unknown or mixed corpus ids.')
  }
  return corpusId
}

function sameCorpusCaptureProductSource(capture: SameCorpusCapture, product: UiResponsivenessSameCorpusProduct): string {
  const source = sameCorpusCaptureOptionalProductSource(capture, product)
  if (!source) {
    throw new Error(`Existing same-corpus capture is missing ${product} source evidence.`)
  }
  return source
}

function sameCorpusCaptureOptionalProductSource(capture: SameCorpusCapture, product: UiResponsivenessSameCorpusProduct): string | null {
  const sources = new Set(
    capture.cases.flatMap((entry) => {
      if (product === 'bilig') {
        return [entry.bilig.source]
      }
      if (product === 'google-sheets') {
        return [entry.googleSheets.source]
      }
      return entry.microsoftExcelWeb ? [entry.microsoftExcelWeb.source] : []
    }),
  )
  if (sources.size > 1) {
    throw new Error(`Existing same-corpus capture has mixed ${product} source evidence.`)
  }
  return [...sources][0] ?? null
}

function sameCorpusScenarioScreenshotPathForProduct(
  proof: SameCorpusScenarioProof,
  caseId: string,
  product: UiResponsivenessSameCorpusProduct,
): string | null {
  return (
    proof.screenshotProof.artifactPaths.find((path) => path.replaceAll('\\', '/').endsWith(`/${caseId}/${product}-sample-1.png`)) ?? null
  )
}

function refreshedCaptureLimitations(limitations: readonly string[], product: RefreshProductArgs['product']): readonly string[] {
  const refreshedLimitation = `The ${product} product evidence was refreshed from the existing same-corpus capture; incumbent evidence remains from the original capture.`
  return limitations.includes(refreshedLimitation) ? limitations : [...limitations, refreshedLimitation]
}
