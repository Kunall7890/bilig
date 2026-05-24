import { createHash } from 'node:crypto'
import { performance } from 'node:perf_hooks'

import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { readBiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface-page.ts'
import { waitForNextFrame } from './ui-responsiveness-same-corpus-page-utils.ts'
import {
  sampleSettledOperation,
  type NonScrollWorkload,
  type ProductOperationSample,
  type SameCorpusProductOperation,
} from './ui-responsiveness-same-corpus-workload-runner.ts'

const visibleResponseTimeoutMs = 5_000

export interface VisibleNonScrollResponseSignature {
  readonly biligPresentedToken: string | null
  readonly product: UiResponsivenessSameCorpusProduct
  readonly screenshotSignature: string | null
  readonly workload: NonScrollWorkload
}

export async function measureVisibleNonScrollResponse(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  _sampleIndex: number,
  runOperation: SameCorpusProductOperation,
): Promise<ProductOperationSample> {
  const before = await readVisibleNonScrollResponseSignature(page, product, workload)
  const startedAt = performance.now()
  await runOperation()
  await waitForVisibleNonScrollResponse(page, product, workload, before)
  return await sampleSettledOperation(page, performance.now() - startedAt)
}

export function visibleNonScrollResponseChanged(
  before: VisibleNonScrollResponseSignature,
  after: VisibleNonScrollResponseSignature,
): boolean {
  if (before.product !== after.product || before.workload !== after.workload) {
    return false
  }
  if (before.product === 'bilig' && before.biligPresentedToken && after.biligPresentedToken) {
    return before.biligPresentedToken !== after.biligPresentedToken
  }
  return Boolean(before.screenshotSignature && after.screenshotSignature && before.screenshotSignature !== after.screenshotSignature)
}

async function waitForVisibleNonScrollResponse(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  before: VisibleNonScrollResponseSignature,
  startedAt = performance.now(),
): Promise<void> {
  if (performance.now() - startedAt >= visibleResponseTimeoutMs) {
    throw new Error(
      `Timed out waiting for browser-visible response after ${workload} on ${product}; operation-only timing is not accepted for same-corpus UI evidence.`,
    )
  }
  await waitForNextFrame(page)
  const after = await readVisibleNonScrollResponseSignature(page, product, workload)
  if (visibleNonScrollResponseChanged(before, after)) {
    return
  }
  await page.waitForTimeout(25)
  return waitForVisibleNonScrollResponse(page, product, workload, before, startedAt)
}

async function readVisibleNonScrollResponseSignature(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
): Promise<VisibleNonScrollResponseSignature> {
  return {
    biligPresentedToken: product === 'bilig' ? await readBiligPresentedResponseToken(page, workload) : null,
    product,
    screenshotSignature: await readViewportScreenshotSignature(page),
    workload,
  }
}

async function readBiligPresentedResponseToken(page: Page, workload: NonScrollWorkload): Promise<string | null> {
  const surface = await readBiligRenderedSurfaceState(page)
  const canvas = surface?.typeGpu
  if (!canvas) {
    return null
  }
  if (workload === 'select-cell') {
    return canvas.presentedSelectionRevision ?? null
  }
  if (workload === 'jump-deep-row') {
    return canvas.presentedViewportRevision ?? null
  }
  return (
    [
      canvas.presentedSemanticMutationRevision,
      canvas.presentedWorkbookRevision,
      canvas.presentedContentSignature,
      canvas.presentedTextSignature,
      canvas.presentedRectSignature,
    ]
      .filter((entry): entry is string => Boolean(entry))
      .join('|') || null
  )
}

async function readViewportScreenshotSignature(page: Page): Promise<string | null> {
  const screenshot = await page.screenshot({ fullPage: false, timeout: 2_000 }).catch(() => null)
  if (!screenshot) {
    return null
  }
  return createHash('sha256').update(screenshot).digest('hex')
}
