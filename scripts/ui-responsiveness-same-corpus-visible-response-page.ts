import { createHash } from 'node:crypto'
import { performance } from 'node:perf_hooks'

import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { BiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface.ts'
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
  readonly biligInteractionVisibleToken: string | null
  readonly product: UiResponsivenessSameCorpusProduct
  readonly screenshotSignature: string | null
  readonly workload: NonScrollWorkload
}

export async function measureVisibleNonScrollResponse(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
  runOperation: SameCorpusProductOperation,
): Promise<ProductOperationSample> {
  return await measureVisibleNonScrollResponseAfterOperationStarted(page, product, workload, sampleIndex, runOperation)
}

export async function measureVisibleNonScrollResponseAfterOperationStarted(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  _sampleIndex: number,
  runOperation: SameCorpusProductOperation,
): Promise<ProductOperationSample> {
  const before = await readVisibleNonScrollResponseSignature(page, product, workload)
  const startedAt = performance.now()
  const operationPromise = runOperation()
  const operationFailurePromise = waitForOperationFailure(operationPromise)
  try {
    const responsePromise =
      product === 'bilig' && before.biligInteractionVisibleToken
        ? waitForBiligVisibleNonScrollResponseToken(page, workload, before.biligInteractionVisibleToken)
        : waitForVisibleNonScrollResponse(page, product, workload, before)
    await Promise.race([responsePromise, operationFailurePromise])
    const interactionVisibleMs = performance.now() - startedAt
    await operationPromise
    return await sampleSettledOperation(page, interactionVisibleMs, 'visible-non-scroll-response')
  } catch (error) {
    await operationPromise.catch(() => undefined)
    throw error
  }
}

async function waitForOperationFailure(operationPromise: Promise<void>): Promise<never> {
  await operationPromise
  await new Promise<never>(() => {})
}

export function visibleNonScrollResponseChanged(
  before: VisibleNonScrollResponseSignature,
  after: VisibleNonScrollResponseSignature,
): boolean {
  if (before.product !== after.product || before.workload !== after.workload) {
    return false
  }
  if (before.product === 'bilig' && before.biligInteractionVisibleToken && after.biligInteractionVisibleToken) {
    return before.biligInteractionVisibleToken !== after.biligInteractionVisibleToken
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
  const biligSurface = product === 'bilig' ? await readBiligRenderedSurfaceState(page) : null
  const biligInteractionVisibleToken = biligInteractionVisibleResponseToken(biligSurface, workload)
  return {
    biligInteractionVisibleToken,
    product,
    screenshotSignature: visibleNonScrollResponseNeedsScreenshot(product, biligInteractionVisibleToken)
      ? await readViewportScreenshotSignature(page)
      : null,
    workload,
  }
}

export function visibleNonScrollResponseNeedsScreenshot(
  product: UiResponsivenessSameCorpusProduct,
  biligInteractionVisibleToken: string | null = null,
): boolean {
  return product !== 'bilig' || biligInteractionVisibleToken === null
}

export function biligInteractionVisibleResponseToken(
  surface: BiligRenderedSurfaceState | null,
  workload: NonScrollWorkload,
): string | null {
  const canvas = surface?.typeGpu
  if (!surface || !canvas) {
    return null
  }
  if (workload === 'edit-visible-cell' || workload === 'formula-edit') {
    return firstVisibleToken(
      surface.editorVisibleRevision,
      surface.formulaVisibleRevision,
      surface.gridEditorVisibleRevision,
      surface.gridInteractionVisibleRevision,
      canvas.currentSemanticMutationRevision,
      canvas.currentWorkbookRevision,
      canvas.currentContentSignature,
      canvas.currentTextSignature,
      canvas.currentRectSignature,
      canvas.visibleLocalRenderRevision,
      canvas.visibleProjectedRenderRevision,
      surface.gridLocalRenderRevision,
      surface.gridProjectedRenderRevision,
    )
  }
  if (workload === 'select-cell') {
    return firstVisibleToken(
      surface.gridSelectionVisibleRevision,
      canvas.currentSelectionRevision,
      canvas.visibleLocalRenderRevision,
      surface.gridLocalRenderRevision,
    )
  }
  if (workload === 'jump-deep-row') {
    return firstVisibleToken(
      surface.gridSelectionVisibleRevision,
      canvas.currentViewportRevision,
      canvas.visibleProjectedRenderRevision,
      surface.gridProjectedRenderRevision,
    )
  }
  return (
    [
      surface.gridInteractionVisibleRevision,
      canvas.currentSemanticMutationRevision,
      canvas.currentWorkbookRevision,
      canvas.currentContentSignature,
      canvas.currentTextSignature,
      canvas.currentRectSignature,
      canvas.visibleLocalRenderRevision,
      canvas.visibleProjectedRenderRevision,
      surface.gridLocalRenderRevision,
      surface.gridProjectedRenderRevision,
    ]
      .filter((entry): entry is string => Boolean(entry))
      .join('|') || null
  )
}

async function waitForBiligVisibleNonScrollResponseToken(page: Page, workload: NonScrollWorkload, beforeToken: string): Promise<void> {
  await page.waitForFunction(
    ({ previousToken, targetWorkload }) => {
      const grid = document.querySelector('[data-testid="sheet-grid"]')
      const canvas = document.querySelector('[data-testid="grid-pane-renderer"]')
      if (!(grid instanceof HTMLElement) || !(canvas instanceof HTMLElement)) {
        return false
      }
      // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright evaluates this helper inside the browser context.
      const firstDocumentToken = (...tokens: readonly (string | null | undefined)[]): string | null =>
        tokens.find((token): token is string => typeof token === 'string' && token.length > 0) ?? null
      const editor = document.querySelector('[data-testid="cell-editor-input"]')
      const editorToken =
        editor instanceof HTMLTextAreaElement
          ? [
              'cell-editor',
              editor.getAttribute('data-editor-target') ?? '',
              editor.value,
              String(editor.selectionStart ?? editor.value.length),
              String(editor.selectionEnd ?? editor.value.length),
            ].join(':')
          : null
      const formulaInput = document.querySelector('[data-testid="formula-input"]')
      const formulaToken =
        formulaInput instanceof HTMLTextAreaElement
          ? [
              'formula-input',
              formulaInput.getAttribute('data-formula-editing') ?? '',
              formulaInput.value,
              String(formulaInput.selectionStart ?? formulaInput.value.length),
              String(formulaInput.selectionEnd ?? formulaInput.value.length),
            ].join(':')
          : null
      const gridInteractionToken = grid.getAttribute('data-render-interaction-visible-revision')
      const gridSelectionToken = grid.getAttribute('data-render-selection-visible-revision')
      const gridEditorToken = grid.getAttribute('data-render-editor-visible-revision')
      const token =
        targetWorkload === 'select-cell'
          ? firstDocumentToken(
              gridSelectionToken,
              canvas.getAttribute('data-v3-current-selection-revision'),
              canvas.getAttribute('data-v3-visible-local-render-revision'),
              grid.getAttribute('data-render-local-revision'),
            )
          : targetWorkload === 'jump-deep-row'
            ? firstDocumentToken(
                gridSelectionToken,
                canvas.getAttribute('data-v3-current-viewport-revision'),
                canvas.getAttribute('data-v3-visible-projected-render-revision'),
                grid.getAttribute('data-render-projected-revision'),
              )
            : targetWorkload === 'edit-visible-cell' || targetWorkload === 'formula-edit'
              ? firstDocumentToken(
                  editorToken,
                  formulaToken,
                  gridEditorToken,
                  gridInteractionToken,
                  canvas.getAttribute('data-v3-current-semantic-mutation-revision'),
                  canvas.getAttribute('data-v3-current-workbook-revision'),
                  canvas.getAttribute('data-v3-current-content-signature'),
                  canvas.getAttribute('data-v3-current-text-signature'),
                  canvas.getAttribute('data-v3-current-rect-signature'),
                  canvas.getAttribute('data-v3-visible-local-render-revision'),
                  canvas.getAttribute('data-v3-visible-projected-render-revision'),
                  grid.getAttribute('data-render-local-revision'),
                  grid.getAttribute('data-render-projected-revision'),
                )
              : [
                  gridInteractionToken,
                  canvas.getAttribute('data-v3-current-semantic-mutation-revision'),
                  canvas.getAttribute('data-v3-current-workbook-revision'),
                  canvas.getAttribute('data-v3-current-content-signature'),
                  canvas.getAttribute('data-v3-current-text-signature'),
                  canvas.getAttribute('data-v3-current-rect-signature'),
                  canvas.getAttribute('data-v3-visible-local-render-revision'),
                  canvas.getAttribute('data-v3-visible-projected-render-revision'),
                  grid.getAttribute('data-render-local-revision'),
                  grid.getAttribute('data-render-projected-revision'),
                ]
                  .filter((entry): entry is string => Boolean(entry))
                  .join('|') || null
      return Boolean(token && token !== previousToken)
    },
    { previousToken: beforeToken, targetWorkload: workload },
    { polling: 'raf', timeout: visibleResponseTimeoutMs },
  )
}

function firstVisibleToken(...tokens: readonly (string | null | undefined)[]): string | null {
  return tokens.find((token): token is string => typeof token === 'string' && token.length > 0) ?? null
}

async function readViewportScreenshotSignature(page: Page): Promise<string | null> {
  const screenshot = await page.screenshot({ fullPage: false, timeout: 2_000 }).catch(() => null)
  if (!screenshot) {
    return null
  }
  return createHash('sha256').update(screenshot).digest('hex')
}
