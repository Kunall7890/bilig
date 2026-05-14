import { performance } from 'node:perf_hooks'

import type { Page } from '@playwright/test'

import type { CaptureArgs } from './ui-responsiveness-same-corpus-args.ts'
import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'
import { collectFrameIntervals, waitForNextFrame } from './ui-responsiveness-same-corpus-page-utils.ts'

export interface ProductOperationSample {
  readonly operationResponseMs: number
  readonly postOperationFrameMs: number
  readonly scrollEventResponseMs?: number
  readonly scrollMovementPx?: number
}

export interface SameCorpusWorkloadRunnerHooks {
  readonly measureVisibleScrollResponseWithRetries: (
    page: Page,
    product: UiResponsivenessSameCorpusProduct,
    deltaX: number,
    deltaY: number,
  ) => Promise<ProductOperationSample>
  readonly movePointerToProductViewport: (page: Page, product: UiResponsivenessSameCorpusProduct) => Promise<void>
}

type NonScrollWorkload = Exclude<
  UiResponsivenessSameCorpusWorkload,
  'open-workbook' | 'scroll-vertical' | 'scroll-horizontal' | 'wide-sheet-navigation'
>

export async function measureProductWorkload(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly captureArgs: CaptureArgs
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleIndex: number
  readonly loadToReadyMs: number
  readonly hooks: SameCorpusWorkloadRunnerHooks
}): Promise<ProductOperationSample> {
  if (args.workload === 'open-workbook') {
    return await sampleSettledOperation(args.page, args.loadToReadyMs)
  }
  if (args.workload === 'scroll-vertical') {
    return await args.hooks.measureVisibleScrollResponseWithRetries(args.page, args.product, 0, args.captureArgs.deltaY)
  }
  if (args.workload === 'scroll-horizontal') {
    return await args.hooks.measureVisibleScrollResponseWithRetries(args.page, args.product, Math.max(args.captureArgs.deltaX, 720), 0)
  }
  if (args.workload === 'wide-sheet-navigation') {
    return await args.hooks.measureVisibleScrollResponseWithRetries(args.page, args.product, Math.max(args.captureArgs.deltaX, 1440), 0)
  }
  return await measureNonScrollProductWorkload(args.page, args.product, args.workload, args.sampleIndex, args.hooks)
}

async function measureNonScrollProductWorkload(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
  hooks: SameCorpusWorkloadRunnerHooks,
): Promise<ProductOperationSample> {
  await hooks.movePointerToProductViewport(page, product)
  const startedAt = performance.now()
  await performProductUiOperation(page, product, workload, sampleIndex)
  return await sampleSettledOperation(page, performance.now() - startedAt)
}

async function sampleSettledOperation(page: Page, operationResponseMs: number): Promise<ProductOperationSample> {
  await waitForNextFrame(page)
  const frameIntervals = await collectFrameIntervals(page, 12)
  return {
    operationResponseMs,
    postOperationFrameMs: percentile(frameIntervals, 0.95),
  }
}

async function performProductUiOperation(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
): Promise<void> {
  if (product === 'bilig') {
    await performBiligUiOperation(page, workload, sampleIndex)
    return
  }
  await performIncumbentUiOperation(page, product, workload, sampleIndex)
}

async function performBiligUiOperation(page: Page, workload: NonScrollWorkload, sampleIndex: number): Promise<void> {
  const nameBox = page.getByTestId('name-box')
  const formulaInput = page.getByTestId('formula-input')
  if (workload === 'select-cell') {
    await nameBox.fill('C12')
    await nameBox.press('Enter')
    return
  }
  if (workload === 'jump-deep-row') {
    await nameBox.fill('A2000')
    await nameBox.press('Enter')
    return
  }
  if (workload === 'fill-format-change') {
    await nameBox.fill('D12')
    await nameBox.press('Enter')
    await page.getByLabel('Bold').click()
    return
  }
  await nameBox.fill(workload === 'formula-edit' ? 'E12' : 'D12')
  await nameBox.press('Enter')
  await formulaInput.fill(workload === 'formula-edit' ? `=${String(sampleIndex + 1)}+1` : `same-corpus-${String(sampleIndex + 1)}`)
  await formulaInput.press('Enter')
}

async function performIncumbentUiOperation(
  page: Page,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  workload: NonScrollWorkload,
  sampleIndex: number,
): Promise<void> {
  if (workload === 'select-cell') {
    await page.keyboard.press('ArrowRight')
    return
  }
  if (workload === 'jump-deep-row') {
    await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowDown' : 'Control+ArrowDown')
    return
  }
  await assertIncumbentEditableForWorkload(page, product, workload)
  if (workload === 'fill-format-change') {
    await page.keyboard.press(process.platform === 'darwin' ? 'Meta+B' : 'Control+B')
    return
  }
  const value = workload === 'formula-edit' ? `=${String(sampleIndex + 1)}+1` : `${product}-same-corpus-${String(sampleIndex + 1)}`
  await page.keyboard.type(value)
  await page.keyboard.press('Enter')
}

async function assertIncumbentEditableForWorkload(
  page: Page,
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  workload: NonScrollWorkload,
): Promise<void> {
  const bodyText = await page
    .locator('body')
    .innerText({ timeout: 2_000 })
    .catch(() => '')
  const blocker = incumbentEditableWorkloadBlocker(product, page.url(), bodyText)
  if (blocker) {
    throw new Error(`Cannot measure ${workload} on ${product}: ${blocker}`)
  }
}

export function incumbentEditableWorkloadBlocker(
  product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>,
  pageUrl: string,
  bodyText: string,
): string | null {
  const normalizedBody = bodyText.replace(/\s+/g, ' ').toLowerCase()
  if (product === 'google-sheets') {
    if (normalizedBody.includes('view only') || normalizedBody.includes('comment only') || normalizedBody.includes('request edit access')) {
      return 'Google Sheets page is read-only; provide an editable same-corpus Google Sheet URL or authenticated storage state.'
    }
    return null
  }
  if (pageUrl.includes('view.officeapps.live.com/op/view.aspx')) {
    return 'Microsoft Excel Web URL is the read-only Office viewer; provide an editable Excel Web workbook URL for edit workloads.'
  }
  if (normalizedBody.includes('view only') || normalizedBody.includes('read-only') || normalizedBody.includes('request edit access')) {
    return 'Microsoft Excel Web page is read-only; provide an editable same-corpus workbook URL or authenticated storage state.'
  }
  return null
}

function percentile(values: readonly number[], percentileValue: number): number {
  if (values.length === 0) {
    throw new Error('Cannot compute percentile for an empty same-corpus UI sample set')
  }
  const sorted = [...values].toSorted((left, right) => left - right)
  const index = Math.ceil(percentileValue * sorted.length) - 1
  return sorted[Math.max(0, Math.min(sorted.length - 1, index))]!
}
