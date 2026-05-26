import type { Locator, Page } from '@playwright/test'

import type { CaptureArgs } from './ui-responsiveness-same-corpus-args.ts'
import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { SameCorpusOperationResponseProof } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { uiSameCorpusWorkloadMutatesWorkbook, type UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'
import { collectFrameIntervals, settleFrames, waitForNextFrame } from './ui-responsiveness-same-corpus-page-utils.ts'

export interface ProductOperationSample {
  readonly operationResponseMs: number
  readonly operationResponseProof: SameCorpusOperationResponseProof
  readonly authoritativeRenderProofMs?: number
  readonly committedTargetProofMs?: number
  readonly postOperationFrameMs: number
  readonly scrollEventResponseMs?: number
  readonly scrollMovementPx?: number
}

export type SameCorpusProductOperation = () => Promise<void>

export interface SameCorpusWorkloadRunnerHooks {
  readonly measureVisibleScrollResponseWithRetries: (
    page: Page,
    product: UiResponsivenessSameCorpusProduct,
    deltaX: number,
    deltaY: number,
  ) => Promise<ProductOperationSample>
  readonly measureVisibleNonScrollResponse: (
    page: Page,
    product: UiResponsivenessSameCorpusProduct,
    workload: NonScrollWorkload,
    sampleIndex: number,
    runOperation: SameCorpusProductOperation,
  ) => Promise<ProductOperationSample>
  readonly movePointerToProductViewport: (page: Page, product: UiResponsivenessSameCorpusProduct) => Promise<void>
}

export type NonScrollWorkload = Exclude<
  UiResponsivenessSameCorpusWorkload,
  'open-workbook' | 'scroll-vertical' | 'scroll-horizontal' | 'wide-sheet-navigation'
>

type SameCorpusKeyboardOperation = { kind: 'press'; key: string } | { kind: 'type'; text: string }
type MutatingSameCorpusWorkload = Extract<NonScrollWorkload, 'edit-visible-cell' | 'formula-edit' | 'fill-format-change'>

const sameCorpusFillColorSwatches = [
  { label: 'light cornflower blue 3', value: '#c9daf8' },
  { label: 'theme green', value: '#34a853' },
  { label: 'light cornflower blue 2', value: '#a4c2f4' },
] as const

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
    return await sampleSettledOperation(args.page, args.loadToReadyMs, 'load-to-ready')
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

export async function restoreProductWorkbookMutation(
  page: Page,
  workload: UiResponsivenessSameCorpusWorkload,
  platform: NodeJS.Platform = process.platform,
): Promise<void> {
  const operations = sameCorpusWorkbookRestoreOperations(workload, platform)
  if (operations.length === 0) {
    return
  }
  await performSameCorpusKeyboardOperations(page, operations)
  await settleFrames(page, 30)
}

async function measureNonScrollProductWorkload(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
  hooks: SameCorpusWorkloadRunnerHooks,
): Promise<ProductOperationSample> {
  await hooks.movePointerToProductViewport(page, product)
  return await hooks.measureVisibleNonScrollResponse(page, product, workload, sampleIndex, async () => {
    await performProductUiOperation(page, product, workload, sampleIndex)
  })
}

export async function sampleSettledOperation(
  page: Page,
  operationResponseMs: number,
  operationResponseProof: SameCorpusOperationResponseProof,
): Promise<ProductOperationSample> {
  await waitForNextFrame(page)
  const frameIntervals = await collectFrameIntervals(page, 12)
  return {
    operationResponseMs,
    operationResponseProof,
    postOperationFrameMs: percentile(frameIntervals, 0.95),
  }
}

async function performProductUiOperation(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
): Promise<void> {
  if (product !== 'bilig') {
    await assertIncumbentEditableForWorkload(page, product, workload)
  }
  if (workload === 'fill-format-change') {
    await performSameCorpusFillColorOperation(page, product, sampleIndex)
    return
  }
  await performSameCorpusKeyboardOperations(page, sameCorpusKeyboardOperations(product, workload, sampleIndex))
}

export function sameCorpusFillColorSwatchLabel(sampleIndex: number): string {
  return sameCorpusFillColorSwatches[sampleIndex % sameCorpusFillColorSwatches.length].label
}

export function sameCorpusFillColorExpectedColor(sampleIndex: number): string {
  return sameCorpusFillColorSwatches[sampleIndex % sameCorpusFillColorSwatches.length].value
}

export function sameCorpusFillColorExpectedColors(): readonly string[] {
  return sameCorpusFillColorSwatches.map((swatch) => swatch.value)
}

export function sameCorpusKeyboardOperations(
  product: UiResponsivenessSameCorpusProduct,
  workload: NonScrollWorkload,
  sampleIndex: number,
  platform: NodeJS.Platform = process.platform,
): readonly SameCorpusKeyboardOperation[] {
  if (workload === 'select-cell') {
    return [{ kind: 'press', key: 'ArrowRight' }]
  }
  if (workload === 'jump-deep-row') {
    return [{ kind: 'press', key: primaryShortcut('ArrowDown', platform) }]
  }
  if (workload === 'fill-format-change') {
    return []
  }
  const value = workload === 'formula-edit' ? `=${String(sampleIndex + 1)}+1` : `${product}-same-corpus-${String(sampleIndex + 1)}`
  return [
    { kind: 'type', text: value },
    { kind: 'press', key: 'Enter' },
  ]
}

export function sameCorpusWorkbookRestoreOperations(
  workload: UiResponsivenessSameCorpusWorkload,
  platform: NodeJS.Platform = process.platform,
): readonly SameCorpusKeyboardOperation[] {
  return sameCorpusWorkloadMutatesWorkbook(workload) ? [{ kind: 'press', key: primaryShortcut('Z', platform) }] : []
}

export function sameCorpusWorkloadMutatesWorkbook(workload: UiResponsivenessSameCorpusWorkload): workload is MutatingSameCorpusWorkload {
  return uiSameCorpusWorkloadMutatesWorkbook(workload)
}

function primaryShortcut(key: string, platform: NodeJS.Platform): string {
  return platform === 'darwin' ? `Meta+${key}` : `Control+${key}`
}

async function performSameCorpusKeyboardOperations(
  page: Page,
  operations: readonly SameCorpusKeyboardOperation[],
  index = 0,
): Promise<void> {
  const operation = operations[index]
  if (!operation) {
    return
  }
  if (operation.kind === 'type') {
    await page.keyboard.type(operation.text)
  } else {
    await page.keyboard.press(operation.key)
  }
  await performSameCorpusKeyboardOperations(page, operations, index + 1)
}

async function performSameCorpusFillColorOperation(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  sampleIndex: number,
): Promise<void> {
  const swatchLabel = sameCorpusFillColorSwatchLabel(sampleIndex)
  const swatchCandidateLabels = sameCorpusFillColorCandidateLabels(swatchLabel)
  await clickFirstAvailableLocator(
    [
      page.getByLabel('Fill color', { exact: true }),
      page.getByRole('button', { name: /^Fill color$/u }),
      page.locator('[aria-label="Fill color"]'),
    ],
    `Cannot open ${product} fill color control`,
  )
  await clickFirstAvailableLocator(
    swatchCandidateLabels.flatMap((label) => [
      page.getByLabel(`Fill color ${label}`, { exact: true }),
      page.getByLabel(new RegExp(`^${escapeRegExp(label)}$`, 'iu')),
      page.getByRole('button', { name: new RegExp(escapeRegExp(label), 'iu') }),
      page.locator(`[aria-label="${label}"]`),
    ]),
    `Cannot choose ${product} fill color swatch "${swatchLabel}"`,
  )
}

async function clickFirstAvailableLocator(
  locators: readonly Locator[],
  errorMessage: string,
  index = 0,
  lastError: unknown = null,
): Promise<void> {
  const locator = locators[index]
  if (!locator) {
    throw new Error(errorMessage, { cause: lastError })
  }
  try {
    await locator.first().click({ timeout: 1_500 })
  } catch (error: unknown) {
    await clickFirstAvailableLocator(locators, errorMessage, index + 1, error)
  }
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function sameCorpusFillColorCandidateLabels(swatchLabel: string): readonly string[] {
  if (swatchLabel === 'theme green') {
    return ['theme green', 'green']
  }
  if (swatchLabel === 'light cornflower blue 2') {
    return ['light cornflower blue 2', 'light blue 2', 'blue']
  }
  return [swatchLabel, 'light cornflower blue', 'blue']
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
    if (normalizedBody.includes(' sign in') || normalizedBody.endsWith('sign in')) {
      return 'Google Sheets page is not authenticated; provide an authenticated storage state for editable same-corpus workloads.'
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
