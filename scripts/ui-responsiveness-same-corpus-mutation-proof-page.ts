import { createHash } from 'node:crypto'

import type { Locator, Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { captureProductScreenshot } from './ui-responsiveness-same-corpus-pixel-proof-page.ts'
import {
  readSameCorpusVisibleSelectedRange,
  type SameCorpusMutationTargetReadback,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'
import { readBiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface-page.ts'
import { settleFrames } from './ui-responsiveness-same-corpus-page-utils.ts'

export interface SameCorpusMutationTargetSelection {
  readonly endAddress: string
  readonly sheetName: string
  readonly startAddress: string
  readonly targetRange: string
}

export interface SameCorpusMutationTargetRevisionProof {
  readonly authoritativeReadbackRevision: string | null
  readonly visibleRenderRevision: string | null
}

export async function readSameCorpusMutationTargetSelection(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sheetName: string
}): Promise<SameCorpusMutationTargetSelection> {
  const selectedRange = await readSameCorpusVisibleSelectedRange(args.page, args.product)
  return normalizeSameCorpusMutationTargetSelection(selectedRange, args.sheetName)
}

export async function selectSameCorpusMutationTargetRange(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly target: SameCorpusMutationTargetSelection
}): Promise<void> {
  if (args.product === 'bilig') {
    const nameBox = args.page.getByTestId('name-box')
    await nameBox.fill(args.target.targetRange)
    await nameBox.press('Enter')
    await settleFrames(args.page, 8)
    return
  }
  await fillFirstAvailableLocator(
    [
      args.page.locator('#t-name-box input'),
      args.page.locator('input[aria-label="Name box"]'),
      args.page.locator('[aria-label="Name box"] input'),
    ],
    args.target.targetRange,
    `Cannot select same-corpus target range ${args.target.targetRange} on ${args.product}`,
  )
  await settleFrames(args.page, 8)
}

export async function readSameCorpusMutationTargetReadback(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly target: SameCorpusMutationTargetSelection
}): Promise<SameCorpusMutationTargetReadback> {
  if (args.product === 'bilig') {
    const biligReadback = await readBiligMutationTargetReadback(args.page, args.target)
    if (biligReadback) {
      return biligReadback
    }
  }
  const formulaBarText = await readVisibleFormulaBarText(args.page, args.product)
  const fillColor = await readSelectedFillColor(args.page, args.product)
  const text = normalizeNullableText(formulaBarText)
  return {
    value: text && !text.startsWith('=') ? text : null,
    formula: text?.startsWith('=') ? text : null,
    fillColor,
    visibleText: text,
  }
}

export async function captureSameCorpusMutationTargetScreenshotSha256(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
): Promise<string | null> {
  const screenshot = await captureProductScreenshot(page, product)
  return screenshot.buffer ? screenshotBufferSha256(screenshot.buffer) : null
}

export async function readSameCorpusMutationTargetRevisionProof(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly target: SameCorpusMutationTargetSelection
  readonly readback: SameCorpusMutationTargetReadback
  readonly screenshotSha256: string | null
}): Promise<SameCorpusMutationTargetRevisionProof> {
  if (args.product === 'bilig') {
    const surface = await readBiligRenderedSurfaceState(args.page)
    return {
      authoritativeReadbackRevision:
        surface?.gridAuthoritativeRenderRevision ??
        surface?.typeGpu?.authoritativeRenderRevision ??
        surface?.typeGpu?.currentWorkbookRevision ??
        null,
      visibleRenderRevision:
        surface?.typeGpu?.visibleRenderRevision ?? surface?.typeGpu?.tileSceneRevision ?? surface?.gridProjectedRenderRevision ?? null,
    }
  }
  const readbackHash = sha256Hex(stableJsonBytes({ product: args.product, readback: args.readback, target: args.target }))
  return {
    authoritativeReadbackRevision: `${args.product}-target-readback-sha256:${readbackHash}`,
    visibleRenderRevision: args.screenshotSha256 ? `${args.product}-screenshot-sha256:${args.screenshotSha256}` : null,
  }
}

export function normalizeSameCorpusMutationTargetSelection(
  selectedRange: string | null,
  sheetName: string,
): SameCorpusMutationTargetSelection {
  const rawRange = selectedRange?.split('!').at(-1)?.replace(/\$/gu, '').trim().toUpperCase() ?? ''
  const match = rawRange.match(/[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?/u)
  const targetRange = match?.[0] ?? 'A1'
  const [startAddress, endAddress = startAddress] = targetRange.split(':')
  return {
    endAddress,
    sheetName,
    startAddress,
    targetRange,
  }
}

async function readBiligMutationTargetReadback(
  page: Page,
  target: SameCorpusMutationTargetSelection,
): Promise<SameCorpusMutationTargetReadback | null> {
  const range = await page.evaluate(
    ({ endAddress, sheetName, startAddress }) => {
      return window.__biligSameCorpusProof?.readRange(sheetName, startAddress, endAddress) ?? null
    },
    {
      endAddress: target.endAddress,
      sheetName: target.sheetName,
      startAddress: target.startAddress,
    },
  )
  const cell = range?.range?.rows?.[0]?.[0]
  if (!cell) {
    return null
  }
  const formula = normalizeNullableText(cell.formula)
  const value = normalizeUnknownCellValue(cell.value ?? cell.input)
  const fillColor = normalizeNullableText(readCellFillColor(cell.style))
  return {
    value,
    formula,
    fillColor,
    visibleText: value ?? formula,
  }
}

async function readVisibleFormulaBarText(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<string | null> {
  if (product === 'bilig') {
    return await page
      .getByTestId('formula-input')
      .inputValue({ timeout: 1_000 })
      .catch(() => null)
  }
  return await page.evaluate(() => {
    const selectors = [
      '#t-formula-bar-input',
      '[aria-label="Formula bar"]',
      '[aria-label*="formula" i]',
      '[contenteditable="true"][aria-label*="formula" i]',
      'textarea[aria-label*="formula" i]',
      'input[aria-label*="formula" i]',
    ]
    for (const selector of selectors) {
      const element = document.querySelector<HTMLInputElement | HTMLTextAreaElement | HTMLElement>(selector)
      const value =
        element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement
          ? element.value
          : (element?.textContent ?? element?.getAttribute('aria-label') ?? '')
      if (value.trim().length > 0) {
        return value
      }
    }
    const active = document.activeElement
    return active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement ? active.value : (active?.textContent ?? null)
  })
}

async function readSelectedFillColor(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<string | null> {
  if (product === 'bilig') {
    const color = await page
      .getByLabel('Fill color', { exact: true })
      .first()
      .getAttribute('data-current-color', { timeout: 1_000 })
      .catch(() => null)
    return normalizeNullableText(color)
  }
  return await page.evaluate(() => {
    const selectors = ['[aria-label="Fill color"]', '#fillColorMenuButton', '[aria-label*="fill color" i]']
    for (const selector of selectors) {
      const element = document.querySelector<HTMLElement>(selector)
      if (!element) {
        continue
      }
      const dataColor = element.getAttribute('data-current-color')
      if (dataColor?.trim()) {
        return dataColor.trim()
      }
      const directColor = getComputedStyle(element).backgroundColor
      if (directColor && directColor !== 'rgba(0, 0, 0, 0)' && directColor !== 'transparent') {
        return directColor
      }
      const swatch = Array.from(element.querySelectorAll<HTMLElement>('*')).find((candidate) => {
        const color = getComputedStyle(candidate).backgroundColor
        return color && color !== 'rgba(0, 0, 0, 0)' && color !== 'transparent'
      })
      if (swatch) {
        return getComputedStyle(swatch).backgroundColor
      }
    }
    return null
  })
}

async function fillFirstAvailableLocator(
  locators: readonly Locator[],
  value: string,
  errorMessage: string,
  index = 0,
  lastError: unknown = null,
): Promise<void> {
  const locator = locators[index]
  if (!locator) {
    throw new Error(errorMessage, { cause: lastError })
  }
  try {
    await locator.first().fill(value, { timeout: 1_500 })
    await locator.first().press('Enter', { timeout: 1_500 })
  } catch (error: unknown) {
    await fillFirstAvailableLocator(locators, value, errorMessage, index + 1, error)
  }
}

function readCellFillColor(style: unknown): string | null {
  if (!style || typeof style !== 'object') {
    return null
  }
  const fill = Reflect.get(style, 'fill')
  if (!fill || typeof fill !== 'object') {
    return null
  }
  const color = Reflect.get(fill, 'backgroundColor')
  return typeof color === 'string' ? color : null
}

function normalizeUnknownCellValue(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null
  }
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  const serialized = JSON.stringify(value)
  return serialized === undefined ? null : serialized
}

function normalizeNullableText(value: string | null | undefined): string | null {
  const trimmed = value?.trim() ?? ''
  return trimmed.length > 0 ? trimmed : null
}

interface ScreenshotBuffer {
  toString(encoding: 'base64'): string
}

function screenshotBufferSha256(screenshotBuffer: ScreenshotBuffer): string {
  return createHash('sha256')
    .update(Buffer.from(screenshotBuffer.toString('base64'), 'base64'))
    .digest('hex')
}

function stableJsonBytes(value: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(stableJsonValue(value)))
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entryValue]) => entryValue !== undefined)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entryValue]) => [key, stableJsonValue(entryValue)]),
    )
  }
  return value
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}
