import { createHash } from 'node:crypto'

import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { captureProductScreenshot } from './ui-responsiveness-same-corpus-pixel-proof-page.ts'
import {
  readSameCorpusVisibleSheetId,
  readSameCorpusVisibleSelectedRange,
  type SameCorpusMutationTargetReadback,
  type SameCorpusMutationTargetScreenshotProof,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'
import { readBiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface-page.ts'
import { settleFrames } from './ui-responsiveness-same-corpus-page-utils.ts'
import { sameCorpusFillColorExpectedColors } from './ui-responsiveness-same-corpus-workload-runner.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

const BILIG_ROW_MARKER_WIDTH = 36
const BILIG_COLUMN_WIDTH = 104
const BILIG_COLUMN_HEADER_HEIGHT = 20
const BILIG_ROW_HEIGHT = 22

export interface SameCorpusMutationTargetSelection {
  readonly endAddress: string
  readonly sheetName: string
  readonly sheetId: string | null
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
  const sheetId = await readSameCorpusVisibleSheetId(args.page, args.product, args.sheetName)
  return normalizeSameCorpusMutationTargetSelection(selectedRange, args.sheetName, sheetId)
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
  if (args.product === 'google-sheets') {
    await selectGoogleSheetsTargetRange(args.page, args.target.targetRange)
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

export interface SameCorpusNameBoxPage {
  readonly keyboard: {
    press(key: string): Promise<void>
  }
  locator(selector: string): SameCorpusFillableLocator
}

interface SameCorpusFillableLocator {
  first(): SameCorpusFillableLocator
  fill(value: string, options?: { readonly timeout?: number }): Promise<void>
  press(key: string, options?: { readonly timeout?: number }): Promise<void>
}

export async function selectGoogleSheetsTargetRange(page: SameCorpusNameBoxPage, targetRange: string): Promise<void> {
  await page.keyboard.press(primaryShortcut('J'))
  await fillFirstAvailableLocator(
    [
      page.locator('#t-name-box'),
      page.locator('input.waffle-name-box'),
      page.locator('input[aria-label="Name box"]'),
      page.locator('[aria-label^="Name box"] input'),
    ],
    targetRange,
    `Cannot select same-corpus target range ${targetRange} on google-sheets`,
  )
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
  return await readSameCorpusVisibleMutationTargetReadback(args)
}

export async function readSameCorpusVisibleMutationTargetReadback(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly target?: SameCorpusMutationTargetSelection
  readonly workload?: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetReadback> {
  if (args.product === 'bilig' && args.workload === 'fill-format-change' && args.target) {
    return {
      value: null,
      formula: null,
      fillColor: await readBiligVisibleGridCellFillColor(args.page, args.target),
      visibleText: null,
      source: 'visible-grid-cell',
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
    source: 'visible-formula-bar',
  }
}

export async function captureSameCorpusMutationTargetScreenshotProof(args: {
  readonly page: Page
  readonly sampleIndex: number
  readonly product: UiResponsivenessSameCorpusProduct
  readonly target: SameCorpusMutationTargetSelection
  readonly phase: SameCorpusMutationTargetScreenshotProof['phase']
  readonly screenshotPath: string
  readonly relativeScreenshotPath: string
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetScreenshotProof> {
  const screenshot = await captureMutationTargetScreenshot(args.page, args.product, args.target, args.screenshotPath)
  return {
    phase: args.phase,
    product: args.product,
    scope: screenshot.scope,
    sampleIndex: args.sampleIndex,
    sheetId: args.target.sheetId,
    sheetName: args.target.sheetName,
    targetRange: args.target.targetRange,
    workload: args.workload,
    screenshotPath: screenshot.captured ? args.relativeScreenshotPath : null,
    screenshotSha256: screenshot.buffer ? screenshotBufferSha256(screenshot.buffer) : null,
  }
}

async function captureMutationTargetScreenshot(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  target: SameCorpusMutationTargetSelection,
  screenshotPath: string,
): Promise<{
  readonly buffer: ScreenshotBuffer | null
  readonly captured: boolean
  readonly scope: SameCorpusMutationTargetScreenshotProof['scope']
}> {
  if (product === 'bilig') {
    const address = sameCorpusCellAddressCoordinates(target.startAddress)
    const buffer = address ? await captureBiligCellInteriorScreenshot(page, address.columnIndex, address.rowIndex, screenshotPath) : null
    return { buffer, captured: Boolean(buffer), scope: 'target-cell' }
  }
  const selectedTargetBox = await firstVisibleTargetBox(page, product)
  if (selectedTargetBox) {
    const buffer = await page.screenshot({
      animations: 'disabled',
      caret: 'hide',
      clip: selectedTargetBox,
      path: screenshotPath,
    })
    return { buffer, captured: true, scope: 'target-cell' }
  }
  const screenshot = await captureProductScreenshot(page, product, screenshotPath)
  return { ...screenshot, scope: 'visible-grid-fallback' }
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
        args.readback.capturedRevision ??
        surface?.gridAuthoritativeRenderRevision ??
        surface?.typeGpu?.authoritativeRenderRevision ??
        surface?.typeGpu?.currentWorkbookRevision ??
        null,
      visibleRenderRevision:
        (args.readback.visibleSceneProofSha256 ? `bilig-visible-scene-sha256:${args.readback.visibleSceneProofSha256}` : null) ??
        surface?.typeGpu?.visibleRenderRevision ??
        surface?.typeGpu?.tileSceneRevision ??
        surface?.gridProjectedRenderRevision ??
        null,
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
  sheetId: string | null = null,
): SameCorpusMutationTargetSelection {
  const rawRange = selectedRange?.split('!').at(-1)?.replace(/\$/gu, '').trim().toUpperCase() ?? ''
  const match = rawRange.match(/^[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?$/u)
  if (!match) {
    throw new Error(`Cannot derive same-corpus mutation target range from visible selection: ${selectedRange ?? 'missing'}`)
  }
  const targetRange = match[0]
  const [startAddress, endAddress = startAddress] = targetRange.split(':')
  return {
    endAddress,
    sheetName,
    sheetId,
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
    batchId: readNullableNumber(range.batchId),
    capturedRevision: normalizeNullableText(normalizeUnknownCellValue(range.capturedRevision)),
    value,
    formula,
    fillColor,
    visibleText: value ?? formula,
    source: 'bilig-authoritative-range',
    visibleSceneProofSha256: range.visibleSceneProof ? sha256Hex(stableJsonBytes(range.visibleSceneProof)) : null,
  }
}

async function readBiligVisibleGridCellFillColor(page: Page, target: SameCorpusMutationTargetSelection): Promise<string | null> {
  const address = sameCorpusCellAddressCoordinates(target.startAddress)
  if (!address) {
    return null
  }
  const screenshot = await captureBiligCellInteriorScreenshot(page, address.columnIndex, address.rowIndex)
  return screenshot ? await readExpectedFillColorFromScreenshot(page, screenshot) : null
}

async function captureBiligCellInteriorScreenshot(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  screenshotPath?: string,
): Promise<ScreenshotBuffer | null> {
  const gridLocator = page.getByTestId('sheet-grid')
  if ((await gridLocator.count().catch(() => 0)) === 0) {
    return null
  }
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    return null
  }
  const geometry = await page.evaluate(
    ({ columnIndex: targetColumnIndex, columnWidthFallback, rowHeightFallback, rowIndex: targetRowIndex, rowMarkerWidth }) => {
      const gridElement = document.querySelector('[data-testid="sheet-grid"]')
      if (!(gridElement instanceof HTMLElement)) {
        return null
      }
      const scrollViewport = document.querySelector('[data-testid="grid-scroll-viewport"]')
      const scrollElement = scrollViewport instanceof HTMLElement ? scrollViewport : null
      // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright evaluates this helper inside the browser context.
      const parseOverrides = (raw: string | null): Record<string, number> => {
        if (!raw) {
          return {}
        }
        try {
          const parsed: unknown = JSON.parse(raw)
          if (!parsed || typeof parsed !== 'object') {
            return {}
          }
          return Object.fromEntries(Object.entries(parsed).filter((entry): entry is [string, number] => typeof entry[1] === 'number'))
        } catch {
          return {}
        }
      }
      const columnWidthDefault = Number(gridElement.getAttribute('data-default-column-width') ?? String(columnWidthFallback))
      const rowHeightDefault = Number(gridElement.getAttribute('data-default-row-height') ?? String(rowHeightFallback))
      const columnOverrides = parseOverrides(gridElement.getAttribute('data-column-width-overrides'))
      const rowOverrides = parseOverrides(gridElement.getAttribute('data-row-height-overrides'))
      const columnWidth = (index: number) => columnOverrides[String(index)] ?? columnWidthDefault
      const rowHeight = (index: number) => rowOverrides[String(index)] ?? rowHeightDefault
      const columnLeft =
        rowMarkerWidth +
        Array.from({ length: targetColumnIndex }, (_, index) => columnWidth(index)).reduce((total, width) => total + width, 0)
      const rowTop = Array.from({ length: targetRowIndex }, (_, index) => rowHeight(index)).reduce((total, height) => total + height, 0)
      return {
        columnLeft,
        columnWidth: columnWidth(targetColumnIndex),
        rowHeight: rowHeight(targetRowIndex),
        rowTop,
        scrollLeft: scrollElement?.scrollLeft ?? 0,
        scrollTop: scrollElement?.scrollTop ?? 0,
      }
    },
    {
      columnIndex,
      columnWidthFallback: BILIG_COLUMN_WIDTH,
      rowHeightFallback: BILIG_ROW_HEIGHT,
      rowIndex,
      rowMarkerWidth: BILIG_ROW_MARKER_WIDTH,
    },
  )
  if (!geometry) {
    return null
  }
  const viewport = page.viewportSize() ?? (await page.evaluate(() => ({ height: window.innerHeight, width: window.innerWidth })))
  const clipX = Math.round(grid.x + geometry.columnLeft - geometry.scrollLeft + 8)
  const clipY = Math.round(grid.y + BILIG_COLUMN_HEADER_HEIGHT + geometry.rowTop - geometry.scrollTop + 5)
  const clipWidth = Math.max(1, Math.round(geometry.columnWidth - 16))
  const clipHeight = Math.max(1, Math.round(geometry.rowHeight - 10))
  const x0 = Math.max(0, clipX)
  const y0 = Math.max(0, clipY)
  const x1 = Math.min(viewport.width, clipX + clipWidth)
  const y1 = Math.min(viewport.height, clipY + clipHeight)
  if (x1 <= x0 || y1 <= y0) {
    return null
  }
  return await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: { height: y1 - y0, width: x1 - x0, x: x0, y: y0 },
    ...(screenshotPath ? { path: screenshotPath } : {}),
  })
}

async function firstVisibleTargetBox(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
): Promise<{ readonly x: number; readonly y: number; readonly width: number; readonly height: number } | null> {
  const selectors =
    product === 'google-sheets'
      ? [
          '.waffle-cell-input',
          '.waffle-active-cell',
          '.waffle-border-cell-active',
          '[class*="active-cell" i]',
          '[class*="selected-cell" i]',
        ]
      : ['.ewr-selection', '.ewr-active-cell', '[class*="active-cell" i]', '[class*="selected-cell" i]']
  const frames = page.frames()
  const targetBoxPromises = selectors.flatMap((selector) =>
    [visibleLocatorBox(page.locator(selector).first())].concat(frames.map((frame) => visibleLocatorBox(frame.locator(selector).first()))),
  )
  const boxes = await Promise.all(targetBoxPromises)
  return boxes.find((box): box is NonNullable<(typeof boxes)[number]> => box !== null) ?? null
}

async function visibleLocatorBox(locator: {
  boundingBox(): Promise<{ x: number; y: number; width: number; height: number } | null>
  count(): Promise<number>
}) {
  if ((await locator.count().catch(() => 0)) === 0) {
    return null
  }
  const box = await locator.boundingBox().catch(() => null)
  if (!box || box.width <= 2 || box.height <= 2) {
    return null
  }
  return {
    x: Math.max(0, Math.floor(box.x)),
    y: Math.max(0, Math.floor(box.y)),
    width: Math.max(1, Math.ceil(box.width)),
    height: Math.max(1, Math.ceil(box.height)),
  }
}

async function readExpectedFillColorFromScreenshot(page: Page, screenshot: ScreenshotBuffer): Promise<string | null> {
  return await page.evaluate(
    async ({ dataUrl, expectedColors }) => {
      const image = await new Promise<HTMLImageElement>((resolveImage, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolveImage(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode rendered fill proof screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d', { willReadFrequently: true })
      if (!context) {
        throw new Error('Missing 2d context for rendered fill proof screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
      const targetColors = expectedColors
        .map((hexColor) => {
          const match = hexColor.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/iu)
          return match
            ? {
                blue: Number.parseInt(match[3] ?? '0', 16),
                green: Number.parseInt(match[2] ?? '0', 16),
                hex: hexColor.toLowerCase(),
                red: Number.parseInt(match[1] ?? '0', 16),
              }
            : null
        })
        .filter(
          (color): color is { readonly blue: number; readonly green: number; readonly hex: string; readonly red: number } => color !== null,
        )
      const counts = Array.from({ length: targetColors.length }, () => 0)
      let colorfulPixels = 0
      for (let index = 0; index < pixels.length; index += 4) {
        const alpha = pixels[index + 3] ?? 0
        const red = pixels[index] ?? 255
        const green = pixels[index + 1] ?? 255
        const blue = pixels[index + 2] ?? 255
        if (alpha <= 220 || Math.max(red, green, blue) - Math.min(red, green, blue) < 28) {
          continue
        }
        colorfulPixels += 1
        for (const [targetIndex, target] of targetColors.entries()) {
          const distance = Math.hypot(red - target.red, green - target.green, blue - target.blue)
          if (distance <= 72) {
            counts[targetIndex] = (counts[targetIndex] ?? 0) + 1
          }
        }
      }
      const best = counts.reduce((currentBest, count, index) => (count > currentBest.count ? { count, index } : currentBest), {
        count: 0,
        index: -1,
      })
      const minimumPixels = Math.max(12, Math.ceil(colorfulPixels * 0.5))
      return best.index >= 0 && best.count >= minimumPixels ? (targetColors[best.index]?.hex ?? null) : null
    },
    {
      dataUrl: `data:image/png;base64,${screenshot.toString('base64')}`,
      expectedColors: sameCorpusFillColorExpectedColors(),
    },
  )
}

function sameCorpusCellAddressCoordinates(address: string): { readonly columnIndex: number; readonly rowIndex: number } | null {
  const match = address
    .trim()
    .toUpperCase()
    .match(/^([A-Z]+)([0-9]+)$/u)
  if (!match) {
    return null
  }
  const columnLetters = match[1] ?? ''
  const rowIndex = Number(match[2]) - 1
  let columnIndex = 0
  for (const letter of columnLetters) {
    columnIndex = columnIndex * 26 + letter.charCodeAt(0) - 64
  }
  return columnIndex > 0 && rowIndex >= 0 ? { columnIndex: columnIndex - 1, rowIndex } : null
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
  locators: readonly SameCorpusFillableLocator[],
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

function readNullableNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function normalizeNullableText(value: string | null | undefined): string | null {
  const trimmed = value?.trim() ?? ''
  return trimmed.length > 0 ? trimmed : null
}

function primaryShortcut(key: string, platform: NodeJS.Platform = process.platform): string {
  return platform === 'darwin' ? `Meta+${key}` : `Control+${key}`
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
