import { createHash } from 'node:crypto'

import type { Page } from '@playwright/test'
import { formatCellDisplayValue, isCellValue } from '@bilig/protocol'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import {
  readSameCorpusVisibleSheetId,
  readSameCorpusVisibleSelectedRange,
  type SameCorpusMutationTargetReadback,
  type SameCorpusMutationTargetScreenshotProof,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'
import { readBiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface-page.ts'
import { settleFrames } from './ui-responsiveness-same-corpus-page-utils.ts'
import { sameCorpusMutationTargetRangeForSample } from './ui-responsiveness-same-corpus-mutation-target-spec.ts'
import {
  firstVisibleTargetBox,
  readExpectedFillColorFromScreenshot,
  readExternalVisibleGridCellReadback,
  readVisibleFormulaBarText,
  sameCorpusCellAddressCoordinates,
  sameCorpusVisibleCellInteriorClip,
} from './ui-responsiveness-same-corpus-mutation-readback.ts'
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

export { sameCorpusMutationTargetRangeForSample } from './ui-responsiveness-same-corpus-mutation-target-spec.ts'
export { readSameCorpusVisibleTargetCellReadbackFromPage } from './ui-responsiveness-same-corpus-mutation-readback.ts'

export async function readSameCorpusDeclaredMutationTargetSelection(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly sheetName: string
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetSelection> {
  const sheetId = await readSameCorpusVisibleSheetId(args.page, args.product, args.sheetName)
  return normalizeSameCorpusMutationTargetSelection(
    sameCorpusMutationTargetRangeForSample(args.workload, args.sampleIndex),
    args.sheetName,
    sheetId,
  )
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
    await assertBiligNameBoxSelectionCommitted(args.page, args.target)
    await focusBiligGridForKeyboardMutation(args.page)
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

async function assertBiligNameBoxSelectionCommitted(page: Page, target: SameCorpusMutationTargetSelection): Promise<void> {
  await page.waitForFunction(
    ({ expectedRange, expectedStatus }) => {
      const nameBox = document.querySelector<HTMLInputElement>('[data-testid="name-box"]')
      const status = document.querySelector('[data-testid="status-selection"]')
      const normalizedNameBox = nameBox?.value.replace(/\$/gu, '').trim().toUpperCase() ?? ''
      const normalizedStatus = status?.textContent?.replace(/\$/gu, '').trim().toUpperCase() ?? ''
      return normalizedNameBox === expectedRange || normalizedStatus === expectedStatus
    },
    {
      expectedRange: target.targetRange.toUpperCase(),
      expectedStatus: `${target.sheetName}!${target.targetRange}`.toUpperCase(),
    },
    { timeout: 3_000 },
  )
}

async function focusBiligGridForKeyboardMutation(page: Page): Promise<void> {
  const focusTarget = page.getByTestId('sheet-grid-focus-target')
  await focusTarget.focus()
  await page.waitForFunction(() => document.activeElement?.getAttribute('data-testid') === 'sheet-grid-focus-target', undefined, {
    timeout: 1_500,
  })
}

export async function readSameCorpusMutationTargetReadback(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly target: SameCorpusMutationTargetSelection
  readonly workload?: UiResponsivenessSameCorpusMutatingWorkload
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
  if (args.product === 'bilig' && args.target) {
    return await readBiligVisibleGridCellReadback(args.page, args.target)
  }
  if (args.product !== 'bilig' && args.target) {
    const targetCellReadback = await readExternalVisibleGridCellReadback({
      page: args.page,
      product: args.product,
      target: args.target,
      workload: args.workload,
    })
    if (targetCellReadback) {
      return targetCellReadback
    }
    return missingSameCorpusVisibleTargetCellReadback()
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

function missingSameCorpusVisibleTargetCellReadback(): SameCorpusMutationTargetReadback {
  return {
    value: null,
    formula: null,
    fillColor: null,
    visibleText: null,
    source: 'unknown',
  }
}

export async function captureSameCorpusMutationTargetScreenshotProof(args: {
  readonly page: Page
  readonly sampleIndex: number
  readonly product: UiResponsivenessSameCorpusProduct
  readonly semanticReadback: SameCorpusMutationTargetReadback
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
    semanticReadback: args.semanticReadback,
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
    const viewport = page.viewportSize() ?? (await page.evaluate(() => ({ height: window.innerHeight, width: window.innerWidth })))
    const clip = sameCorpusVisibleCellInteriorClip(selectedTargetBox, viewport)
    if (!clip) {
      return { buffer: null, captured: false, scope: 'visible-grid-fallback' }
    }
    const buffer = await page.screenshot({
      animations: 'disabled',
      caret: 'hide',
      clip,
      path: screenshotPath,
    })
    return { buffer, captured: true, scope: 'target-cell' }
  }
  return { buffer: null, captured: false, scope: 'visible-grid-fallback' }
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
  const value = normalizeUnknownCellDisplayValue(cell.value ?? cell.input)
  const fillColor = normalizeNullableText(readCellFillColor(cell.style))
  return {
    batchId: readNullableNumber(range.batchId),
    capturedRevision: normalizeNullableText(normalizeUnknownCellDisplayValue(range.capturedRevision)),
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

async function readBiligVisibleGridCellReadback(
  page: Page,
  target: SameCorpusMutationTargetSelection,
): Promise<SameCorpusMutationTargetReadback> {
  const [visibleText, fillColor, visibleSceneProofSha256] = await Promise.all([
    readBiligVisibleGridCellText(page, target),
    readBiligVisibleGridCellFillColor(page, target),
    readBiligVisibleSceneProofSha256(page, target),
  ])
  const text = normalizeNullableText(visibleText)
  return {
    value: text && !text.startsWith('=') ? text : null,
    formula: text?.startsWith('=') ? text : null,
    fillColor,
    visibleText: text,
    source: 'visible-grid-cell',
    visibleSceneProofSha256,
  }
}

async function readBiligVisibleSceneProofSha256(page: Page, target: SameCorpusMutationTargetSelection): Promise<string | null> {
  const visibleSceneProof = await page.evaluate(
    ({ endAddress, sheetName, startAddress }) =>
      window.__biligSameCorpusProof?.readRange(sheetName, startAddress, endAddress)?.visibleSceneProof ?? null,
    {
      endAddress: target.endAddress,
      sheetName: target.sheetName,
      startAddress: target.startAddress,
    },
  )
  return visibleSceneProof ? sha256Hex(stableJsonBytes(visibleSceneProof)) : null
}

async function readBiligVisibleGridCellText(page: Page, target: SameCorpusMutationTargetSelection): Promise<string | null> {
  const address = sameCorpusCellAddressCoordinates(target.startAddress)
  if (!address) {
    return null
  }
  return await page.evaluate(({ columnIndex, rowIndex }) => {
    const layer = document.querySelector('[data-testid="grid-native-text-layer"]')
    if (!(layer instanceof HTMLElement)) {
      return null
    }
    const runs = Array.from(layer.querySelectorAll<HTMLElement>('[data-native-text-run]')).filter(
      (run) =>
        run.getAttribute('data-native-text-run-col') === String(columnIndex) &&
        run.getAttribute('data-native-text-run-row') === String(rowIndex),
    )
    const text = runs
      .map((run) => run.textContent ?? '')
      .join('')
      .trim()
    return text.length > 0 ? text : null
  }, address)
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

function normalizeUnknownCellDisplayValue(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null
  }
  if (isCellValue(value)) {
    return normalizeNullableText(formatCellDisplayValue(value, undefined))
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
