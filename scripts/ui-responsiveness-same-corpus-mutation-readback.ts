import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type { SameCorpusMutationTargetSelection } from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import { sameCorpusFillColorExpectedColors } from './ui-responsiveness-same-corpus-workload-runner.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

interface ScreenshotBuffer {
  toString(encoding: 'base64'): string
}

interface SameCorpusViewportSize {
  readonly width: number
  readonly height: number
}

export interface SameCorpusVisibleCellInteriorClip {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export async function readExternalVisibleGridCellReadback(args: {
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly target: SameCorpusMutationTargetSelection
  readonly workload?: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetReadback | null> {
  const selectedTargetBox = await firstVisibleTargetBox(args.page, args.product)
  if (!selectedTargetBox) {
    return null
  }
  const targetCellReadback = await args.page
    .evaluate(readSameCorpusVisibleTargetCellReadbackFromPage, {
      targetBox: selectedTargetBox,
      targetRange: args.target.targetRange,
    })
    .catch(
      (): SameCorpusMutationTargetReadback => ({
        fillColor: null,
        formula: null,
        source: 'unknown',
        value: null,
        visibleText: null,
      }),
    )
  const screenshotFillColor =
    args.workload === 'fill-format-change' ? await readExternalVisibleGridCellFillColor(args.page, selectedTargetBox) : null
  if (targetCellReadback.source !== 'visible-grid-cell' && screenshotFillColor === null) {
    return null
  }
  if (args.workload !== 'fill-format-change' && targetCellReadback.visibleText === null) {
    return null
  }
  return {
    ...targetCellReadback,
    fillColor: screenshotFillColor,
    source: 'visible-grid-cell',
  }
}

export function readSameCorpusVisibleTargetCellReadbackFromPage(args: {
  readonly targetBox: { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
  readonly targetRange?: string
}): SameCorpusMutationTargetReadback {
  const targetBox = args.targetBox
  const candidates = targetCellCandidates()
    .map((element) => ({ element, score: targetOverlapScore(element) }))
    .filter((candidate) => candidate.score > 0)
    .toSorted((left, right) => right.score - left.score)
  const targetText = normalizeText(
    candidates.map((candidate) => textFromElement(candidate.element, args.targetRange)).find((candidateText) => candidateText !== null),
  )
  const fillColor =
    candidates.map((candidate) => visibleBackgroundColor(candidate.element)).find((color): color is string => color !== null) ?? null
  const source = candidates.length > 0 ? 'visible-grid-cell' : 'unknown'
  return {
    value: targetText && !targetText.startsWith('=') ? targetText : null,
    formula: targetText?.startsWith('=') ? targetText : null,
    fillColor,
    visibleText: targetText,
    source,
  }

  function targetCellCandidates(): HTMLElement[] {
    const seen = new Set<HTMLElement>()
    const add = (element: Element | null | undefined): void => {
      if (!(element instanceof HTMLElement)) {
        return
      }
      const targetCell = nearestGridCellCandidate(element)
      if (!targetCell || seen.has(targetCell) || isExcludedChromeElement(targetCell) || isTransientEditorElement(targetCell)) {
        return
      }
      seen.add(targetCell)
    }
    for (const selector of [
      '.waffle-cell',
      '.waffle-active-cell',
      '[role="gridcell"]',
      '[aria-selected="true"]',
      '.active-cell',
      '.selected-cell',
    ]) {
      for (const element of Array.from(document.querySelectorAll(selector))) {
        add(element)
      }
    }
    const points = [
      { x: targetBox.x + targetBox.width / 2, y: targetBox.y + targetBox.height / 2 },
      { x: targetBox.x + 3, y: targetBox.y + 3 },
      { x: targetBox.x + targetBox.width - 3, y: targetBox.y + 3 },
      { x: targetBox.x + 3, y: targetBox.y + targetBox.height - 3 },
      { x: targetBox.x + targetBox.width - 3, y: targetBox.y + targetBox.height - 3 },
    ]
    if (typeof document.elementsFromPoint === 'function') {
      for (const point of points) {
        for (const element of document.elementsFromPoint(point.x, point.y)) {
          add(element)
        }
      }
    }
    return [...seen]
  }

  function targetOverlapScore(element: HTMLElement): number {
    const rect = element.getBoundingClientRect()
    if (rect.width <= 0 || rect.height <= 0) {
      return 0
    }
    const cellLike = isGridCellCandidate(element)
    if (!cellLike) {
      return 0
    }
    if (!hasTargetCellLikeGeometry(rect, targetBox)) {
      return 0
    }
    const overlapWidth = Math.max(0, Math.min(rect.right, targetBox.x + targetBox.width) - Math.max(rect.left, targetBox.x))
    const overlapHeight = Math.max(0, Math.min(rect.bottom, targetBox.y + targetBox.height) - Math.max(rect.top, targetBox.y))
    const overlapArea = overlapWidth * overlapHeight
    if (overlapArea <= 0) {
      return 0
    }
    const smallerArea = Math.max(1, Math.min(rect.width * rect.height, targetBox.width * targetBox.height))
    const overlapRatio = overlapArea / smallerArea
    const roleBoost = cellLike ? 20 : 0
    const textBoost = textFromElement(element, args.targetRange) !== null ? 5 : 0
    return overlapRatio * 100 + roleBoost + textBoost
  }

  function textFromElement(element: HTMLElement, targetRange: string | undefined): string | null {
    if (element instanceof HTMLScriptElement || element instanceof HTMLStyleElement || element instanceof HTMLLinkElement) {
      return null
    }
    if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
      return null
    }
    return normalizeText(element.textContent ?? '') ?? visibleTextFromAriaLabel(element, targetRange)
  }

  function visibleBackgroundColor(element: HTMLElement): string | null {
    for (const candidate of [element, ...Array.from(element.querySelectorAll<HTMLElement>('*'))]) {
      if (!elementPaintsTargetInterior(candidate)) {
        continue
      }
      const color = normalizeBackgroundColor(getComputedStyle(candidate).backgroundColor)
      if (color !== null && !isKnownSelectionChromeBackground(candidate, color)) {
        return color
      }
    }
    return null
  }

  // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright serializes this helper with the page function.
  function normalizeText(value: string | null | undefined): string | null {
    const trimmedText =
      value
        ?.replace(/[\u200b-\u200f\ufeff]/gu, '')
        .replace(/\s+/gu, ' ')
        .trim() ?? ''
    return trimmedText.length > 0 && trimmedText.length <= 512 ? trimmedText : null
  }

  // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright serializes this helper with the page function.
  function normalizeBackgroundColor(value: string | null | undefined): string | null {
    const color = value?.trim() ?? ''
    if (!color || color === 'transparent' || color === 'rgba(0, 0, 0, 0)' || color === 'rgb(255, 255, 255)') {
      return null
    }
    return color
  }

  // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright serializes this helper with the page function.
  function isExcludedChromeElement(element: HTMLElement): boolean {
    const className = element.className.toString().toLowerCase()
    const ariaLabel = element.getAttribute('aria-label')?.toLowerCase() ?? ''
    if (
      /\b(?:range-border|waffle-border-cell|fill-handle|autofill-handle)\b/u.test(className) ||
      /\b(?:fill handle|autofill handle|selection border)\b/u.test(ariaLabel)
    ) {
      return true
    }
    return Boolean(
      element.closest(
        '#t-formula-bar-input, #t-formula-bar, #t-name-box, input.waffle-name-box, [aria-label="Formula bar"], [aria-label="Name box"], .docs-toolbar, .waffle-menu, .range-border, .waffle-border-cell-active',
      ),
    )
  }

  // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright serializes this helper with the page function.
  function isTransientEditorElement(element: HTMLElement): boolean {
    const className = element.className.toString().toLowerCase()
    return (
      element instanceof HTMLInputElement ||
      element instanceof HTMLTextAreaElement ||
      element.isContentEditable ||
      element.matches('[contenteditable="true"]') ||
      element.querySelector('input, textarea, [contenteditable="true"]') !== null ||
      /\b(?:waffle-cell-input|cell-input|formula-input)\b/u.test(className)
    )
  }

  function nearestGridCellCandidate(element: HTMLElement): HTMLElement | null {
    for (let candidate: HTMLElement | null = element; candidate; candidate = candidate.parentElement) {
      if (isExcludedChromeElement(candidate) || isTransientEditorElement(candidate)) {
        return null
      }
      if (isGridCellCandidate(candidate)) {
        return candidate
      }
    }
    return null
  }

  // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright serializes this helper with the page function.
  function isGridCellCandidate(element: HTMLElement): boolean {
    const classList = element.classList
    return (
      element.getAttribute('role') === 'gridcell' ||
      classList.contains('waffle-cell') ||
      classList.contains('waffle-active-cell') ||
      classList.contains('active-cell') ||
      classList.contains('selected-cell')
    )
  }

  function elementPaintsTargetInterior(element: HTMLElement): boolean {
    const rect = element.getBoundingClientRect()
    if (rect.width <= 0 || rect.height <= 0) {
      return false
    }
    const insetX = Math.min(Math.max(2, targetBox.width * 0.15), Math.max(2, targetBox.width / 2 - 1))
    const insetY = Math.min(Math.max(2, targetBox.height * 0.15), Math.max(2, targetBox.height / 2 - 1))
    const interiorLeft = targetBox.x + insetX
    const interiorTop = targetBox.y + insetY
    const interiorRight = targetBox.x + targetBox.width - insetX
    const interiorBottom = targetBox.y + targetBox.height - insetY
    const overlapWidth = Math.max(0, Math.min(rect.right, interiorRight) - Math.max(rect.left, interiorLeft))
    const overlapHeight = Math.max(0, Math.min(rect.bottom, interiorBottom) - Math.max(rect.top, interiorTop))
    const interiorArea = Math.max(1, (interiorRight - interiorLeft) * (interiorBottom - interiorTop))
    return (overlapWidth * overlapHeight) / interiorArea >= 0.5
  }

  function visibleTextFromAriaLabel(element: HTMLElement, targetRange: string | undefined): string | null {
    const label = normalizeText(element.getAttribute('aria-label'))
    if (!label) {
      return null
    }
    const normalizedTargetRange = targetRange?.replace(/\$/gu, '').trim().toUpperCase() ?? ''
    if (normalizedTargetRange && label.toUpperCase() === normalizedTargetRange) {
      return null
    }
    if (normalizedTargetRange && /^[A-Z]+[0-9]+$/u.test(normalizedTargetRange)) {
      const prefixPattern = new RegExp(`^${escapeRegExp(normalizedTargetRange)}(?:\\s+|\\s*[:,-]\\s*)(.+)$`, 'iu')
      const match = label.match(prefixPattern)
      const visibleText = normalizeText(match?.[1])
      if (visibleText) {
        return visibleText
      }
    }
    return label
  }

  // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright serializes this helper with the page function.
  function escapeRegExp(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  }

  // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright serializes this helper with the page function.
  function isKnownSelectionChromeBackground(element: HTMLElement, color: string): boolean {
    const normalizedColor = color.replace(/\s+/gu, '').toLowerCase()
    if (
      normalizedColor !== 'rgb(11,87,208)' &&
      normalizedColor !== 'rgba(11,87,208,1)' &&
      normalizedColor !== 'rgb(26,115,232)' &&
      normalizedColor !== 'rgba(26,115,232,1)'
    ) {
      return false
    }
    const className = element.className.toString().toLowerCase()
    return (
      element.getAttribute('aria-selected') === 'true' ||
      /\b(?:active|selected|selection|range-border|waffle-border-cell)\b/u.test(className) ||
      element.closest('.range-border, .waffle-border-cell-active, [aria-selected="true"]') !== null
    )
  }
}

function hasTargetCellLikeGeometry(
  rect: DOMRect,
  expected: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
): boolean {
  const minimumWidth = Math.max(4, expected.width * 0.35)
  const minimumHeight = Math.max(4, expected.height * 0.35)
  const maximumWidth = Math.max(expected.width + 8, expected.width * 1.35)
  const maximumHeight = Math.max(expected.height + 8, expected.height * 1.35)
  return rect.width >= minimumWidth && rect.height >= minimumHeight && rect.width <= maximumWidth && rect.height <= maximumHeight
}

export async function readExpectedFillColorFromScreenshot(page: Page, screenshot: ScreenshotBuffer): Promise<string | null> {
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
        let nearestTargetIndex = -1
        let nearestTargetDistance = Number.POSITIVE_INFINITY
        for (const [targetIndex, target] of targetColors.entries()) {
          const distance = Math.hypot(red - target.red, green - target.green, blue - target.blue)
          if (distance < nearestTargetDistance) {
            nearestTargetDistance = distance
            nearestTargetIndex = targetIndex
          }
        }
        if (nearestTargetIndex >= 0 && nearestTargetDistance <= 72) {
          counts[nearestTargetIndex] = (counts[nearestTargetIndex] ?? 0) + 1
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

export function sameCorpusCellAddressCoordinates(address: string): { readonly columnIndex: number; readonly rowIndex: number } | null {
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

export async function readVisibleFormulaBarText(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<string | null> {
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

async function readExternalVisibleGridCellFillColor(
  page: Page,
  targetBox: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
): Promise<string | null> {
  const viewport = page.viewportSize() ?? (await page.evaluate(() => ({ height: window.innerHeight, width: window.innerWidth })))
  const clip = sameCorpusVisibleCellInteriorClip(targetBox, viewport)
  if (!clip) {
    return null
  }
  const screenshot = await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip,
  })
  return await readExpectedFillColorFromScreenshot(page, screenshot)
}

export function sameCorpusVisibleCellInteriorClip(
  targetBox: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  viewport: SameCorpusViewportSize,
): SameCorpusVisibleCellInteriorClip | null {
  const xInset = sameCorpusVisibleCellInteriorInset(targetBox.width, 3)
  const yInset = sameCorpusVisibleCellInteriorInset(targetBox.height, 3)
  const x0 = Math.max(0, Math.floor(targetBox.x + xInset))
  const y0 = Math.max(0, Math.floor(targetBox.y + yInset))
  const x1 = Math.min(viewport.width, Math.ceil(targetBox.x + targetBox.width - xInset))
  const y1 = Math.min(viewport.height, Math.ceil(targetBox.y + targetBox.height - yInset))
  if (x1 <= x0 || y1 <= y0) {
    return null
  }
  return { height: y1 - y0, width: x1 - x0, x: x0, y: y0 }
}

function sameCorpusVisibleCellInteriorInset(size: number, minimumInset: number): number {
  if (size <= 2) {
    return 0
  }
  return Math.min(Math.max(minimumInset, Math.floor(size * 0.2)), Math.floor((size - 1) / 2))
}

export async function firstVisibleTargetBox(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
): Promise<{ readonly x: number; readonly y: number; readonly width: number; readonly height: number } | null> {
  if (product === 'google-sheets') {
    const activeCellBorderBox = await googleSheetsActiveCellBorderBox(page).catch(() => null)
    if (activeCellBorderBox) {
      return activeCellBorderBox
    }
  }
  const selectors =
    product === 'google-sheets'
      ? ['.waffle-active-cell', '.waffle-border-cell-active', '[class*="active-cell" i]', '[class*="selected-cell" i]']
      : ['.ewr-selection', '.ewr-active-cell', '[class*="active-cell" i]', '[class*="selected-cell" i]']
  const frames = page.frames()
  const targetBoxPromises = selectors.flatMap((selector) =>
    [visibleLocatorBox(page.locator(selector).first())].concat(frames.map((frame) => visibleLocatorBox(frame.locator(selector).first()))),
  )
  const boxes = await Promise.all(targetBoxPromises)
  return boxes.find((box): box is NonNullable<(typeof boxes)[number]> => box !== null) ?? null
}

async function googleSheetsActiveCellBorderBox(
  page: Page,
): Promise<{ readonly x: number; readonly y: number; readonly width: number; readonly height: number } | null> {
  return await page.evaluate(() => {
    const borders = Array.from(document.querySelectorAll<HTMLElement>('.range-border.active-cell-border')).map((element) =>
      element.getBoundingClientRect(),
    )
    const visibleBorders = borders.filter((rect) => rect.width > 0 && rect.height > 0)
    if (visibleBorders.length < 4) {
      return null
    }
    const left = Math.min(...visibleBorders.map((rect) => rect.left))
    const top = Math.min(...visibleBorders.map((rect) => rect.top))
    const right = Math.max(...visibleBorders.map((rect) => rect.right))
    const bottom = Math.max(...visibleBorders.map((rect) => rect.bottom))
    if (right - left <= 2 || bottom - top <= 2) {
      return null
    }
    return {
      x: Math.max(0, Math.floor(left)),
      y: Math.max(0, Math.floor(top)),
      width: Math.max(1, Math.ceil(right - left)),
      height: Math.max(1, Math.ceil(bottom - top)),
    }
  })
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
