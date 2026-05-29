import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-proof.ts'

export interface SameCorpusNameBoxReaderPage {
  readonly keyboard: {
    press(key: string): Promise<void>
  }
  locator(selector: string): SameCorpusReadableLocator
}

interface SameCorpusReadableLocator {
  first(): SameCorpusReadableLocator
  inputValue(options?: { readonly timeout?: number }): Promise<string>
}

export async function readSameCorpusVisibleSelectedRange(page: Page, product: UiResponsivenessSameCorpusProduct): Promise<string | null> {
  if (product === 'bilig') {
    const status = await page
      .locator('[data-testid="status-selection"]')
      .first()
      .textContent({ timeout: 1_000 })
      .catch(() => null)
    if (status && status.trim().length > 0) {
      return status.trim()
    }
    const nameBox = await page
      .locator('[data-testid="name-box"]')
      .first()
      .inputValue({ timeout: 1_000 })
      .catch(() => null)
    return nameBox && nameBox.trim().length > 0 ? nameBox.trim() : null
  }
  if (product === 'google-sheets') {
    const selectedRange = await readGoogleSheetsNameBoxSelection(page)
    if (selectedRange) {
      return selectedRange
    }
  }
  return await page.evaluate(() => {
    const selectors = [
      '#t-name-box input',
      'input[aria-label="Name box"]',
      '[aria-label="Name box"] input',
      '[aria-label="Name box"]',
      '[aria-label*="selected cell" i]',
      '[aria-label*="selection" i]',
    ]
    for (const selector of selectors) {
      const element = document.querySelector<HTMLInputElement | HTMLElement>(selector)
      const value =
        element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement
          ? element.value
          : (element?.getAttribute('aria-label') ?? element?.textContent ?? '')
      const trimmed = value.trim()
      if (trimmed.length > 0 && trimmed.length <= 64) {
        return trimmed
      }
    }
    const activeElement = document.activeElement
    const activeLabel = activeElement?.getAttribute('aria-label') ?? ''
    const activeLabelMatch = activeLabel.match(/(?:cell|range)\s+([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)/iu)
    return activeLabelMatch?.[1] ?? null
  })
}

export async function readSameCorpusVisibleSheetId(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  sheetName: string,
): Promise<string | null> {
  if (product === 'bilig') {
    const sheetIdentity = await page.evaluate((name) => window.__biligSameCorpusProof?.readSheetIdentity?.(name) ?? null, sheetName)
    return typeof sheetIdentity?.sheetId === 'number' && Number.isSafeInteger(sheetIdentity.sheetId) ? String(sheetIdentity.sheetId) : null
  }
  if (product === 'google-sheets') {
    return googleSheetsVisibleSheetId(page.url())
  }
  return null
}

export async function readGoogleSheetsNameBoxSelection(page: SameCorpusNameBoxReaderPage): Promise<string | null> {
  await page.keyboard.press(primaryShortcut('J')).catch(() => null)
  const value = await page
    .locator('#t-name-box')
    .first()
    .inputValue({ timeout: 1_500 })
    .catch(() => null)
  await page.keyboard.press('Escape').catch(() => null)
  return value && value.trim().length > 0 ? value.trim() : null
}

function googleSheetsVisibleSheetId(url: string): string | null {
  try {
    const parsed = new URL(url)
    const searchGid = parsed.searchParams.get('gid')?.trim()
    const hashGid = parsed.hash.match(/(?:^#|[&#])gid=([^&#]+)/u)?.[1]?.trim()
    const gid = searchGid || hashGid
    return gid && /^\d+$/u.test(gid) ? `gid:${gid}` : null
  } catch {
    return null
  }
}

function primaryShortcut(key: string, platform: NodeJS.Platform = process.platform): string {
  return platform === 'darwin' ? `Meta+${key}` : `Control+${key}`
}
