import { expect, test, type Page } from '@playwright/test'
import {
  PRIMARY_MODIFIER,
  clickProductCell,
  countDarkReadbackPixelsInCell,
  createTestDocumentId,
  installTypeGpuCellReadbackHarness,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('@browser-webgpu @browser-deep web app keeps an in-cell Delete clear committed after clicking away', async ({ page }) => {
  const staleText = 'editor-delete-clickaway'
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(createTestDocumentId('playwright-editor-delete-click-away'))}&persist=0`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await formulaInput.fill(staleText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(staleText)
  await expectCellRenderedText(page, 1, 1, staleText, 'visible')

  await page.getByTestId('sheet-grid-focus-target').focus()
  await page.keyboard.press('F2')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue(staleText)

  await page.keyboard.press(`${PRIMARY_MODIFIER}+A`)
  await page.keyboard.press('Delete')
  await expect(cellEditor).toHaveValue('')

  await clickProductCell(page, 2, 1)
  await expect(cellEditor).toHaveCount(0)
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 1, 1, staleText, 'hidden')

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await expect(formulaInput).toHaveValue('')
  await expectCellRenderedText(page, 1, 1, staleText, 'hidden')
})

test('@browser-webgpu @browser-deep web app keeps active in-cell undo and redo local to the draft editor', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-editor-local-undo-redo')
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=B2`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')
  const cellEditor = page.getByTestId('cell-editor-input')
  const redoShortcut = PRIMARY_MODIFIER === 'Meta' ? `${PRIMARY_MODIFIER}+Shift+Z` : `${PRIMARY_MODIFIER}+Y`

  await clickProductCell(page, 3, 3)
  await formulaInput.fill('workbook-history-sentinel')
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue('workbook-history-sentinel')

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await page.keyboard.press('a')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('a')
  await cellEditor.press('b')
  await expect(cellEditor).toHaveValue('ab')
  await cellEditor.press('c')
  await expect(cellEditor).toHaveValue('abc')

  await cellEditor.press(`${PRIMARY_MODIFIER}+Z`)
  await expect(cellEditor).toHaveValue('ab')
  await expect(formulaInput).toHaveValue('ab')

  await cellEditor.press(`${PRIMARY_MODIFIER}+Z`)
  await expect(cellEditor).toHaveValue('a')
  await expect(formulaInput).toHaveValue('a')

  await cellEditor.press(redoShortcut)
  await expect(cellEditor).toHaveValue('ab')
  await expect(formulaInput).toHaveValue('ab')

  await page.keyboard.press('Enter')
  await expect(cellEditor).toHaveCount(0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B3')
  await expect.poll(() => nativeTextRunsInclude(page, 'abc')).toBe(false)
  await expectCellRenderedText(page, 1, 1, 'ab', 'visible')
  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('ab')
  await expect.poll(() => nativeTextRunsInclude(page, 'ab')).toBe(true)
  await clickProductCell(page, 3, 3)
  await expect(formulaInput).toHaveValue('workbook-history-sentinel')
})

async function nativeTextRunsInclude(page: Page, text: string): Promise<boolean> {
  return await page.evaluate((needle) => {
    const nativeTextIncludes = Array.from(document.querySelectorAll('[data-native-text-run]')).some(
      (run) => run.textContent?.includes(needle) ?? false,
    )
    if (nativeTextIncludes) {
      return true
    }
    const typeGpu = document.querySelector('[data-testid="grid-pane-renderer"]')
    if (!(typeGpu instanceof HTMLElement) || typeGpu.getAttribute('data-v3-frame-proof-status') !== 'presented') {
      return false
    }
    if (Number(typeGpu.getAttribute('data-v3-presented-text-run-count') ?? '0') <= 0) {
      return false
    }
    const formulaInput = document.querySelector('[data-testid="formula-input"]')
    const selectedValue = formulaInput instanceof HTMLInputElement || formulaInput instanceof HTMLTextAreaElement ? formulaInput.value : ''
    const resolvedValue = document.querySelector('[data-testid="formula-resolved-value"]')?.textContent ?? ''
    return selectedValue.includes(needle) || resolvedValue.includes(needle)
  }, text)
}

async function expectCellRenderedText(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  text: string,
  expected: 'hidden' | 'visible',
): Promise<void> {
  const hasTypeGpuCanvas = (await page.getByTestId('grid-pane-renderer').count()) > 0
  if (!hasTypeGpuCanvas) {
    await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(expected === 'visible')
    return
  }
  const darkPixelPoll = expect.poll(async () => await countDarkReadbackPixelsInCell(page, columnIndex, rowIndex), {
    message: `cell ${columnIndex}:${rowIndex} rendered text should be ${expected}`,
  })
  if (expected === 'visible') {
    await darkPixelPoll.toBeGreaterThan(4)
    return
  }
  await darkPixelPoll.toBeLessThanOrEqual(2)
}
