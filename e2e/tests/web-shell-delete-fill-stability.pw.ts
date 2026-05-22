import { expect, test, type Page } from '@playwright/test'
import {
  clickProductCell,
  countGreenFillReadbackPixelsInCell,
  createTestDocumentId,
  expectToolbarColor,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  getToolbarButton,
  installTypeGpuCellReadbackHarness,
  PRIMARY_MODIFIER,
  PRODUCT_HEADER_HEIGHT,
  pickToolbarPresetColor,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

test('@browser-webgpu @browser-deep web app keeps deleted filled cells stable after click-away and viewport churn', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-delete-fill-stability')
  const text = 'delete-fill-no-ghost'
  await page.setViewportSize({ width: 1166, height: 820 })
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(text)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(text)

  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expect
    .poll(() => countGreenFillReadbackPixelsInCell(page, 1, 1), {
      message: 'setup should visibly paint B2 green before deletion',
      timeout: 5_000,
    })
    .toBeGreaterThan(120)

  await clickProductCell(page, 1, 1)
  await grid.press('Delete')
  await expect(formulaInput).toHaveValue('')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4))),
    'delete must clear text without flashing the retained fill to default',
  ).toBeGreaterThan(120)

  await clickProductCell(page, 4, 4)
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 1, 1)).toBeGreaterThan(120)

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 900
    viewport.scrollLeft = 220
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 1, 1)).toBe(0)

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    viewport.scrollTop = 0
    viewport.scrollLeft = 0
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }))
  })
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 1, 1)).toBeGreaterThan(120)

  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('')
})

test('@browser-webgpu @browser-deep web app preserves filled-cell presentation when formula-bar clear commits as delete', async ({
  page,
}) => {
  const documentId = createTestDocumentId('playwright-formula-clear-fill-stability')
  const text = 'formula-clear-keeps-fill'
  await page.setViewportSize({ width: 1166, height: 820 })
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 3, 9)
  await formulaInput.fill(text)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(text)

  await clickProductCell(page, 3, 9)
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await waitForGreenFillPaint(page, 3, 9, 'setup should visibly paint D10 green before formula-bar clear')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 3, 9, 4))),
    'setup should visibly paint D10 green before formula-bar clear',
  ).toBeGreaterThan(120)

  await formulaInput.fill('')
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue('')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 3, 9, 4))),
    'formula-bar clear should not flash the retained fill to default while commit delete is optimistic',
  ).toBeGreaterThan(120)

  await clickProductCell(page, 5, 11)
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 3, 9)).toBeGreaterThan(120)

  await clickProductCell(page, 3, 9)
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 3, 9)).toBeGreaterThan(120)
})

test('@browser-webgpu @browser-deep web app applies fill color after moving text into an empty tile range', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-move-text-fill-range')
  const text = 'moved-fill-stability'
  await page.setViewportSize({ width: 1166, height: 820 })
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(text)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(text)

  await dragProductSelectionBorder(page, 1, 1, 3, 4)
  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('')
  await clickProductCell(page, 3, 4)
  await expect(formulaInput).toHaveValue(text)

  await clickProductCell(page, 3, 4)
  await clickProductCell(page, 5, 7, { shift: true })
  await pickToolbarPresetColor(page, 'Fill color', 'green')

  const fillProofCells = [
    [3, 4],
    [4, 4],
    [5, 4],
    [3, 7],
    [4, 7],
    [5, 7],
  ] as const
  await expect
    .poll(
      async () => {
        const pixelCounts = await Promise.all(
          fillProofCells.map(([columnIndex, rowIndex]) => countGreenFillReadbackPixelsInCell(page, columnIndex, rowIndex)),
        )
        return Math.min(...pixelCounts)
      },
      {
        message: 'moved text range should paint green fill across visible occupied and empty cells',
        timeout: 5_000,
      },
    )
    .toBeGreaterThan(120)

  await grid.press('Delete')
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 3, 4)).toBeGreaterThan(120)

  await clickProductCell(page, 7, 9)
  await clickProductCell(page, 3, 4)
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 3, 4)).toBeGreaterThan(120)
})

test('@browser-webgpu @browser-deep web app keeps fill undo and redo visually stable from grid keyboard ownership', async ({ page }) => {
  const documentId = createTestDocumentId('playwright-fill-undo-redo-stability')
  const redoShortcut = PRIMARY_MODIFIER === 'Meta' ? 'Meta+Shift+Z' : 'Control+Y'
  await page.setViewportSize({ width: 1166, height: 820 })
  await installTypeGpuCellReadbackHarness(page)
  await page.goto(`/?document=${encodeURIComponent(documentId)}&persist=0&sheet=Sheet1&cell=A1`)
  await waitForWorkbookReady(page)

  const grid = page.getByTestId('sheet-grid')
  const formulaInput = page.getByTestId('formula-input')

  await clickProductCell(page, 1, 1)
  await pickToolbarPresetColor(page, 'Fill color', 'green')
  await expectToolbarColor(getToolbarButton(page, 'Fill color'), '#00ff00')
  await waitForGreenFillPaint(page, 1, 1, 'setup should paint B2 green before undoing the style mutation')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4))),
    'setup should paint B2 green before undoing the style mutation',
  ).toBeGreaterThan(120)

  await expect
    .poll(async () => page.evaluate(() => document.activeElement?.getAttribute('data-testid') ?? null), {
      message: 'toolbar style command must return keyboard ownership to the grid before undo',
    })
    .toBe('sheet-grid-focus-target')

  await grid.press(`${PRIMARY_MODIFIER}+Z`)
  await expect(formulaInput).toHaveValue('')
  await expect.poll(() => countGreenFillReadbackPixelsInCell(page, 1, 1)).toBe(0)
  await expectToolbarColor(getToolbarButton(page, 'Fill color'), '#ffffff')
  await expect(page.getByTestId('name-box')).toHaveValue('B2')

  await grid.press(redoShortcut)
  await expect(formulaInput).toHaveValue('')
  await expectToolbarColor(getToolbarButton(page, 'Fill color'), '#00ff00')
  expect(
    Math.min(...(await sampleGreenFillPixelsAcrossFrames(page, 1, 1, 4))),
    'redo should restore the green fill without flashing through default grid paint',
  ).toBeGreaterThan(120)
  await expect(page.getByTestId('name-box')).toHaveValue('B2')
})

async function waitForGreenFillPaint(page: Page, columnIndex: number, rowIndex: number, message: string) {
  await expect
    .poll(() => countGreenFillReadbackPixelsInCell(page, columnIndex, rowIndex), {
      message,
      timeout: 5_000,
    })
    .toBeGreaterThan(120)
}

async function sampleGreenFillPixelsAcrossFrames(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  remainingSamples: number,
  samples: readonly number[] = [],
): Promise<readonly number[]> {
  if (remainingSamples <= 0) {
    return samples
  }
  const pixels = await countGreenFillReadbackPixelsInCell(page, columnIndex, rowIndex)
  await page.waitForTimeout(50)
  return await sampleGreenFillPixelsAcrossFrames(page, columnIndex, rowIndex, remainingSamples - 1, [...samples, pixels])
}

async function dragProductSelectionBorder(page: Page, startColumn: number, startRow: number, targetColumn: number, targetRow: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const startTop = await getProductRowTop(page, startRow)
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetTop = await getProductRowTop(page, targetRow)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetHeight = await getProductRowHeight(page, targetRow)

  await page.mouse.move(grid.x + startLeft + 3, grid.y + PRODUCT_HEADER_HEIGHT + startTop + 2)
  await page.mouse.down()
  await page.mouse.move(
    grid.x + targetLeft + Math.floor(targetWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + targetTop + Math.floor(targetHeight / 2),
    { steps: 12 },
  )
  await page.mouse.up()
}
