import { expect, test, type Page } from '@playwright/test'
import {
  createTestDocumentId,
  dragProductBodySelection,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductFillHandleDragPoints,
  getProductRowHeight,
  getProductRowTop,
  gotoWorkbookShell,
  PRODUCT_HEADER_HEIGHT,
  waitForWorkbookReady,
} from './web-shell-helpers.js'

async function getProductCellRangeBox(
  page: Page,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): Promise<{
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const leftColumn = Math.min(startColumn, endColumn)
  const rightColumn = Math.max(startColumn, endColumn)
  const topRow = Math.min(startRow, endRow)
  const bottomRow = Math.max(startRow, endRow)
  const left = await getProductColumnLeft(page, leftColumn)
  const right = await getProductColumnLeft(page, rightColumn)
  const rightWidth = await getProductColumnWidth(page, rightColumn)
  const top = await getProductRowTop(page, topRow)
  const bottom = await getProductRowTop(page, bottomRow)
  const bottomHeight = await getProductRowHeight(page, bottomRow)
  return {
    x: grid.x + left,
    y: grid.y + PRODUCT_HEADER_HEIGHT + top,
    width: right + rightWidth - left,
    height: bottom + bottomHeight - top,
  }
}

async function expectVisualRectNear(
  locator: ReturnType<Page['locator']>,
  expected: {
    readonly x: number
    readonly y: number
    readonly width: number
    readonly height: number
  },
  label: string,
): Promise<void> {
  await expect(locator, `${label} should be unique`).toHaveCount(1)
  const actual = await locator.boundingBox()
  if (!actual) {
    throw new Error(`${label} is not visible`)
  }
  expect(actual.x, `${label} x`).toBeCloseTo(expected.x, 0)
  expect(actual.y, `${label} y`).toBeCloseTo(expected.y, 0)
  expect(actual.width, `${label} width`).toBeCloseTo(expected.width, 0)
  expect(actual.height, `${label} height`).toBeCloseTo(expected.height, 0)
}

test('@browser-ci web app keeps the source selection visible while fill-dragging downward inside the grid', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 420 })
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('fill-handle-down-selection'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, 2, 2, 2, 5)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 8 })

  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="selection-border"]'),
    await getProductCellRangeBox(page, 1, 1, 2, 2),
    'source selection border while dragging downward',
  )
  await expectVisualRectNear(
    page.locator('[data-grid-fill-preview="true"]'),
    await getProductCellRangeBox(page, 1, 3, 2, 5),
    'downward fill preview',
  )

  await page.mouse.up()
})

test('@browser-ci web app keeps the source selection visible while fill-dragging rightward inside the grid', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 420 })
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('fill-handle-right-selection'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, 2, 2, 3, 2)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 8 })

  await expectVisualRectNear(
    page.locator('[data-grid-selection-visual-role="selection-border"]'),
    await getProductCellRangeBox(page, 1, 1, 2, 2),
    'source selection border while dragging rightward',
  )
  await expectVisualRectNear(
    page.locator('[data-grid-fill-preview="true"]'),
    await getProductCellRangeBox(page, 3, 1, 3, 2),
    'rightward fill preview',
  )

  await page.mouse.up()
})

test('@browser-ci web app keeps the fill preview visible while fill-dragging below the grid', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 420 })
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('fill-handle-below-grid-selection'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const { sourceX, sourceY } = await getProductFillHandleDragPoints(page, 2, 2, 2, 2)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(sourceX, grid.y + grid.height + 16, { steps: 8 })

  await expect(page.locator('[data-grid-fill-preview="true"]')).toBeVisible()
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

  await page.mouse.up()
})

test('@browser-ci web app keeps the fill preview visible while fill-dragging past the right grid edge', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 420 })
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('fill-handle-right-grid-selection'))}&persist=0`)
  await waitForWorkbookReady(page)

  await dragProductBodySelection(page, 1, 1, 2, 2)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const { sourceX, sourceY } = await getProductFillHandleDragPoints(page, 2, 2, 2, 2)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(grid.x + grid.width + 16, sourceY, { steps: 8 })

  await expect(page.locator('[data-grid-fill-preview="true"]')).toBeVisible()
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')

  await page.mouse.up()
})
