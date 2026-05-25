import { randomUUID } from 'node:crypto'
import { expect, type Locator, type Page } from '@playwright/test'

export const PRODUCT_ROW_MARKER_WIDTH = 40
export const PRODUCT_COLUMN_WIDTH = 104
export const PRODUCT_HEADER_HEIGHT = 22
export const PRODUCT_ROW_HEIGHT = 22
export const PRIMARY_MODIFIER = process.platform === 'darwin' ? 'Meta' : 'Control'
export const remoteSyncEnabled = process.env['BILIG_E2E_REMOTE_SYNC'] !== '0'

export function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

export function createTestDocumentId(prefix: string): string {
  return `${prefix}-${randomUUID()}`
}

function parseDimensionOverrides(raw: string | null): Record<string, number> {
  if (!raw) {
    return {}
  }
  const parsed: unknown = JSON.parse(raw)
  if (typeof parsed !== 'object' || parsed === null) {
    return {}
  }
  const entries = Object.entries(parsed).filter((entry): entry is [string, number] => typeof entry[1] === 'number')
  return Object.fromEntries(entries)
}

function readProductGridDimensionAttributes(
  page: Page,
  attributes: {
    readonly defaultAttribute: string
    readonly overridesAttribute: string
  },
): Promise<{
  readonly defaultSizeRaw: string | null
  readonly overridesRaw: string | null
}> {
  return page.evaluate(({ defaultAttribute, overridesAttribute }) => {
    const grid = document.querySelector('[data-testid="sheet-grid"]')
    if (!grid) {
      throw new Error('sheet grid is not attached')
    }
    return {
      defaultSizeRaw: grid.getAttribute(defaultAttribute),
      overridesRaw: grid.getAttribute(overridesAttribute),
    }
  }, attributes)
}

export async function getProductColumnWidth(page: Page, columnIndex: number) {
  const { defaultSizeRaw, overridesRaw } = await readProductGridDimensionAttributes(page, {
    defaultAttribute: 'data-default-column-width',
    overridesAttribute: 'data-column-width-overrides',
  })
  const defaultWidth = Number(defaultSizeRaw ?? String(PRODUCT_COLUMN_WIDTH))
  const overrides = parseDimensionOverrides(overridesRaw)
  return overrides[String(columnIndex)] ?? defaultWidth
}

export async function waitForProductColumnWidthChange(
  page: Page,
  columnIndex: number,
  previousWidth: number,
  timeoutMs = 10_000,
): Promise<number> {
  return await page.evaluate(
    ({ columnIndex: targetColumnIndex, defaultColumnWidth, previousWidth: initialWidth, timeoutMs: waitTimeoutMs }) =>
      new Promise<number>((resolve, reject) => {
        const grid = document.querySelector('[data-testid="sheet-grid"]')
        if (!grid) {
          reject(new Error('sheet grid is not attached'))
          return
        }

        const readColumnWidth = () => {
          const defaultWidth = Number(grid.getAttribute('data-default-column-width') ?? String(defaultColumnWidth))
          const overridesRaw = grid.getAttribute('data-column-width-overrides')
          if (!overridesRaw) {
            return defaultWidth
          }
          const parsed: unknown = JSON.parse(overridesRaw)
          if (typeof parsed !== 'object' || parsed === null) {
            return defaultWidth
          }
          const override = Object.entries(parsed).find(
            (entry): entry is [string, number] => entry[0] === String(targetColumnIndex) && typeof entry[1] === 'number',
          )
          return override?.[1] ?? defaultWidth
        }

        let timeout: ReturnType<typeof window.setTimeout> | null = null
        const observer = new MutationObserver(() => {
          const nextWidth = readColumnWidth()
          if (nextWidth === initialWidth) {
            return
          }
          observer.disconnect()
          if (timeout !== null) {
            window.clearTimeout(timeout)
          }
          resolve(nextWidth)
        })

        const currentWidth = readColumnWidth()
        if (currentWidth !== initialWidth) {
          resolve(currentWidth)
          return
        }

        timeout = window.setTimeout(() => {
          observer.disconnect()
          reject(
            new Error(
              `column ${String(targetColumnIndex)} width did not change from ${String(initialWidth)} within ${String(waitTimeoutMs)}ms`,
            ),
          )
        }, waitTimeoutMs)
        observer.observe(grid, {
          attributeFilter: ['data-column-width-overrides', 'data-default-column-width'],
          attributes: true,
        })
      }),
    { columnIndex, defaultColumnWidth: PRODUCT_COLUMN_WIDTH, previousWidth, timeoutMs },
  )
}

export async function getProductRowHeight(page: Page, rowIndex: number) {
  const { defaultSizeRaw, overridesRaw } = await readProductGridDimensionAttributes(page, {
    defaultAttribute: 'data-default-row-height',
    overridesAttribute: 'data-row-height-overrides',
  })
  const defaultHeight = Number(defaultSizeRaw ?? String(PRODUCT_ROW_HEIGHT))
  const overrides = parseDimensionOverrides(overridesRaw)
  return overrides[String(rowIndex)] ?? defaultHeight
}

export async function getProductRowTop(page: Page, rowIndex: number) {
  const heights = await Promise.all(Array.from({ length: rowIndex }, (_, index) => getProductRowHeight(page, index)))
  return heights.reduce((total, height) => total + height, 0)
}

export async function getProductColumnLeft(page: Page, columnIndex: number) {
  const widths = await Promise.all(Array.from({ length: columnIndex }, (_, index) => getProductColumnWidth(page, index)))
  return PRODUCT_ROW_MARKER_WIDTH + widths.reduce((total, width) => total + width, 0)
}

export async function getBox(locator: Locator) {
  await expect(locator).toBeVisible()
  const box = await locator.boundingBox()
  if (!box) {
    throw new Error('locator is not visible')
  }
  return box
}

export async function clickProductCell(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  options?: {
    shift?: boolean
  },
) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  if (options?.shift) {
    await page.keyboard.down('Shift')
  }
  try {
    await page.mouse.click(
      grid.x + columnLeft + Math.floor(columnWidth / 2),
      grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
    )
  } finally {
    if (options?.shift) {
      await page.keyboard.up('Shift')
    }
  }
}
