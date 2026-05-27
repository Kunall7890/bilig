import type { Locator, Page } from '@playwright/test'
import { describe, expect, it } from 'vitest'

import { performSameCorpusFillColorOperation } from '../ui-responsiveness-same-corpus-workload-runner.ts'

describe('same-corpus workload runner', () => {
  it('uses exact scoped Google Sheets swatches instead of broad color-family fallbacks', async () => {
    const events: string[] = []
    const page = fakeFillColorPage(events)

    await performSameCorpusFillColorOperation(page, 'google-sheets', 0)

    expect(events).toEqual(['label:Fill color', expect.stringContaining('.goog-menu')])
    expect(events[1]).toContain('light cornflower blue 3')
    expect(events[1]).not.toContain('light cornflower blue"]')
  })
})

function fakeFillColorPage(events: string[]): Page {
  const locatorFor = (name: string): Locator => {
    const locator = {
      first: () => ({
        click: async () => {
          events.push(name)
          if (name === 'label:Fill color') {
            return
          }
          if (name.startsWith('selector:') && name.includes('.goog-menu') && name.includes('light cornflower blue 3')) {
            return
          }
          throw new Error(`unexpected fill locator: ${name}`)
        },
      }),
    }
    // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- This test supplies the minimal Playwright Locator surface used by the fill-color path.
    return locator as unknown as Locator
  }

  const page = {
    getByLabel: (name: string | RegExp) => locatorFor(`label:${String(name)}`),
    getByRole: (role: string, options?: { readonly name?: string | RegExp }) => locatorFor(`role:${role}:${String(options?.name ?? '')}`),
    locator: (selector: string) => locatorFor(`selector:${selector}`),
  }
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- This test supplies the minimal Playwright Page surface used by the fill-color path.
  return page as unknown as Page
}
