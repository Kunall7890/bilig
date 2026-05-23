import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'

export function sameCorpusChromiumLaunchOptions(headless: boolean): { readonly args: string[]; readonly headless: boolean } {
  return {
    args: ['--enable-unsafe-webgpu', '--ignore-gpu-blocklist'],
    headless,
  }
}

export async function waitForNextFrame(page: Page): Promise<void> {
  await page.evaluate(async () => {
    await new Promise<void>((resolveFrame) => requestAnimationFrame(() => resolveFrame()))
  })
}

export async function collectFrameIntervals(page: Page, frameCount: number): Promise<number[]> {
  const frameIntervals = await page.evaluate(async (targetFrameCount) => {
    const intervals: number[] = []
    let previous = performance.now()
    await new Promise<void>((finish) => {
      const step = (now: number): void => {
        intervals.push(now - previous)
        previous = now
        if (intervals.length >= targetFrameCount) {
          finish()
          return
        }
        requestAnimationFrame(step)
      }
      requestAnimationFrame(step)
    })
    return intervals
  }, frameCount)
  return frameIntervals
}

export async function settleFrames(page: Page, frames: number): Promise<void> {
  await page.evaluate(async (frameCount) => {
    await Array.from({ length: frameCount }).reduce<Promise<void>>(async (previous) => {
      await previous
      await new Promise<void>((resolveFrame) => requestAnimationFrame(() => resolveFrame()))
    }, Promise.resolve())
  }, frames)
}

export function productLimitations(product: UiResponsivenessSameCorpusProduct, storageStatePath: string | null): string[] {
  const authLimitations = storageStatePath ? ['Browser context used an explicit Playwright storage state for authenticated access.'] : []
  if (product === 'bilig') {
    return [
      'Bilig timing is captured from the supplied app URL and benchmarkCorpus route; 10x claims require the page runtime proof to report a production build.',
      ...authLimitations,
    ]
  }
  if (product === 'google-sheets') {
    return [
      'Google Sheets timing requires the supplied URL to be browser-accessible and loaded with the same benchmark corpus.',
      ...authLimitations,
    ]
  }
  return [
    'Microsoft Excel Web timing requires the supplied URL to be browser-accessible and loaded with the same benchmark corpus.',
    ...authLimitations,
  ]
}
