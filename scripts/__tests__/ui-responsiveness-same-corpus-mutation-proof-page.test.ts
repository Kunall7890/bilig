// @vitest-environment jsdom

import { afterEach, describe, expect, it } from 'vitest'
import type { Page } from '@playwright/test'

import {
  captureSameCorpusMutationTargetScreenshotProof,
  readSameCorpusVisibleMutationTargetReadback,
  readSameCorpusVisibleTargetCellReadbackFromPage,
} from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'

describe('same-corpus mutation target page proof helpers', () => {
  afterEach(() => {
    document.body.innerHTML = ''
  })

  it('reads Google Sheets target-cell text and fill from grid DOM instead of formula bar chrome', () => {
    document.body.innerHTML = `
      <input id="t-formula-bar-input" value="stale formula bar text" />
      <div class="waffle-cell active-cell" style="background-color: rgb(52, 168, 83);">grid committed text</div>
    `
    const cell = document.querySelector<HTMLElement>('.waffle-cell')
    expect(cell).not.toBeNull()
    setRect(cell!, { height: 22, width: 104, x: 120, y: 80 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 120, y: 80 } })).toEqual({
      fillColor: 'rgb(52, 168, 83)',
      formula: null,
      source: 'visible-grid-cell',
      value: 'grid committed text',
      visibleText: 'grid committed text',
    })
  })

  it('uses an active in-grid editor only when it spatially overlaps the selected cell', () => {
    document.body.innerHTML = `
      <div class="waffle-cell-input" contenteditable="true">2</div>
      <div class="waffle-cell-input" contenteditable="true">outside target</div>
    `
    const [targetEditor, outsideEditor] = Array.from(document.querySelectorAll<HTMLElement>('.waffle-cell-input'))
    if (!targetEditor || !outsideEditor) {
      throw new Error('Test setup failed to create grid editors')
    }
    setRect(targetEditor, { height: 22, width: 104, x: 220, y: 144 })
    setRect(outsideEditor, { height: 22, width: 104, x: 420, y: 144 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 220, y: 144 } })).toMatchObject({
      source: 'visible-grid-cell',
      value: '2',
      visibleText: '2',
    })
  })

  it('treats an empty overlapping grid cell as target-cell proof instead of falling back to chrome', () => {
    document.body.innerHTML = `
      <input id="t-formula-bar-input" value="stale formula bar text" />
      <div class="waffle-cell active-cell"></div>
    `
    const cell = document.querySelector<HTMLElement>('.waffle-cell')
    expect(cell).not.toBeNull()
    setRect(cell!, { height: 22, width: 104, x: 120, y: 144 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 120, y: 144 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: null,
      visibleText: null,
    })
  })

  it('returns an unknown source when no target-cell grid evidence overlaps the selection', () => {
    document.body.innerHTML = `
      <input id="t-formula-bar-input" value="formula bar only" />
      <div class="waffle-cell">outside target</div>
    `
    const cell = document.querySelector<HTMLElement>('.waffle-cell')
    setRect(cell!, { height: 22, width: 104, x: 420, y: 144 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 120, y: 144 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
      value: null,
      visibleText: null,
    })
  })

  it('does not fall back to formula-bar chrome when external target-cell evidence is absent', async () => {
    const page = fakePageWithoutTargetBox()

    await expect(
      readSameCorpusVisibleMutationTargetReadback({
        page,
        product: 'google-sheets',
        target: {
          endAddress: 'C5',
          sheetId: 'sheet-id',
          sheetName: 'Sheet1',
          startAddress: 'C5',
          targetRange: 'C5',
        },
      }),
    ).resolves.toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
      value: null,
      visibleText: null,
    })
  })

  it('does not write a misleading external target screenshot when the selected cell box is missing', async () => {
    const page = fakePageWithExternalGridButNoTargetBox()

    await expect(
      captureSameCorpusMutationTargetScreenshotProof({
        page,
        phase: 'after',
        product: 'google-sheets',
        relativeScreenshotPath: 'same-corpus/google-sheets/missing-target.png',
        sampleIndex: 0,
        screenshotPath: '/tmp/bilig-missing-target-should-not-be-written.png',
        semanticReadback: {
          fillColor: null,
          formula: null,
          source: 'unknown',
          value: null,
          visibleText: null,
        },
        target: {
          endAddress: 'C5',
          sheetId: 'sheet-id',
          sheetName: 'Sheet1',
          startAddress: 'C5',
          targetRange: 'C5',
        },
        workload: 'edit-visible-cell',
      }),
    ).resolves.toMatchObject({
      scope: 'visible-grid-fallback',
      screenshotPath: null,
      screenshotSha256: null,
    })
  })
})

function fakePageWithoutTargetBox(): Page {
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises this narrow no-target-box page surface.
  return {
    frames: () => [],
    locator: () => ({
      boundingBox: async () => null,
      count: async () => 0,
      first() {
        return this
      },
    }),
    evaluate: async () => {
      throw new Error('Formula bar fallback must not be used for external target-cell proof')
    },
    screenshot: async () => {
      throw new Error('Whole-grid fallback screenshot must not be written for missing external target-cell proof')
    },
  } as unknown as Page
}

function fakePageWithExternalGridButNoTargetBox(): Page {
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises missing target box plus present grid surface.
  return {
    frames: () => [],
    locator: (selector: string) => ({
      boundingBox: async () => null,
      count: async () => (selector === '.grid-scrollable-wrapper' ? 1 : 0),
      first() {
        return this
      },
      screenshot: async () => {
        throw new Error('Whole-grid fallback screenshot must not be written for missing external target-cell proof')
      },
    }),
  } as unknown as Page
}

function setRect(
  element: HTMLElement,
  rect: { readonly height: number; readonly width: number; readonly x: number; readonly y: number },
): void {
  element.getBoundingClientRect = () =>
    ({
      bottom: rect.y + rect.height,
      height: rect.height,
      left: rect.x,
      right: rect.x + rect.width,
      top: rect.y,
      width: rect.width,
      x: rect.x,
      y: rect.y,
      toJSON: () => rect,
    }) as DOMRect
}
