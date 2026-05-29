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

  it('does not count Google Sheets selected-cell chrome as target fill color', () => {
    document.body.innerHTML = `
      <div class="waffle-cell active-cell" aria-selected="true" style="background-color: rgb(11, 87, 208);">segment-5</div>
    `
    const cell = document.querySelector<HTMLElement>('.waffle-cell')
    expect(cell).not.toBeNull()
    setRect(cell!, { height: 22, width: 104, x: 120, y: 80 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 120, y: 80 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: 'segment-5',
      visibleText: 'segment-5',
    })
  })

  it('rejects active in-grid editor text as committed target-cell proof', () => {
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

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 220, y: 144 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
      value: null,
      visibleText: null,
    })
  })

  it('prefers stale committed grid text over overlapping editor text so the validator can reject drift', () => {
    document.body.innerHTML = `
      <div class="waffle-cell">stale committed text</div>
      <div class="waffle-cell-input" contenteditable="true">same-corpus-edit-1</div>
    `
    const cell = document.querySelector<HTMLElement>('.waffle-cell')
    const editor = document.querySelector<HTMLElement>('.waffle-cell-input')
    expect(cell).not.toBeNull()
    expect(editor).not.toBeNull()
    setRect(cell!, { height: 22, width: 104, x: 220, y: 144 })
    setRect(editor!, { height: 22, width: 104, x: 220, y: 144 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 220, y: 144 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: 'stale committed text',
      visibleText: 'stale committed text',
    })
  })

  it('reads committed target text from a real grid-cell child without accepting broad wrapper text', () => {
    document.body.innerHTML = `
      <div class="waffle-cell-wrapper">wrong wrapper text</div>
      <div role="gridcell"><span>committed aria child text</span></div>
    `
    const wrapper = document.querySelector<HTMLElement>('.waffle-cell-wrapper')
    const cell = document.querySelector<HTMLElement>('[role="gridcell"]')
    expect(wrapper).not.toBeNull()
    expect(cell).not.toBeNull()
    setRect(wrapper!, { height: 220, width: 520, x: 80, y: 80 })
    setRect(cell!, { height: 22, width: 104, x: 220, y: 144 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 220, y: 144 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: 'committed aria child text',
      visibleText: 'committed aria child text',
    })
  })

  it('rejects bare selected wrappers as target-cell text proof', () => {
    document.body.innerHTML = `
      <div aria-selected="true">selected wrapper text from chrome</div>
      <div class="waffle-cell">outside target</div>
    `
    const wrapper = document.querySelector<HTMLElement>('[aria-selected="true"]')
    const outsideCell = document.querySelector<HTMLElement>('.waffle-cell')
    expect(wrapper).not.toBeNull()
    expect(outsideCell).not.toBeNull()
    setRect(wrapper!, { height: 220, width: 520, x: 80, y: 80 })
    setRect(outsideCell!, { height: 22, width: 104, x: 420, y: 144 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 220, y: 144 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
      value: null,
      visibleText: null,
    })
  })

  it('rejects grid-looking border strips as target-cell proof', () => {
    document.body.innerHTML = `
      <div class="waffle-cell">border-only text</div>
    `
    const strip = document.querySelector<HTMLElement>('.waffle-cell')
    expect(strip).not.toBeNull()
    setRect(strip!, { height: 2, width: 104, x: 120, y: 80 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 120, y: 80 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
      value: null,
      visibleText: null,
    })
  })

  it('ignores child colors that only paint target selection chrome', () => {
    document.body.innerHTML = `
      <div class="waffle-cell">
        <div class="range-border" style="background-color: rgb(52, 168, 83);"></div>
      </div>
    `
    const cell = document.querySelector<HTMLElement>('.waffle-cell')
    const border = document.querySelector<HTMLElement>('.range-border')
    expect(cell).not.toBeNull()
    expect(border).not.toBeNull()
    setRect(cell!, { height: 22, width: 104, x: 120, y: 80 })
    setRect(border!, { height: 2, width: 104, x: 120, y: 80 })

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 120, y: 80 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: null,
      visibleText: null,
    })
  })

  it('uses a target grid-cell aria label while stripping the coordinate prefix', () => {
    document.body.innerHTML = `
      <div role="gridcell" aria-label="F5 same-corpus-edit-1"></div>
    `
    const cell = document.querySelector<HTMLElement>('[role="gridcell"]')
    expect(cell).not.toBeNull()
    setRect(cell!, { height: 22, width: 104, x: 120, y: 144 })

    expect(
      readSameCorpusVisibleTargetCellReadbackFromPage({
        targetBox: { height: 22, width: 104, x: 120, y: 144 },
        targetRange: 'F5',
      }),
    ).toEqual({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: 'same-corpus-edit-1',
      visibleText: 'same-corpus-edit-1',
    })
  })

  it('rejects an aria-only cell coordinate without committed target text', () => {
    document.body.innerHTML = `
      <div role="gridcell" aria-label="F5"></div>
    `
    const cell = document.querySelector<HTMLElement>('[role="gridcell"]')
    expect(cell).not.toBeNull()
    setRect(cell!, { height: 22, width: 104, x: 120, y: 144 })

    expect(
      readSameCorpusVisibleTargetCellReadbackFromPage({
        targetBox: { height: 22, width: 104, x: 120, y: 144 },
        targetRange: 'F5',
      }),
    ).toEqual({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: null,
      visibleText: null,
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

  it('rejects Google Sheets selection-border elements as target-cell proof', () => {
    document.body.innerHTML = `
      <div class="waffle-border-cell-active range-border active-cell-border" style="background-color: rgb(11, 87, 208);"></div>
      <div class="autofill-handle" aria-label="Fill handle" style="background-color: rgb(11, 87, 208);"></div>
    `
    for (const element of Array.from(document.querySelectorAll<HTMLElement>('div'))) {
      setRect(element, { height: 22, width: 104, x: 120, y: 144 })
    }

    expect(readSameCorpusVisibleTargetCellReadbackFromPage({ targetBox: { height: 22, width: 104, x: 120, y: 144 } })).toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
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

  it('does not use Google Sheets formula-bar text as target-cell visible proof', async () => {
    const page = fakeGoogleSheetsCanvasTargetPage({ formulaBarText: 'note-4-5', selectedRange: 'F5' })

    await expect(
      readSameCorpusVisibleMutationTargetReadback({
        page,
        product: 'google-sheets',
        target: {
          endAddress: 'F5',
          sheetId: 'gid:160971404',
          sheetName: 'WideGrid',
          startAddress: 'F5',
          targetRange: 'F5',
        },
        workload: 'edit-visible-cell',
      }),
    ).resolves.toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
      value: null,
      visibleText: null,
    })
  })

  it('does not write a target-cell screenshot when semantic proof came from formula-bar chrome', async () => {
    const page = fakePageWithExternalTargetBox(() => {
      throw new Error('Formula-bar semantic readback must not produce a target-cell screenshot')
    })

    await expect(
      captureSameCorpusMutationTargetScreenshotProof({
        page,
        phase: 'after',
        product: 'google-sheets',
        relativeScreenshotPath: 'same-corpus/google-sheets/formula-bar-target.png',
        sampleIndex: 0,
        screenshotPath: '/tmp/bilig-formula-bar-target-should-not-be-written.png',
        semanticReadback: {
          fillColor: null,
          formula: null,
          source: 'visible-formula-bar',
          value: 'same-corpus-edit-1',
          visibleText: 'same-corpus-edit-1',
        },
        target: {
          endAddress: 'F5',
          sheetId: 'gid:160971404',
          sheetName: 'WideGrid',
          startAddress: 'F5',
          targetRange: 'F5',
        },
        workload: 'edit-visible-cell',
      }),
    ).resolves.toMatchObject({
      scope: 'visible-grid-fallback',
      screenshotPath: null,
      screenshotSha256: null,
    })
  })

  it('does not use an active editor overlay as external target-cell screenshot geometry', async () => {
    const page = fakePageWithGoogleSheetsEditorOverlay()

    await expect(
      captureSameCorpusMutationTargetScreenshotProof({
        page,
        phase: 'after',
        product: 'google-sheets',
        relativeScreenshotPath: 'same-corpus/google-sheets/editor-overlay.png',
        sampleIndex: 0,
        screenshotPath: '/tmp/bilig-editor-overlay-should-not-be-written.png',
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

  it('does not call an empty Google Sheets target cell committed text proof for edit/formula workloads', async () => {
    const page = fakeGoogleSheetsEmptyTargetPage()

    await expect(
      readSameCorpusVisibleMutationTargetReadback({
        page,
        product: 'google-sheets',
        target: {
          endAddress: 'F5',
          sheetId: 'gid:160971404',
          sheetName: 'WideGrid',
          startAddress: 'F5',
          targetRange: 'F5',
        },
        workload: 'edit-visible-cell',
      }),
    ).resolves.toEqual({
      fillColor: null,
      formula: null,
      source: 'unknown',
      value: null,
      visibleText: null,
    })
  })

  it('captures external mutation target screenshots from the cell interior, not the selection border', async () => {
    let capturedClip: { height: number; width: number; x: number; y: number } | null = null
    const page = fakePageWithExternalTargetBox((clip) => {
      capturedClip = clip
    })

    await expect(
      captureSameCorpusMutationTargetScreenshotProof({
        page,
        phase: 'after',
        product: 'google-sheets',
        relativeScreenshotPath: 'same-corpus/google-sheets/target.png',
        sampleIndex: 0,
        screenshotPath: '/tmp/bilig-target-interior.png',
        semanticReadback: {
          fillColor: 'rgb(52, 168, 83)',
          formula: null,
          source: 'visible-grid-cell',
          value: 'grid committed text',
          visibleText: 'grid committed text',
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
      scope: 'target-cell',
      screenshotPath: 'same-corpus/google-sheets/target.png',
    })

    expect(capturedClip).toEqual({ height: 14, width: 64, x: 140, y: 84 })
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

function fakePageWithExternalTargetBox(
  onScreenshotClip: (clip: { readonly height: number; readonly width: number; readonly x: number; readonly y: number }) => void,
): Page {
  const page = {
    frames: () => [],
    viewportSize: () => ({ height: 200, width: 300 }),
    locator: (selector: string) => ({
      boundingBox: async () => (selector === '.waffle-active-cell' ? { height: 22, width: 104, x: 120, y: 80 } : null),
      count: async () => (selector === '.waffle-active-cell' ? 1 : 0),
      first() {
        return this
      },
    }),
    evaluate: async () => null,
    screenshot: async (options: {
      readonly clip?: { readonly height: number; readonly width: number; readonly x: number; readonly y: number }
    }) => {
      if (options.clip) {
        onScreenshotClip(options.clip)
      }
      return Buffer.from('target screenshot')
    },
  }
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises the external target screenshot path.
  return page as unknown as Page
}

function fakePageWithGoogleSheetsEditorOverlay(): Page {
  const page = {
    frames: () => [],
    locator: (selector: string) => ({
      boundingBox: async () => (selector === '.waffle-cell-input' ? { height: 22, width: 104, x: 120, y: 80 } : null),
      count: async () => (selector === '.waffle-cell-input' ? 1 : 0),
      first() {
        return this
      },
    }),
    evaluate: async () => null,
    screenshot: async () => {
      throw new Error('Editor overlay must not be captured as target-cell proof')
    },
  }
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises editor-overlay exclusion.
  return page as unknown as Page
}

function fakePageWithoutTargetBox(): Page {
  const page = {
    frames: () => [],
    keyboard: {
      press: async () => undefined,
    },
    locator: () => ({
      boundingBox: async () => null,
      count: async () => 0,
      first() {
        return this
      },
      inputValue: async () => '',
    }),
    evaluate: async () => {
      throw new Error('Formula bar fallback must not be used for external target-cell proof')
    },
    screenshot: async () => {
      throw new Error('Whole-grid fallback screenshot must not be written for missing external target-cell proof')
    },
  }
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises this narrow no-target-box page surface.
  return page as unknown as Page
}

function fakePageWithExternalGridButNoTargetBox(): Page {
  const page = {
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
  }
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises missing target box plus present grid surface.
  return page as unknown as Page
}

function fakeGoogleSheetsCanvasTargetPage(args: { readonly formulaBarText: string; readonly selectedRange: string }): Page {
  let evaluateCallCount = 0
  const page = {
    frames: () => [],
    keyboard: {
      press: async () => undefined,
    },
    locator: (selector: string) => ({
      boundingBox: async () => (selector === '.waffle-cell-input' ? { height: 22, width: 104, x: 120, y: 80 } : null),
      count: async () => (selector === '.waffle-cell-input' ? 1 : 0),
      first() {
        return this
      },
      inputValue: async () => args.selectedRange,
    }),
    evaluate: async () => {
      evaluateCallCount += 1
      if (evaluateCallCount === 1) {
        return null
      }
      if (evaluateCallCount === 2) {
        return {
          fillColor: null,
          formula: null,
          source: 'visible-grid-cell',
          value: null,
          visibleText: null,
        }
      }
      return args.formulaBarText
    },
  }
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises the canvas-grid readback fallback path.
  return page as unknown as Page
}

function fakeGoogleSheetsEmptyTargetPage(): Page {
  const page = {
    frames: () => [],
    locator: (selector: string) => ({
      boundingBox: async () => (selector === '.waffle-cell-input' ? { height: 22, width: 104, x: 120, y: 80 } : null),
      count: async () => (selector === '.waffle-cell-input' ? 1 : 0),
      first() {
        return this
      },
      inputValue: async () => 'F5',
    }),
    evaluate: async () => ({
      fillColor: null,
      formula: null,
      source: 'visible-grid-cell',
      value: null,
      visibleText: null,
    }),
  }
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- Unit test only exercises empty external target-cell readback.
  return page as unknown as Page
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
