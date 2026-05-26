import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { describe, expect, it } from 'vitest'

import { exportXlsx } from '../../packages/excel-import/src/index.js'
import { buildWorkbookBenchmarkCorpus } from '../../packages/benchmarks/src/workbook-corpus.js'
import {
  assertSameCorpusBrowserRunAllowed,
  assertSameCorpusCaptureCurrentContractEvidenceReady,
  assertSameCorpusCaptureEvidenceReady,
  assertSameCorpusPreflightReady,
  assertProductionBiligEvidenceSource,
  buildSameCorpusFingerprint,
  buildSameCorpusCaptureArtifact,
  collectSameCorpusProductMeasurements,
  parseCaptureArgs,
  parseEmitXlsxArgs,
  parsePreflightArgs,
  parseSaveStorageStateArgs,
  verifyXlsxCorpusFingerprint,
} from '../capture-ui-responsiveness-same-corpus.ts'
import {
  parseSameCorpusCapture,
  sameCorpusScenarioCaseFields,
  type SameCorpusOperationResponseProof,
} from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadMutatesWorkbook,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusMutatingWorkload,
  type UiResponsivenessSameCorpusWorkload,
} from '../ui-responsiveness-same-corpus-workloads.ts'
import {
  buildCaptureScenarioProof,
  isSameCorpusProductPixelGridProofComplete,
  type SameCorpusMutationTargetProof,
  type SameCorpusProductVisualProof,
} from '../ui-responsiveness-same-corpus-proof.ts'
import {
  normalizeSameCorpusMutationTargetSelection,
  selectGoogleSheetsTargetRange,
  sameCorpusMutationTargetRangeForSample,
  type SameCorpusNameBoxPage,
} from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'
import { sameCorpusChromiumLaunchOptions } from '../ui-responsiveness-same-corpus-page-utils.ts'
import { sameCorpusScrollProbeSelectorsForProduct } from '../ui-responsiveness-same-corpus-scroll-page.ts'
import { readGoogleSheetsNameBoxSelection, type SameCorpusNameBoxReaderPage } from '../ui-responsiveness-same-corpus-semantic-proof.ts'
import {
  incumbentEditableWorkloadBlocker,
  measureProductWorkload,
  sameCorpusEditVisibleCellValue,
  sameCorpusFillColorSwatchLabel,
  sameCorpusFormulaEditFormula,
  sameCorpusKeyboardOperations,
  sameCorpusWorkbookRestoreOperations,
  sameCorpusWorkloadMutatesWorkbook,
} from '../ui-responsiveness-same-corpus-workload-runner.ts'
import {
  biligInteractionVisibleResponseToken,
  measureVisibleNonScrollResponse,
  visibleNonScrollResponseChanged,
  visibleNonScrollResponseNeedsScreenshot,
  type VisibleNonScrollResponseSignature,
} from '../ui-responsiveness-same-corpus-visible-response-page.ts'

const sameCorpusFixtureCheckedCells = [
  { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
  { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
  { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
] as const
const wideMixedSameCorpusFingerprint = buildSameCorpusFingerprint(buildWorkbookBenchmarkCorpus('wide-mixed-250k')).corpusFingerprint

describe('same-corpus UI responsiveness capture CLI', () => {
  it('builds a default Bilig benchmark URL from the selected corpus', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--corpus',
      'dense-mixed-250k',
    ])

    expect(args).toMatchObject({
      allowIncompleteEvidence: false,
      biligProductionHost: '127.0.0.1',
      biligProductionPort: 4180,
      biligUrl: 'http://localhost:5173/?benchmarkCorpus=dense-mixed-250k',
      biligUrlSource: 'default-dev',
      biligStorageStatePath: null,
      corpusId: 'dense-mixed-250k',
      deltaX: 0,
      deltaY: 720,
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      googleSheetsStorageStatePath: null,
      headless: true,
      microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      microsoftExcelWebStorageStatePath: null,
      readyTimeoutMs: 60000,
      sampleCount: 3,
      storageStatePath: null,
    })
    expect(args.outputPath.endsWith('/tmp/ui-capture.json')).toBe(true)
  })

  it('accepts explicit browser and workload options', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--bilig-url',
      'http://127.0.0.1:4173/?benchmarkCorpus=wide-mixed-250k',
      '--samples',
      '5',
      '--delta-x',
      '1024',
      '--delta-y',
      '0',
      '--ready-timeout-ms',
      '120000',
      '--storage-state',
      'tmp/shared-state.json',
      '--google-sheets-storage-state',
      'tmp/google-state.json',
      '--microsoft-excel-web-storage-state',
      'tmp/microsoft-state.json',
      '--bilig-storage-state',
      'tmp/bilig-state.json',
      '--headed',
    ])

    expect(args).toMatchObject({
      allowIncompleteEvidence: false,
      biligProductionHost: '127.0.0.1',
      biligProductionPort: 4180,
      biligUrl: 'http://127.0.0.1:4173/?benchmarkCorpus=wide-mixed-250k',
      biligUrlSource: 'explicit',
      deltaX: 1024,
      deltaY: 0,
      headless: false,
      readyTimeoutMs: 120000,
      sampleCount: 5,
    })
    expect(args.storageStatePath?.endsWith('/tmp/shared-state.json')).toBe(true)
    expect(args.googleSheetsStorageStatePath?.endsWith('/tmp/google-state.json')).toBe(true)
    expect(args.microsoftExcelWebStorageStatePath?.endsWith('/tmp/microsoft-state.json')).toBe(true)
    expect(args.biligStorageStatePath?.endsWith('/tmp/bilig-state.json')).toBe(true)
  })

  it('parses production Bilig preview capture options', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--serve-bilig-production',
      '--bilig-production-host',
      '0.0.0.0',
      '--bilig-production-port',
      '4181',
    ])

    expect(args).toMatchObject({
      biligProductionHost: '0.0.0.0',
      biligProductionPort: 4181,
      biligUrl: 'http://127.0.0.1:4181/?benchmarkCorpus=wide-mixed-250k&persist=0',
      biligUrlSource: 'served-production',
    })
  })

  it('rejects ambiguous Bilig production serving and explicit URL options', () => {
    expect(() =>
      parseCaptureArgs([
        '--output',
        'tmp/ui-capture.json',
        '--google-sheets-url',
        'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        '--serve-bilig-production',
        '--bilig-url',
        'https://example.test',
      ]),
    ).toThrow('Use either --serve-bilig-production or --bilig-url, not both.')
  })

  it('fails fast before dominance captures use the default dev Bilig URL', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
    ])

    expect(() => assertProductionBiligEvidenceSource(args)).toThrow(/needs production Bilig runtime proof/u)
  })

  it('allows explicit diagnostic captures to use the default dev Bilig URL', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--allow-incomplete-evidence',
    ])

    expect(() => assertProductionBiligEvidenceSource(args)).not.toThrow()
  })

  it('rejects missing incumbent URLs because the generated proof must be comparable', () => {
    expect(() => parseCaptureArgs(['--output', 'tmp/ui-capture.json'])).toThrow('Missing required arguments.')
  })

  it('rejects blank capture argument values', () => {
    expect(() =>
      parseCaptureArgs(['--output', '   ', '--google-sheets-url', 'https://docs.google.com/spreadsheets/d/sheet-id/edit']),
    ).toThrow('Missing value after --output')
  })

  it('parses incomplete-evidence override for diagnostic captures only', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--allow-incomplete-evidence',
    ])

    expect(args.allowIncompleteEvidence).toBe(true)
  })

  it('parses incumbent-only same-corpus preflight options', () => {
    const args = parsePreflightArgs([
      '--preflight',
      '--output',
      'tmp/preflight.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--google-sheets-storage-state',
      'tmp/google-state.json',
      '--ready-timeout-ms',
      '90000',
    ])

    expect(args).toMatchObject({
      corpusId: 'wide-mixed-250k',
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      headless: true,
      microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      readyTimeoutMs: 90000,
    })
    expect(args?.outputPath?.endsWith('/tmp/preflight.json')).toBe(true)
    expect(args?.googleSheetsStorageStatePath?.endsWith('/tmp/google-state.json')).toBe(true)
  })

  it('allows same-corpus preflight for one incumbent while diagnosing access setup', () => {
    const args = parsePreflightArgs([
      '--preflight',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--headed',
    ])

    expect(args).toMatchObject({
      googleSheetsUrl: null,
      headless: false,
      microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
    })
  })

  it('rejects same-corpus preflight with no incumbent URL', () => {
    expect(() => parsePreflightArgs(['--preflight'])).toThrow('Same-corpus preflight requires')
  })

  it('keeps incumbent preflight blockers structured before failing readiness', () => {
    expect(() =>
      assertSameCorpusPreflightReady({
        mode: 'preflight',
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        requiredProductCount: 2,
        checkedProductCount: 1,
        readyProductCount: 0,
        blockedProductCount: 1,
        allCheckedProductsReady: false,
        products: [
          {
            product: 'google-sheets',
            source: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
            finalUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
            title: 'Google Sheets',
            status: 'blocked',
            blocker:
              'Cannot preflight same-corpus editable workloads on google-sheets: Google Sheets page is not authenticated; provide an authenticated storage state.',
            corpusVerification: null,
            editableMutationProof: null,
            limitations: ['storage state was not provided'],
          },
        ],
      }),
    ).toThrow('Google Sheets page is not authenticated')
  })

  it('parses XLSX emission mode for same-corpus setup', () => {
    const args = parseEmitXlsxArgs(['--emit-xlsx', 'tmp/ui-corpus', '--corpus', 'wide-mixed-variable-250k'])

    expect(args).toMatchObject({
      check: false,
      corpusId: 'wide-mixed-variable-250k',
    })
    expect(args?.targetDirectory.endsWith('/tmp/ui-corpus')).toBe(true)
  })

  it('parses checked XLSX fixture mode', () => {
    const args = parseEmitXlsxArgs(['--emit-xlsx', 'packages/benchmarks/baselines/ui-same-corpus', '--check'])

    expect(args).toMatchObject({
      check: true,
      corpusId: 'wide-mixed-250k',
    })
    expect(args?.targetDirectory.endsWith('/packages/benchmarks/baselines/ui-same-corpus')).toBe(true)
  })

  it('rejects XLSX emission mode when the next flag would be consumed as the directory', () => {
    expect(() => parseEmitXlsxArgs(['--emit-xlsx', '--check'])).toThrow('Missing directory after --emit-xlsx')
  })

  it('builds deterministic literal-cell fingerprints for same-corpus verification', () => {
    const corpus = buildWorkbookBenchmarkCorpus('wide-mixed-250k')
    const fingerprint = buildSameCorpusFingerprint(corpus)

    expect(fingerprint).toMatchObject({
      sheetName: 'WideGrid',
      materializedCells: 250000,
    })
    expect(fingerprint.checkedCells.length).toBeGreaterThanOrEqual(3)
    expect(fingerprint.checkedCells[0]).toEqual({ address: 'A1', expected: 'metric-1' })
  })

  it('verifies same-corpus XLSX bytes before accepting external timing evidence', () => {
    const corpus = buildWorkbookBenchmarkCorpus('wide-mixed-250k')
    const verification = verifyXlsxCorpusFingerprint(Buffer.from(exportXlsx(corpus.snapshot)), corpus, 'microsoft-excel-web-source-xlsx')

    expect(verification).toMatchObject({
      verified: true,
      method: 'microsoft-excel-web-source-xlsx',
      sheetName: 'WideGrid',
      materializedCells: 250000,
      sourceWorkbookSha256: buildSameCorpusFingerprint(corpus).corpusFingerprint.snapshotSha256,
    })
    expect(verification.checkedCells.length).toBeGreaterThanOrEqual(3)
    expect(verification.checkedCells.every((cell) => cell.expected === cell.actual)).toBe(true)
  })

  it('rejects operation-only measurements before writing a same-corpus capture', async () => {
    await expect(
      collectSameCorpusProductMeasurements(
        {
          biligUrl: 'http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k',
          googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
          microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
        },
        async (product, url) => ({
          product,
          source: url,
          operationResponseMsSamples: [10, 11, 12],
          ...(product === 'bilig' ? { authoritativeRenderProofMsSamples: [14, 15, 16] } : {}),
          postOperationFrameMsSamples: [8, 9, 10],
          corpusVerification: {
            verified: true,
            method:
              product === 'bilig'
                ? 'bilig-benchmark-state'
                : product === 'google-sheets'
                  ? 'google-sheets-xlsx-export'
                  : 'microsoft-excel-web-source-xlsx',
            sheetName: 'WideGrid',
            materializedCells: 250000,
            checkedCells: [],
          },
          limitations: [],
        }),
      ),
    ).rejects.toThrow('same-corpus UI measurement for bilig is missing scroll-event response samples')
  })

  it('allows operation-only measurements for non-scroll same-corpus workloads', async () => {
    const measuredWorkloads: string[] = []
    const measurements = await collectSameCorpusProductMeasurements(
      {
        biligUrl: 'http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k',
        googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      },
      async (product, url, workload) => {
        measuredWorkloads.push(workload)
        return {
          product,
          source: url,
          operationResponseMsSamples: [10, 11, 12],
          ...(product === 'bilig' ? { authoritativeRenderProofMsSamples: [14, 15, 16] } : {}),
          postOperationFrameMsSamples: [8, 9, 10],
          corpusVerification: {
            verified: true,
            method:
              product === 'bilig'
                ? 'bilig-benchmark-state'
                : product === 'google-sheets'
                  ? 'google-sheets-xlsx-export'
                  : 'microsoft-excel-web-source-xlsx',
            sheetName: 'WideGrid',
            materializedCells: 250000,
            checkedCells: [],
          },
          limitations: [],
        }
      },
      'edit-visible-cell',
    )

    expect(measuredWorkloads).toEqual(['edit-visible-cell', 'edit-visible-cell', 'edit-visible-cell'])
    expect(measurements.bilig.source).toBe('http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k')
    expect(measurements.googleSheets.scrollEventResponseMsSamples).toBeUndefined()
  })

  it('declares the fixed same-corpus workload suite in capture order', () => {
    expect(requiredUiResponsivenessSameCorpusWorkloads).toEqual([
      'open-workbook',
      'select-cell',
      'edit-visible-cell',
      'scroll-vertical',
      'scroll-horizontal',
      'jump-deep-row',
      'formula-edit',
      'fill-format-change',
      'wide-sheet-navigation',
    ])
  })

  it('uses grid keyboard parity for same-corpus non-scroll Bilig workloads', () => {
    expect(sameCorpusKeyboardOperations('select-cell', 0, 'darwin')).toEqual([{ kind: 'press', key: 'ArrowRight' }])
    expect(sameCorpusKeyboardOperations('jump-deep-row', 0, 'darwin')).toEqual([{ kind: 'press', key: 'Meta+ArrowDown' }])
    expect(sameCorpusKeyboardOperations('fill-format-change', 0, 'linux')).toEqual([])
    expect(sameCorpusFillColorSwatchLabel(0)).toBe('light cornflower blue 3')
    expect(sameCorpusFillColorSwatchLabel(1)).toBe('theme green')
    expect(sameCorpusKeyboardOperations('formula-edit', 1, 'darwin')).toEqual([
      { kind: 'type', text: sameCorpusFormulaEditFormula(1) },
      { kind: 'press', key: 'Enter' },
    ])
    expect(sameCorpusKeyboardOperations('edit-visible-cell', 2, 'darwin')).toEqual([
      { kind: 'type', text: sameCorpusEditVisibleCellValue(2) },
      { kind: 'press', key: 'Enter' },
    ])
    expect(sameCorpusWorkbookRestoreOperations('edit-visible-cell', 'darwin')).toEqual([{ kind: 'press', key: 'Meta+Z' }])
    expect(sameCorpusWorkbookRestoreOperations('formula-edit', 'linux')).toEqual([{ kind: 'press', key: 'Control+Z' }])
    expect(sameCorpusWorkbookRestoreOperations('fill-format-change', 'darwin')).toEqual([{ kind: 'press', key: 'Meta+Z' }])
    expect(sameCorpusWorkbookRestoreOperations('select-cell', 'darwin')).toEqual([])
    expect(sameCorpusWorkbookRestoreOperations('scroll-vertical', 'darwin')).toEqual([])
    expect(sameCorpusWorkloadMutatesWorkbook('edit-visible-cell')).toBe(true)
    expect(sameCorpusWorkloadMutatesWorkbook('formula-edit')).toBe(true)
    expect(sameCorpusWorkloadMutatesWorkbook('fill-format-change')).toBe(true)
    expect(sameCorpusWorkloadMutatesWorkbook('jump-deep-row')).toBe(false)
    expect(sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0)).toBe('C5')
    expect(sameCorpusMutationTargetRangeForSample('formula-edit', 1)).toBe('D6')
    expect(sameCorpusMutationTargetRangeForSample('fill-format-change', 2)).toBe('E7')
  })

  it('routes non-scroll timings through the browser-visible response barrier', async () => {
    const events: string[] = []
    // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- This test supplies the minimal Playwright Page surface used by the keyboard workload path.
    const page = {
      keyboard: {
        press: async (key: string) => {
          events.push(`press:${key}`)
        },
        type: async (text: string) => {
          events.push(`type:${text}`)
        },
      },
    } as unknown as Page
    const sample = await measureProductWorkload({
      page,
      product: 'bilig',
      captureArgs: parseCaptureArgs([
        '--output',
        'tmp/ui-capture.json',
        '--google-sheets-url',
        'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        '--allow-incomplete-evidence',
      ]),
      workload: 'select-cell',
      sampleIndex: 0,
      loadToReadyMs: 999,
      hooks: {
        measureVisibleScrollResponseWithRetries: async () => {
          throw new Error('scroll hook should not be used for select-cell')
        },
        measureVisibleNonScrollResponse: async (_page, product, workload, sampleIndex, runOperation) => {
          events.push(`visible:${product}:${workload}:${String(sampleIndex)}`)
          await runOperation()
          return { operationResponseMs: 42, postOperationFrameMs: 7 }
        },
        movePointerToProductViewport: async () => {
          events.push('pointer')
        },
      },
    })

    expect(sample).toEqual({ operationResponseMs: 42, postOperationFrameMs: 7 })
    expect(events).toEqual(['pointer', 'visible:bilig:select-cell:0', 'press:ArrowRight'])
  })

  it('requires Bilig interaction-visible state movement instead of arbitrary screenshot noise', () => {
    const before: VisibleNonScrollResponseSignature = {
      biligInteractionVisibleToken: 'selection-revision-1',
      product: 'bilig',
      screenshotSignature: 'screenshot-a',
      workload: 'select-cell',
    }

    expect(
      visibleNonScrollResponseChanged(before, {
        ...before,
        screenshotSignature: 'screenshot-b',
      }),
    ).toBe(false)
    expect(
      visibleNonScrollResponseChanged(before, {
        ...before,
        biligInteractionVisibleToken: 'selection-revision-2',
      }),
    ).toBe(true)
    expect(
      visibleNonScrollResponseChanged(
        {
          biligInteractionVisibleToken: null,
          product: 'google-sheets',
          screenshotSignature: 'screenshot-a',
          workload: 'select-cell',
        },
        {
          biligInteractionVisibleToken: null,
          product: 'google-sheets',
          screenshotSignature: 'screenshot-b',
          workload: 'select-cell',
        },
      ),
    ).toBe(true)
  })

  it('measures Bilig visible response after the operation starts instead of after authoritative completion', async () => {
    let operationCompleted = false
    let evaluateCount = 0
    const surface = {
      dpr: 1,
      editorVisibleRevision: null,
      fallback: null,
      formulaVisibleRevision: null,
      gridAuthoritativeRenderRevision: '0',
      gridEditorVisibleRevision: 'grid-editor-before',
      gridHeight: 480,
      gridInteractionVisibleRevision: 'grid-before',
      gridLocalRenderRevision: '1',
      gridProjectedRenderRevision: '2',
      gridSelectionVisibleRevision: 'grid-selection-before',
      gridWidth: 720,
      nativeRectCount: 4,
      nativeRectLayerMounted: true,
      nativeTextLayerMounted: true,
      nativeTextRunCount: 3,
      typeGpu: {
        authoritativeRenderRevision: '0',
        backendStatus: 'ready',
        currentContentSignature: 'content-current',
        currentRectSignature: 'rect-current',
        currentSceneEpochSignature: 'epoch-current',
        currentSceneOwnershipSignature: 'scene-current',
        currentSelectionRevision: 'selection-current',
        currentSemanticMutationRevision: 'semantic-current',
        currentTextSignature: 'text-current',
        currentViewportRevision: 'viewport-current',
        currentWorkbookRevision: 'workbook-current',
        frameProofStatus: 'presented',
        headerPaneCount: 3,
        localRenderRevision: '1',
        mode: 'typegpu-v3',
        pixelHeight: 480,
        pixelWidth: 720,
        presentedContentSignature: 'content-current',
        presentedRectSignature: 'rect-current',
        presentedSelectionRevision: 'selection-current',
        presentedSemanticMutationRevision: 'semantic-current',
        presentedTextSignature: 'text-current',
        presentedViewportRevision: 'viewport-current',
        presentedWorkbookRevision: 'workbook-current',
        projectedRenderRevision: '2',
        tilePaneCount: 6,
        tileSceneRevision: '2',
        visibleLocalRenderRevision: '1',
        visibleProjectedRenderRevision: '2',
      },
    }
    // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion -- This test supplies the minimal Playwright Page surface used by the visible-response barrier.
    const page = {
      evaluate: async (_callback: unknown, arg?: unknown) => {
        evaluateCount += 1
        if (evaluateCount === 1) {
          return surface
        }
        if (typeof arg === 'number') {
          return Array.from({ length: arg }, () => 1)
        }
        return undefined
      },
      waitForFunction: async () => {
        expect(operationCompleted).toBe(false)
      },
    } as unknown as Page

    let finishOperation: (() => void) | null = null
    const operationPromise = new Promise<void>((resolve) => {
      finishOperation = () => {
        operationCompleted = true
        resolve()
      }
    })
    const samplePromise = measureVisibleNonScrollResponse(page, 'bilig', 'edit-visible-cell', 0, async () => {
      await operationPromise
    })
    await new Promise<void>((resolve) => setTimeout(resolve, 0))
    expect(operationCompleted).toBe(false)
    finishOperation?.()
    const sample = await samplePromise

    expect(operationCompleted).toBe(true)
    expect(sample.operationResponseMs).toBeLessThan(50)
    expect(sample.operationResponseProof).toBe('visible-non-scroll-response')
  })

  it('keys Bilig non-scroll response timing from local visible revisions, not authoritative presented proof', () => {
    const surface = {
      dpr: 2,
      editorVisibleRevision: 'cell-editor:WideGrid!F6:typed:5:5',
      fallback: null,
      formulaVisibleRevision: null,
      gridAuthoritativeRenderRevision: 'authoritative-1',
      gridEditorVisibleRevision: 'grid-editor-current',
      gridHeight: 480,
      gridInteractionVisibleRevision: 'grid-interaction-current',
      gridLocalRenderRevision: 'local-7',
      gridProjectedRenderRevision: 'projected-9',
      gridSelectionVisibleRevision: 'grid-selection-current',
      gridWidth: 720,
      nativeRectCount: 4,
      nativeRectLayerMounted: true,
      nativeTextLayerMounted: true,
      nativeTextRunCount: 3,
      typeGpu: {
        authoritativeRenderRevision: 'authoritative-1',
        backendStatus: 'ready',
        currentContentSignature: 'content-current',
        currentRectSignature: 'rect-current',
        currentSceneEpochSignature: 'epoch-current',
        currentSceneOwnershipSignature: 'scene-current',
        currentSelectionRevision: 'selection-current',
        currentSemanticMutationRevision: 'semantic-current',
        currentTextSignature: 'text-current',
        currentViewportRevision: 'viewport-current',
        currentWorkbookRevision: 'workbook-current',
        frameProofStatus: 'pending',
        headerPaneCount: 3,
        localRenderRevision: 'local-7',
        mode: 'typegpu-v3',
        pixelHeight: 960,
        pixelWidth: 1440,
        presentedContentSignature: 'content-stale',
        presentedRectSignature: 'rect-stale',
        presentedSelectionRevision: 'selection-stale',
        presentedSemanticMutationRevision: 'semantic-stale',
        presentedTextSignature: 'text-stale',
        presentedViewportRevision: 'viewport-stale',
        presentedWorkbookRevision: 'workbook-stale',
        projectedRenderRevision: 'projected-9',
        tilePaneCount: 6,
        tileSceneRevision: 'tile-9',
        visibleLocalRenderRevision: 'local-7',
        visibleProjectedRenderRevision: 'projected-9',
      },
    }

    expect(biligInteractionVisibleResponseToken(surface, 'select-cell')).toBe('grid-selection-current')
    expect(biligInteractionVisibleResponseToken(surface, 'jump-deep-row')).toBe('grid-selection-current')
    expect(biligInteractionVisibleResponseToken(surface, 'edit-visible-cell')).toContain('cell-editor:WideGrid!F6:typed')
    expect(biligInteractionVisibleResponseToken(surface, 'edit-visible-cell')).not.toContain('semantic-stale')
  })

  it('does not charge screenshot capture to Bilig visible-response timing once TypeGPU tokens exist', () => {
    expect(visibleNonScrollResponseNeedsScreenshot('bilig', 'selection-revision-2')).toBe(false)
    expect(visibleNonScrollResponseNeedsScreenshot('bilig', null)).toBe(true)
    expect(visibleNonScrollResponseNeedsScreenshot('google-sheets')).toBe(true)
    expect(visibleNonScrollResponseNeedsScreenshot('microsoft-excel-web')).toBe(true)
  })

  it('requires visual proof for Bilig and Google Sheets in each same-corpus scenario', () => {
    const proof = buildCaptureScenarioProof({
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export'),
      microsoftExcelWeb: sameCorpusCaptureMeasurement('microsoft-excel-web', 'microsoft-excel-web-source-xlsx'),
      visualProofs: [
        sameCorpusVisualProof('bilig', 'typegpu-visible-canvas'),
        sameCorpusVisualProof('google-sheets', 'google-sheets-visible-grid'),
      ],
    })

    expect(proof.screenshotProof).toMatchObject({
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      missingProducts: [],
    })
    expect(proof.pixelGridProof).toMatchObject({
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      missingProducts: [],
    })
    expect(proof.semanticUiProof).toMatchObject({
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      missingProducts: [],
    })
  })

  it('rejects mutating semantic UI proof that lacks target readback and restore evidence', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'edit-visible-cell',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'edit-visible-cell'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'edit-visible-cell'),
      visualProofs: [
        sameCorpusVisualProof('bilig', 'typegpu-visible-canvas', 'same-corpus-wide-mixed-250k-edit-visible-cell'),
        sameCorpusVisualProof('google-sheets', 'google-sheets-visible-grid', 'same-corpus-wide-mixed-250k-edit-visible-cell'),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig', 'google-sheets'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining(['semantic UI mutation target proof for edit-visible-cell covers 0/3 samples']),
        }),
        expect.objectContaining({
          product: 'google-sheets',
          invalidReasons: expect.arrayContaining(['semantic UI mutation target proof for edit-visible-cell covers 0/3 samples']),
        }),
      ]),
    )
  })

  it('rejects mutating semantic UI proof with stale target values, missing restore, and revision drift', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'edit-visible-cell',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'edit-visible-cell'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'edit-visible-cell'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
          corruptFirstMutationTargetProof,
        ),
        sameCorpusVisualProofWithMutationProofs(
          'google-sheets',
          'google-sheets-visible-grid',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
          corruptFirstMutationTargetProof,
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig', 'google-sheets'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for edit-visible-cell is missing the target range',
            'semantic UI mutation target proof for edit-visible-cell did not prove undo restore',
            'semantic UI mutation target proof for edit-visible-cell has unverified undo restore status',
            'semantic UI mutation target proof for edit-visible-cell is missing authoritative readback revision',
            'semantic UI mutation target proof for edit-visible-cell is missing visible render revision',
            'semantic UI mutation target proof for edit-visible-cell is missing screenshot artifact path',
            'semantic UI mutation target proof for edit-visible-cell did not prove the intended committed target value',
          ]),
        }),
      ]),
    )
  })

  it('rejects Bilig mutation target proof backed only by visible editor text', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'edit-visible-cell',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'edit-visible-cell'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'edit-visible-cell'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
          forceVisibleEditorReadbackSource,
        ),
        sameCorpusVisualProof(
          'google-sheets',
          'google-sheets-visible-grid',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for edit-visible-cell used visible editor text instead of Bilig authoritative range readback',
          ]),
        }),
      ]),
    )
  })

  it('rejects mutation target proof whose target range does not match the rendered selection', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'fill-format-change',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'fill-format-change'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'fill-format-change'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-fill-format-change',
          'fill-format-change',
          moveMutationTargetAwayFromRenderedSelection,
        ),
        sameCorpusVisualProofWithMutationProofs(
          'google-sheets',
          'google-sheets-visible-grid',
          'same-corpus-wide-mixed-250k-fill-format-change',
          'fill-format-change',
          moveMutationTargetAwayFromRenderedSelection,
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig', 'google-sheets'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for fill-format-change target range does not match the rendered selection',
          ]),
        }),
        expect.objectContaining({
          product: 'google-sheets',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for fill-format-change target range does not match the rendered selection',
          ]),
        }),
      ]),
    )
  })

  it('rejects Bilig mutation target proof whose rendered readback does not match authoritative readback', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'edit-visible-cell',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'edit-visible-cell'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'edit-visible-cell'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
          driftVisibleTargetReadback,
        ),
        sameCorpusVisualProof(
          'google-sheets',
          'google-sheets-visible-grid',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for edit-visible-cell rendered readback does not match Bilig authoritative range readback',
          ]),
        }),
      ]),
    )
  })

  it('rejects fill-format mutation target proof without a rendered post-mutation fill color', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'fill-format-change',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'fill-format-change'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'fill-format-change'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-fill-format-change',
          'fill-format-change',
          removeRenderedPostMutationFill,
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig', 'google-sheets'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for fill-format-change is missing rendered post-mutation fill color',
          ]),
        }),
      ]),
    )
  })

  it('rejects fill-format mutation target proof whose rendered color does not match the intended swatch', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'fill-format-change',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'fill-format-change'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'fill-format-change'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-fill-format-change',
          'fill-format-change',
          driftIntendedFillColor,
        ),
      ],
    })

    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for fill-format-change did not prove the intended fill color',
          ]),
        }),
      ]),
    )
  })

  it('rejects mutation target proof that does not prove the intended operation payload', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'formula-edit',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'formula-edit'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'formula-edit'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-formula-edit',
          'formula-edit',
          driftIntendedFormulaPayload,
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig', 'google-sheets'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for formula-edit did not prove the intended edited formula',
          ]),
        }),
      ]),
    )
  })

  it('rejects mutation target screenshots whose embedded identity drifts from the timed sample', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'edit-visible-cell',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'edit-visible-cell'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'edit-visible-cell'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
          driftMutationTargetScreenshotIdentity,
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig', 'google-sheets'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for edit-visible-cell has mismatched before screenshot product',
            'semantic UI mutation target proof for edit-visible-cell has mismatched before screenshot workload',
            'semantic UI mutation target proof for edit-visible-cell has mismatched before screenshot sample',
            'semantic UI mutation target proof for edit-visible-cell has mismatched before screenshot sheet identity',
          ]),
        }),
      ]),
    )
  })

  it('rejects mutation target proof whose top-level product drifts from the visual product', () => {
    const proof = buildCaptureScenarioProof({
      workload: 'edit-visible-cell',
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', 'edit-visible-cell'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', 'edit-visible-cell'),
      visualProofs: [
        sameCorpusVisualProofWithMutationProofs(
          'bilig',
          'typegpu-visible-canvas',
          'same-corpus-wide-mixed-250k-edit-visible-cell',
          'edit-visible-cell',
          driftMutationTargetProofProduct,
        ),
      ],
    })

    expect(proof.semanticUiProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig', 'google-sheets'],
    })
    expect(proof.semanticUiProof.productVerdicts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          product: 'bilig',
          invalidReasons: expect.arrayContaining(['semantic UI mutation target proof for edit-visible-cell has mismatched product']),
        }),
      ]),
    )
  })

  it('normalizes same-corpus mutation target selections to an explicit cell range', () => {
    expect(normalizeSameCorpusMutationTargetSelection('WideGrid!$C$5:$D$7', 'WideGrid')).toEqual({
      endAddress: 'D7',
      sheetName: 'WideGrid',
      sheetId: null,
      startAddress: 'C5',
      targetRange: 'C5:D7',
    })
    expect(() => normalizeSameCorpusMutationTargetSelection('selected row', 'WideGrid')).toThrow(
      'Cannot derive same-corpus mutation target range from visible selection: selected row',
    )
  })

  it('selects Google Sheets target ranges through the visible name-box shortcut', async () => {
    const calls: string[] = []
    const page = fakeGoogleSheetsNameBoxPage(calls, 'A1')

    await selectGoogleSheetsTargetRange(page, 'C7')

    expect(calls).toEqual([
      `${testPrimaryShortcut()}+J`,
      'locator:#t-name-box',
      'locator:input.waffle-name-box',
      'locator:input[aria-label="Name box"]',
      'locator:[aria-label^="Name box"] input',
      'fill:C7',
      'press:Enter',
    ])
  })

  it('reads Google Sheets selected ranges through the visible name-box shortcut', async () => {
    const calls: string[] = []
    const page = fakeGoogleSheetsNameBoxPage(calls, 'D9')

    await expect(readGoogleSheetsNameBoxSelection(page)).resolves.toBe('D9')
    expect(calls).toEqual([`${testPrimaryShortcut()}+J`, 'locator:#t-name-box', 'inputValue', 'keyboard:Escape'])
  })

  it('writes same-corpus capture artifacts with a fresh run manifest', () => {
    const scenarioProof = buildCaptureScenarioProof({
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export'),
      visualProofs: [
        sameCorpusVisualProof('bilig', 'typegpu-visible-canvas'),
        sameCorpusVisualProof('google-sheets', 'google-sheets-visible-grid'),
      ],
    })
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: [
        {
          id: 'same-corpus-wide-mixed-250k-open-workbook',
          corpusCaseId: 'wide-mixed-250k',
          materializedCells: 250_000,
          workload: 'open-workbook',
          ...sameCorpusScenarioCaseFields(scenarioProof),
          scenarioProof,
          bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state'),
          googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export'),
        },
      ],
    })

    expect(capture.runManifest).toMatchObject({
      artifactGenerator: 'scripts/capture-ui-responsiveness-same-corpus.ts',
      biligAuthoritativeRenderProofCaseCount: 1,
      caseCount: 1,
      contractVersion: 'same-corpus-ui-v8',
      currentContractEvidenceComplete: false,
      googleSheetsTenXRequirementSatisfied: false,
      requiredProducts: ['bilig', 'google-sheets'],
      sampleCount: 3,
      scenarioSummaryFieldCaseCount: 1,
      strictRenderedGridProofCaseCount: 1,
      semanticUiProofCaseCount: 1,
      requiredMutationTargetProofCaseCount: 3,
      mutationTargetProofCaseCount: 0,
      requiredMutationTargetProofSampleCount: 18,
      mutationTargetProofSampleCount: 0,
      requiredCommittedTargetProofTimingCaseCount: 3,
      committedTargetProofTimingCaseCount: 0,
      requiredCommittedTargetProofTimingSampleCount: 18,
      committedTargetProofTimingSampleCount: 0,
      tenXMeanAndP95CaseCount: 0,
    })
    expect(capture.runManifest.captureRunSignature).toMatch(/^[a-f0-9]{64}$/u)
    expect(capture.runManifest.invalidReasons).toContain(
      'missing required workloads: select-cell, edit-visible-cell, scroll-vertical, scroll-horizontal, jump-deep-row, formula-edit, fill-format-change, wide-sheet-navigation',
    )
    expect(capture.runManifest.invalidReasons).toContain('Bilig production runtime proof covers 0/9 cases')
    expect(capture.runManifest.invalidReasons).toContain('mutation target proof covers 0/3 mutating cases')
    expect(capture.runManifest.invalidReasons).toContain('mutation target proof covers 0/18 required per-sample product proofs')
    expect(capture.runManifest.invalidReasons).toContain('committed target proof timing covers 0/3 mutating cases')
    expect(capture.runManifest.invalidReasons).toContain('committed target proof timing covers 0/18 required per-sample product timings')
    expect(capture.runManifest.invalidReasons).toContain('not every required workload is 10x against Google Sheets')
    expect(capture.cases[0]).toMatchObject({
      biligMeanMs: scenarioProof.biligMeanMs,
      biligP95Ms: scenarioProof.biligP95Ms,
      googleMeanMs: scenarioProof.googleMeanMs,
      googleP95Ms: scenarioProof.googleP95Ms,
      meanRatio: scenarioProof.meanRatio,
      p95Ratio: scenarioProof.p95Ratio,
      screenshotProof: scenarioProof.screenshotProof,
      pixelGridProof: scenarioProof.pixelGridProof,
    })
    expect(parseSameCorpusCapture(capture).runManifest.captureRunSignature).toBe(capture.runManifest.captureRunSignature)
  })

  it('counts Bilig production-runtime proof in same-corpus capture manifests', () => {
    const scenarioProof = buildCaptureScenarioProof({
      bilig: {
        ...sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state'),
        biligRuntimeProof: sameCorpusBiligRuntimeProof('production'),
      },
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export'),
      visualProofs: [
        sameCorpusVisualProof('bilig', 'typegpu-visible-canvas'),
        sameCorpusVisualProof('google-sheets', 'google-sheets-visible-grid'),
      ],
    })
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: [
        {
          id: 'same-corpus-wide-mixed-250k-open-workbook',
          corpusCaseId: 'wide-mixed-250k',
          materializedCells: 250_000,
          workload: 'open-workbook',
          ...sameCorpusScenarioCaseFields(scenarioProof),
          scenarioProof,
          bilig: {
            ...sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state'),
            biligRuntimeProof: sameCorpusBiligRuntimeProof('production'),
          },
          googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export'),
        },
      ],
    })

    expect(capture.runManifest.biligProductionRuntimeProofCaseCount).toBe(1)
    expect(capture.runManifest.scenarioSummaryFieldCaseCount).toBe(1)
    expect(capture.runManifest.biligAuthoritativeRenderProofCaseCount).toBe(1)
    expect(capture.runManifest.semanticUiProofCaseCount).toBe(1)
    expect(capture.runManifest.invalidReasons).toContain('Bilig production runtime proof covers 1/9 cases')
    expect(parseSameCorpusCapture(capture).cases[0]?.bilig.biligRuntimeProof?.verified).toBe(true)
  })

  it('requires Bilig authoritative rendered-proof timing in current same-corpus manifests', () => {
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: requiredUiResponsivenessSameCorpusWorkloads.map((workload) =>
        sameCorpusCaptureCase({
          workload,
          biligRuntimeProof: sameCorpusBiligRuntimeProof('production'),
          includeAuthoritativeRenderProofTiming: false,
        }),
      ),
    })

    expect(capture.runManifest.currentContractEvidenceComplete).toBe(false)
    expect(capture.runManifest.biligAuthoritativeRenderProofCaseCount).toBe(0)
    expect(capture.runManifest.invalidReasons).toContain('Bilig authoritative render proof timing covers 0/9 cases')
  })

  it('rejects capture artifacts with incomplete browser-visible evidence by default', () => {
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: [
        sameCorpusCaptureCase({
          workload: 'open-workbook',
          biligRuntimeProof: null,
        }),
      ],
    })

    expect(() => assertSameCorpusCaptureEvidenceReady(capture)).toThrow(
      /Same-corpus UI capture artifact is not valid claim-grade Google Sheets 10x evidence/u,
    )
    expect(() => assertSameCorpusCaptureEvidenceReady(capture)).toThrow(/Bilig production runtime proof covers 0\/9 cases/u)
  })

  it('separates current-contract diagnostics from claim-grade Google Sheets 10x readiness', () => {
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: requiredUiResponsivenessSameCorpusWorkloads.map((workload) =>
        sameCorpusCaptureCase({
          workload,
          biligRuntimeProof: sameCorpusBiligRuntimeProof('production'),
        }),
      ),
    })

    expect(capture.runManifest.currentContractEvidenceComplete).toBe(true)
    expect(capture.runManifest.googleSheetsTenXRequirementSatisfied).toBe(false)
    expect(capture.runManifest.invalidReasons).toEqual(['not every required workload is 10x against Google Sheets'])
    expect(() => assertSameCorpusCaptureCurrentContractEvidenceReady(capture)).not.toThrow()
    expect(() => assertSameCorpusCaptureEvidenceReady(capture)).toThrow(/not every required workload is 10x against Google Sheets/u)
  })

  it('downgrades legacy Bilig canvas evidence that lacks strict rendered-frame proof', () => {
    const proof = buildCaptureScenarioProof({
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export'),
      visualProofs: [
        {
          ...sameCorpusVisualProof('bilig', 'typegpu-visible-canvas'),
          pixelGridProof: {
            product: 'bilig',
            captured: true,
            method: 'typegpu-visible-canvas',
            viewportPixelWidth: 1440,
            viewportPixelHeight: 900,
            evidence: ['mode=typegpu-v3', 'tilePaneCount=6', 'headerPaneCount=3'],
          },
        },
        sameCorpusVisualProof('google-sheets', 'google-sheets-visible-grid'),
      ],
    })

    expect(proof.pixelGridProof).toMatchObject({
      captured: false,
      missingProducts: ['bilig'],
    })
  })

  it('rejects each missing Bilig visible-frame proof field and revision mismatch', () => {
    const baseProof = sameCorpusVisualProof('bilig', 'typegpu-visible-canvas').pixelGridProof
    const baseEvidence = baseProof.evidence

    expect(isSameCorpusProductPixelGridProofComplete(baseProof)).toBe(true)

    const invalidEvidenceCases: readonly {
      readonly label: string
      readonly evidence: readonly string[]
    }[] = [
      ...[
        'pixelGridProofVersion',
        'pixelSampleSource',
        'visibleGridLinePixels',
        'verticalLineRuns',
        'horizontalLineRuns',
        'verticalLineCoverageBands',
        'horizontalLineCoverageBands',
        'largestVerticalLineGapPx',
        'largestHorizontalLineGapPx',
        'mode',
        'contractVersion',
        'backendStatus',
        'frameProofStatus',
        'frameProofSignature',
        'hasPresentedFrame',
        'hasPresentedVisibleFrame',
        'presentedFrameProofSignature',
        'currentSceneEpochSignature',
        'currentSceneOwnershipSignature',
        'presentedSceneEpochSignature',
        'presentedSceneOwnershipSignature',
        'currentWorkbookRevision',
        'presentedWorkbookRevision',
        'currentSemanticMutationRevision',
        'presentedSemanticMutationRevision',
        'currentViewportRevision',
        'presentedViewportRevision',
        'currentSelectionRevision',
        'presentedSelectionRevision',
        'currentFillHandleRevision',
        'presentedFillHandleRevision',
        'currentContentSignature',
        'presentedContentSignature',
        'currentTextRunCount',
        'presentedTextRunCount',
        'currentTextSignature',
        'presentedTextSignature',
        'currentRectCount',
        'presentedRectCount',
        'currentRectSignature',
        'presentedRectSignature',
        'tilePaneCount',
        'headerPaneCount',
        'presentedTilePaneCount',
        'presentedHeaderPaneCount',
        'expectedPixelWidth',
        'expectedPixelHeight',
        'canvasPixelWidth',
        'canvasPixelHeight',
        'canvasCoversViewport',
        'gridAuthoritativeRevision',
        'typeGpuAuthoritativeRevision',
        'visibleAuthoritativeRevision',
        'gridLocalRevision',
        'typeGpuLocalRevision',
        'visibleLocalRevision',
        'gridProjectedRevision',
        'typeGpuProjectedRevision',
        'visibleProjectedRevision',
        'tileSceneRevision',
        'visibleRenderRevision',
      ].map((key) => ({ label: `missing ${key}`, evidence: evidenceWithoutKey(baseEvidence, key) })),
      { label: 'wrong pixel proof contract', evidence: replaceEvidence(baseEvidence, 'pixelGridProofVersion', 'legacy-grid-pixels') },
      { label: 'wrong pixel proof source', evidence: replaceEvidence(baseEvidence, 'pixelSampleSource', 'dom-rectangle') },
      { label: 'too few gridline pixels', evidence: replaceEvidence(baseEvidence, 'visibleGridLinePixels', '23') },
      { label: 'too few vertical line runs', evidence: replaceEvidence(baseEvidence, 'verticalLineRuns', '2') },
      { label: 'too few horizontal line runs', evidence: replaceEvidence(baseEvidence, 'horizontalLineRuns', '2') },
      { label: 'too few vertical coverage bands', evidence: replaceEvidence(baseEvidence, 'verticalLineCoverageBands', '4') },
      { label: 'too few horizontal coverage bands', evidence: replaceEvidence(baseEvidence, 'horizontalLineCoverageBands', '4') },
      { label: 'large vertical grid gap', evidence: replaceEvidence(baseEvidence, 'largestVerticalLineGapPx', '385') },
      { label: 'large horizontal grid gap', evidence: replaceEvidence(baseEvidence, 'largestHorizontalLineGapPx', '161') },
      { label: 'wrong renderer mode', evidence: replaceEvidence(baseEvidence, 'mode', 'canvas2d') },
      { label: 'wrong render proof contract', evidence: replaceEvidence(baseEvidence, 'contractVersion', 'same-corpus-ui-v2') },
      { label: 'backend not ready', evidence: replaceEvidence(baseEvidence, 'backendStatus', 'pending') },
      { label: 'frame not presented', evidence: replaceEvidence(baseEvidence, 'frameProofStatus', 'pending') },
      { label: 'missing current frame presentation', evidence: replaceEvidence(baseEvidence, 'hasPresentedFrame', 'false') },
      {
        label: 'stale presented frame signature',
        evidence: replaceEvidence(baseEvidence, 'presentedFrameProofSignature', 'frame-stale'),
      },
      {
        label: 'stale presented scene ownership',
        evidence: replaceEvidence(baseEvidence, 'presentedSceneOwnershipSignature', 'scene-stale'),
      },
      {
        label: 'stale presented scene epoch',
        evidence: replaceEvidence(baseEvidence, 'presentedSceneEpochSignature', 'epoch-stale'),
      },
      {
        label: 'stale presented workbook revision',
        evidence: replaceEvidence(baseEvidence, 'presentedWorkbookRevision', 'workbook-stale'),
      },
      {
        label: 'stale presented semantic mutation revision',
        evidence: replaceEvidence(baseEvidence, 'presentedSemanticMutationRevision', 'semantic-stale'),
      },
      {
        label: 'stale presented viewport revision',
        evidence: replaceEvidence(baseEvidence, 'presentedViewportRevision', 'viewport-stale'),
      },
      {
        label: 'stale presented selection revision',
        evidence: replaceEvidence(baseEvidence, 'presentedSelectionRevision', 'selection-stale'),
      },
      {
        label: 'stale presented fill handle revision',
        evidence: replaceEvidence(baseEvidence, 'presentedFillHandleRevision', 'fill-stale'),
      },
      { label: 'visible frame not presented', evidence: replaceEvidence(baseEvidence, 'hasPresentedVisibleFrame', 'false') },
      {
        label: 'stale presented content signature',
        evidence: replaceEvidence(baseEvidence, 'presentedContentSignature', 'content-stale'),
      },
      { label: 'stale presented text signature', evidence: replaceEvidence(baseEvidence, 'presentedTextSignature', 'text-stale') },
      { label: 'stale presented rect signature', evidence: replaceEvidence(baseEvidence, 'presentedRectSignature', 'rect-stale') },
      { label: 'stale presented text run count', evidence: replaceEvidence(baseEvidence, 'presentedTextRunCount', '11') },
      { label: 'stale presented rect count', evidence: replaceEvidence(baseEvidence, 'presentedRectCount', '87') },
      { label: 'empty current rect payload', evidence: replaceEvidence(baseEvidence, 'currentRectCount', '0') },
      { label: 'empty presented rect payload', evidence: replaceEvidence(baseEvidence, 'presentedRectCount', '0') },
      { label: 'empty tile panes', evidence: replaceEvidence(baseEvidence, 'tilePaneCount', '0') },
      { label: 'empty header panes', evidence: replaceEvidence(baseEvidence, 'headerPaneCount', '0') },
      { label: 'empty presented tile panes', evidence: replaceEvidence(baseEvidence, 'presentedTilePaneCount', '0') },
      { label: 'empty presented header panes', evidence: replaceEvidence(baseEvidence, 'presentedHeaderPaneCount', '0') },
      { label: 'partial presented tile panes', evidence: replaceEvidence(baseEvidence, 'presentedTilePaneCount', '5') },
      { label: 'partial presented header panes', evidence: replaceEvidence(baseEvidence, 'presentedHeaderPaneCount', '2') },
      { label: 'undersized canvas width', evidence: replaceEvidence(baseEvidence, 'canvasPixelWidth', '1437') },
      { label: 'undersized canvas height', evidence: replaceEvidence(baseEvidence, 'canvasPixelHeight', '897') },
      { label: 'canvas does not cover viewport', evidence: replaceEvidence(baseEvidence, 'canvasCoversViewport', 'false') },
      {
        label: 'TypeGPU authoritative revision differs from grid',
        evidence: replaceEvidence(baseEvidence, 'typeGpuAuthoritativeRevision', 'rev-2'),
      },
      {
        label: 'visible authoritative revision differs from grid',
        evidence: replaceEvidence(baseEvidence, 'visibleAuthoritativeRevision', 'rev-2'),
      },
      { label: 'TypeGPU local revision differs from grid', evidence: replaceEvidence(baseEvidence, 'typeGpuLocalRevision', 'rev-local-1') },
      { label: 'visible local revision differs from grid', evidence: replaceEvidence(baseEvidence, 'visibleLocalRevision', 'rev-local-1') },
      { label: 'TypeGPU revision differs from grid', evidence: replaceEvidence(baseEvidence, 'typeGpuProjectedRevision', 'rev-2') },
      { label: 'visible revision differs from grid', evidence: replaceEvidence(baseEvidence, 'visibleProjectedRevision', 'rev-2') },
      {
        label: 'visible render revision differs from tile scene',
        evidence: replaceEvidence(baseEvidence, 'visibleRenderRevision', 'scene-6'),
      },
    ]

    for (const entry of invalidEvidenceCases) {
      expect(
        isSameCorpusProductPixelGridProofComplete({
          ...baseProof,
          evidence: entry.evidence,
        }),
        entry.label,
      ).toBe(false)
    }
  })

  it('downgrades incumbent grid evidence that only proves a large DOM rectangle exists', () => {
    const proof = buildCaptureScenarioProof({
      bilig: sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state'),
      googleSheets: sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export'),
      visualProofs: [
        sameCorpusVisualProof('bilig', 'typegpu-visible-canvas'),
        {
          ...sameCorpusVisualProof('google-sheets', 'google-sheets-visible-grid'),
          pixelGridProof: {
            product: 'google-sheets',
            captured: true,
            method: 'google-sheets-visible-grid',
            viewportPixelWidth: 1440,
            viewportPixelHeight: 900,
            evidence: ['selector=.grid-scrollable-wrapper', 'cssWidth=720', 'cssHeight=450'],
          },
        },
      ],
    })

    expect(proof.pixelGridProof).toMatchObject({
      captured: false,
      missingProducts: ['google-sheets'],
    })
  })

  it('rejects read-only incumbent edit surfaces before timing same-corpus edits', () => {
    expect(
      incumbentEditableWorkloadBlocker(
        'google-sheets',
        'https://docs.google.com/spreadsheets/d/example/edit',
        'File Edit View Comment only Share',
      ),
    ).toContain('read-only')
    expect(
      incumbentEditableWorkloadBlocker(
        'microsoft-excel-web',
        'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
        'Excel workbook viewer',
      ),
    ).toContain('read-only Office viewer')
    expect(
      incumbentEditableWorkloadBlocker(
        'google-sheets',
        'https://docs.google.com/spreadsheets/d/example/edit',
        'File Edit Insert Format Data Tools Extensions Help Share Sign in',
      ),
    ).toContain('authenticated')
    expect(
      incumbentEditableWorkloadBlocker(
        'google-sheets',
        'https://docs.google.com/spreadsheets/d/example/edit',
        'File Edit Insert Format Data Tools Extensions Help',
      ),
    ).toBeNull()
  })

  it('launches same-corpus capture browsers with WebGPU enabled for TypeGPU proof', () => {
    expect(sameCorpusChromiumLaunchOptions(true)).toEqual({
      args: ['--enable-unsafe-webgpu', '--ignore-gpu-blocklist'],
      headless: true,
    })
  })

  it('keeps product-specific scroll probe targets explicit', () => {
    expect(sameCorpusScrollProbeSelectorsForProduct('bilig')).toEqual(['[data-testid="grid-scroll-viewport"]'])
    expect(sameCorpusScrollProbeSelectorsForProduct('google-sheets')).toEqual([
      '.native-scrollbar-y',
      '.native-scrollbar-x',
      '.grid-scrollable-wrapper',
    ])
    expect(sameCorpusScrollProbeSelectorsForProduct('microsoft-excel-web')).toEqual(['.ewr-grdcontarea-grid'])
  })

  it('parses storage-state bootstrap mode for authenticated capture', () => {
    const args = parseSaveStorageStateArgs([
      '--save-storage-state',
      'tmp/google-state.json',
      '--auth-product',
      'google-sheets',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--corpus',
      'wide-mixed-variable-250k',
      '--ready-timeout-ms',
      '180000',
    ])

    expect(args).toMatchObject({
      authUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      corpusId: 'wide-mixed-variable-250k',
      headless: false,
      product: 'google-sheets',
      readyTimeoutMs: 180000,
    })
    expect(args?.targetPath.endsWith('/tmp/google-state.json')).toBe(true)
  })

  it('requires an auth URL in storage-state bootstrap mode', () => {
    expect(() => parseSaveStorageStateArgs(['--save-storage-state', 'tmp/state.json'])).toThrow('Missing auth URL.')
  })

  it('rejects storage-state bootstrap mode when the next flag would be consumed as the output path', () => {
    expect(() =>
      parseSaveStorageStateArgs([
        '--save-storage-state',
        '--auth-product',
        'google-sheets',
        '--google-sheets-url',
        'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      ]),
    ).toThrow('Missing file path after --save-storage-state')
  })

  it('blocks Playwright-backed capture modes while the local resource guard is active', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ui-same-corpus-guard-'))
    const coordinationDir = join(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(
      join(coordinationDir, '20260508T092619Z-codex-memory-pressure-stop.md'),
      '# Memory pressure stop\n\nStatus: active on 2026-05-08T09:26:19Z.\n',
    )

    expect(() => assertSameCorpusBrowserRunAllowed(rootDir, {})).toThrow(/same-corpus UI browser capture/u)
    expect(() => assertSameCorpusBrowserRunAllowed(rootDir, { BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD: '1' })).not.toThrow()
  })

  it('rejects unknown corpus ids', () => {
    expect(() =>
      parseCaptureArgs([
        '--output',
        'tmp/ui-capture.json',
        '--google-sheets-url',
        'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        '--microsoft-excel-web-url',
        'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
        '--corpus',
        'tiny-demo',
      ]),
    ).toThrow('Unexpected workbook benchmark corpus id: tiny-demo')
  })
})

function sameCorpusCaptureMeasurement(
  product: 'bilig' | 'google-sheets' | 'microsoft-excel-web',
  method: 'bilig-benchmark-state' | 'google-sheets-xlsx-export' | 'microsoft-excel-web-source-xlsx',
  workload: UiResponsivenessSameCorpusWorkload = 'open-workbook',
) {
  const sourceWorkbookSha256 = product === 'bilig' ? 'a'.repeat(64) : product === 'google-sheets' ? 'b'.repeat(64) : 'c'.repeat(64)
  const operationResponseProof = expectedSameCorpusOperationResponseProof(workload)
  return {
    product,
    source: product === 'bilig' ? 'http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k' : 'https://example.com/sheet',
    operationResponseMsSamples: [10, 11, 12],
    operationResponseProofs: [operationResponseProof, operationResponseProof, operationResponseProof],
    ...(product === 'bilig' ? { authoritativeRenderProofMsSamples: [15, 16, 17] } : {}),
    ...(uiSameCorpusWorkloadMutatesWorkbook(workload) ? { committedTargetProofMsSamples: [40, 41, 42] } : {}),
    postOperationFrameMsSamples: [8, 9, 10],
    corpusVerification: {
      verified: true,
      method,
      sheetName: 'WideGrid',
      materializedCells: 250_000,
      corpusFingerprint: wideMixedSameCorpusFingerprint,
      sourceWorkbookSha256,
      checkedCells: sameCorpusFixtureCheckedCells,
    },
    limitations: [],
  }
}

function sameCorpusCaptureCase(args: {
  readonly workload: (typeof requiredUiResponsivenessSameCorpusWorkloads)[number]
  readonly biligRuntimeProof: ReturnType<typeof sameCorpusBiligRuntimeProof> | null
  readonly includeAuthoritativeRenderProofTiming?: boolean
}) {
  const bilig = {
    ...sameCorpusCaptureMeasurement('bilig', 'bilig-benchmark-state', args.workload),
    ...(args.includeAuthoritativeRenderProofTiming === false ? { authoritativeRenderProofMsSamples: undefined } : {}),
    ...(args.biligRuntimeProof ? { biligRuntimeProof: args.biligRuntimeProof } : {}),
  }
  const googleSheets = sameCorpusCaptureMeasurement('google-sheets', 'google-sheets-xlsx-export', args.workload)
  const scenarioProof = buildCaptureScenarioProof({
    bilig,
    googleSheets,
    workload: args.workload,
    visualProofs: [
      sameCorpusVisualProof('bilig', 'typegpu-visible-canvas', `same-corpus-wide-mixed-250k-${args.workload}`, args.workload),
      sameCorpusVisualProof('google-sheets', 'google-sheets-visible-grid', `same-corpus-wide-mixed-250k-${args.workload}`, args.workload),
    ],
  })
  return {
    id: `same-corpus-wide-mixed-250k-${args.workload}`,
    corpusCaseId: 'wide-mixed-250k' as const,
    materializedCells: 250_000,
    workload: args.workload,
    ...sameCorpusScenarioCaseFields(scenarioProof),
    scenarioProof,
    bilig,
    googleSheets,
  }
}

function expectedSameCorpusOperationResponseProof(workload: UiResponsivenessSameCorpusWorkload): SameCorpusOperationResponseProof {
  if (workload === 'open-workbook') {
    return 'load-to-ready'
  }
  if (uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
    return 'visible-scroll-movement'
  }
  return 'visible-non-scroll-response'
}

function sameCorpusBiligRuntimeProof(buildKind: 'development' | 'production' | 'unknown') {
  const prod = buildKind === 'production'
  const dev = buildKind === 'development'
  return {
    product: 'bilig' as const,
    source: 'http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k',
    verificationMethod: 'window.__biligRuntimeBuild' as const,
    requiredBuildKind: 'production' as const,
    actualBuildKind: buildKind,
    mode: prod ? 'production' : buildKind,
    dev,
    prod,
    remoteSyncEnabled: false,
    entryRoute: 'workbook',
    sampleCount: 3,
    verified: prod,
    samples: [0, 1, 2].map((sampleIndex) => ({
      sampleIndex,
      present: true,
      app: 'bilig-web',
      buildKind,
      mode: prod ? 'production' : buildKind,
      dev,
      prod,
      remoteSyncEnabled: false,
      entryRoute: 'workbook',
    })),
  }
}

function sameCorpusVisualProof(
  product: SameCorpusProductVisualProof['product'],
  method: SameCorpusProductVisualProof['pixelGridProof']['method'],
  caseId = 'same-corpus-wide-mixed-250k-open-workbook',
  workload: UiResponsivenessSameCorpusWorkload = 'open-workbook',
): SameCorpusProductVisualProof {
  return {
    product,
    screenshotPath: `tmp/${caseId}/${product}-sample-1.png`,
    screenshotCaptured: true,
    pixelGridProof: {
      product,
      captured: true,
      method,
      viewportPixelWidth: 1440,
      viewportPixelHeight: 900,
      evidence:
        product === 'bilig'
          ? [
              'gridCssWidth=720',
              'gridCssHeight=450',
              'devicePixelRatio=2',
              'expectedPixelWidth=1440',
              'expectedPixelHeight=900',
              'contractVersion=same-corpus-ui-v8',
              'gridAuthoritativeRevision=rev-3',
              'gridLocalRevision=rev-local-2',
              'gridProjectedRevision=rev-3',
              'fallbackMounted=false',
              'mode=typegpu-v3',
              'backendStatus=ready',
              'frameProofStatus=presented',
              'frameProofSignature=frame-current',
              'hasPresentedFrame=true',
              'hasPresentedVisibleFrame=true',
              'presentedFrameProofSignature=frame-current',
              'currentSceneEpochSignature=epoch-current',
              'currentSceneOwnershipSignature=scene-current',
              'presentedSceneEpochSignature=epoch-current',
              'presentedSceneOwnershipSignature=scene-current',
              'currentWorkbookRevision=workbook-current',
              'presentedWorkbookRevision=workbook-current',
              'currentSemanticMutationRevision=semantic-current',
              'presentedSemanticMutationRevision=semantic-current',
              'currentViewportRevision=viewport-current',
              'presentedViewportRevision=viewport-current',
              'currentSelectionRevision=selection-current',
              'presentedSelectionRevision=selection-current',
              'currentFillHandleRevision=fill-current',
              'presentedFillHandleRevision=fill-current',
              'currentContentSignature=content-current',
              'presentedContentSignature=content-current',
              'currentTextRunCount=12',
              'presentedTextRunCount=12',
              'currentTextSignature=text-current',
              'presentedTextSignature=text-current',
              'currentRectCount=88',
              'presentedRectCount=88',
              'currentRectSignature=rect-current',
              'presentedRectSignature=rect-current',
              'tilePaneCount=6',
              'headerPaneCount=3',
              'presentedTilePaneCount=6',
              'presentedHeaderPaneCount=3',
              'canvasPixelWidth=1440',
              'canvasPixelHeight=900',
              'canvasCoversViewport=true',
              'typeGpuAuthoritativeRevision=rev-3',
              'typeGpuLocalRevision=rev-local-2',
              'typeGpuProjectedRevision=rev-3',
              'visibleAuthoritativeRevision=rev-3',
              'visibleLocalRevision=rev-local-2',
              'visibleProjectedRevision=rev-3',
              'tileSceneRevision=scene-7',
              'visibleRenderRevision=scene-7',
              ...strictPixelGridEvidence(),
            ]
          : strictPixelGridEvidence(),
    },
    semanticUiProof: {
      product,
      captured: true,
      method:
        product === 'bilig'
          ? 'bilig-visible-semantic-readback'
          : product === 'google-sheets'
            ? 'google-sheets-visible-semantic-readback'
            : 'excel-web-visible-semantic-readback',
      sheetName: 'WideGrid',
      sheetId: sameCorpusFixtureSheetId(product),
      selectedRange: sameCorpusSemanticSelectedRange(workload),
      checkedCells: sameCorpusFixtureCheckedCells,
      authoritativeRenderRevision: product === 'bilig' ? 'rev-3' : null,
      visibleRenderRevision: product === 'bilig' ? 'scene-7' : null,
      screenshotSha256: 'a'.repeat(64),
      mutationTargetProofs: sameCorpusMutationTargetProofs(product, workload),
      evidence: [
        'sheetName=WideGrid',
        `sheetId=${sameCorpusFixtureSheetId(product)}`,
        `selectedRange=${sameCorpusSemanticSelectedRange(workload)}`,
        `checkedCellCount=${String(sameCorpusFixtureCheckedCells.length)}`,
        'screenshotSha256=' + 'a'.repeat(64),
        ...(product === 'bilig' ? ['authoritativeRenderRevision=rev-3', 'visibleRenderRevision=scene-7'] : []),
      ],
    },
  }
}

function sameCorpusFixtureSheetId(product: SameCorpusProductVisualProof['product']): string {
  if (product === 'bilig') {
    return 'sheet-wide-grid'
  }
  if (product === 'google-sheets') {
    return 'gid:160971404'
  }
  return 'excel-web-sheet-wide-grid'
}

function sameCorpusSemanticSelectedRange(workload: UiResponsivenessSameCorpusWorkload): string {
  return uiSameCorpusWorkloadMutatesWorkbook(workload) ? sameCorpusMutationTargetRangeForSample(workload, 0) : 'A1'
}

function sameCorpusVisualProofWithMutationProofs(
  product: SameCorpusProductVisualProof['product'],
  method: SameCorpusProductVisualProof['pixelGridProof']['method'],
  caseId: string,
  workload: UiResponsivenessSameCorpusWorkload,
  mapProof: (proof: SameCorpusMutationTargetProof) => SameCorpusMutationTargetProof,
): SameCorpusProductVisualProof {
  const visualProof = sameCorpusVisualProof(product, method, caseId, workload)
  const semanticUiProof = visualProof.semanticUiProof
  if (!semanticUiProof) {
    throw new Error('Expected semantic UI proof fixture')
  }
  return {
    ...visualProof,
    semanticUiProof: {
      ...semanticUiProof,
      mutationTargetProofs: semanticUiProof.mutationTargetProofs.map(mapProof),
    },
  }
}

function sameCorpusMutationTargetProofs(product: SameCorpusProductVisualProof['product'], workload: UiResponsivenessSameCorpusWorkload) {
  if (!uiSameCorpusWorkloadMutatesWorkbook(workload)) {
    return []
  }
  return [0, 1, 2].map((sampleIndex) => {
    const committedTargetProofMs = 40 + sampleIndex
    const operationStartedAtMs = 1000 + sampleIndex * 100
    return {
      product,
      sampleIndex,
      committedTargetProofMs,
      operationStartedAtMs,
      postMutationProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs,
      restoreProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs + 80,
      workload,
      intendedOperation: workload,
      intendedPayload: sameCorpusMutationTargetIntendedPayload(workload, sampleIndex),
      sheetName: 'WideGrid',
      sheetId: sameCorpusFixtureSheetId(product),
      targetRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
      before: sameCorpusMutationReadback(product, workload, 'before', sampleIndex),
      after: sameCorpusMutationReadback(product, workload, 'after', sampleIndex),
      restored: sameCorpusMutationReadback(product, workload, 'before', sampleIndex),
      visibleAfter: sameCorpusVisibleMutationReadback(product, workload, 'after', sampleIndex),
      visibleAfterSelectedRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
      visibleRestored: sameCorpusVisibleMutationReadback(product, workload, 'before', sampleIndex),
      visibleRestoredSelectedRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
      authoritativeReadbackRevision: sameCorpusAuthoritativeReadbackRevision(product, sampleIndex),
      visibleRenderRevision: sameCorpusVisibleRenderRevision(product, sampleIndex),
      targetScreenshots: sameCorpusMutationTargetScreenshots(product, workload, sampleIndex),
      screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${sampleIndex + 1}-after.png`,
      screenshotSha256: sameCorpusMutationTargetScreenshotSha256(sampleIndex, 'after'),
      undoRestoreStatus: 'verified' as const,
    }
  })
}

function sameCorpusMutationTargetScreenshots(
  product: SameCorpusProductVisualProof['product'],
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['targetScreenshots'] {
  return {
    before: sameCorpusMutationTargetScreenshot(product, workload, sampleIndex, 'before'),
    after: sameCorpusMutationTargetScreenshot(product, workload, sampleIndex, 'after'),
    restored: sameCorpusMutationTargetScreenshot(product, workload, sampleIndex, 'restored'),
  }
}

function sameCorpusMutationTargetScreenshot(
  product: SameCorpusProductVisualProof['product'],
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
  phase: 'before' | 'after' | 'restored',
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before'] {
  return {
    phase,
    product,
    scope: 'target-cell',
    sampleIndex,
    sheetId: sameCorpusFixtureSheetId(product),
    sheetName: 'WideGrid',
    targetRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
    workload,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${sampleIndex + 1}-${phase}.png`,
    screenshotSha256: sameCorpusMutationTargetScreenshotSha256(sampleIndex, phase),
  }
}

function sameCorpusMutationTargetScreenshotSha256(sampleIndex: number, phase: 'before' | 'after' | 'restored'): string {
  const hexChars = '0123456789abcdef'
  const phaseOffset = phase === 'before' ? 1 : phase === 'after' ? 5 : 9
  return hexChars[(sampleIndex + phaseOffset) % hexChars.length]?.repeat(64) ?? '0'.repeat(64)
}

function sameCorpusMutationTargetIntendedPayload(
  workload: UiResponsivenessSameCorpusWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['intendedPayload'] {
  if (workload === 'formula-edit') {
    return { kind: 'formula', formula: sameCorpusFormulaEditFormula(sampleIndex) }
  }
  if (workload === 'fill-format-change') {
    const swatches = [
      { label: 'light cornflower blue 3', value: '#c9daf8' },
      { label: 'theme green', value: '#34a853' },
      { label: 'light cornflower blue 2', value: '#a4c2f4' },
    ] as const
    const swatch = swatches[sampleIndex % swatches.length]
    return { kind: 'fill-color', expectedFillColor: swatch.value, swatchLabel: swatch.label }
  }
  return { kind: 'cell-value', value: sameCorpusEditVisibleCellValue(sampleIndex) }
}

function sameCorpusExpectedFillColor(sampleIndex: number): string {
  const colors = ['#c9daf8', '#34a853', '#a4c2f4'] as const
  return colors[sampleIndex % colors.length]
}

function corruptFirstMutationTargetProof(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (proof.sampleIndex !== 0) {
    return proof
  }
  return {
    ...proof,
    after: { ...proof.before },
    authoritativeReadbackRevision: null,
    restored: { ...proof.after },
    screenshotPath: null,
    targetRange: '',
    undoRestoreStatus: 'failed',
    visibleAfter: { ...proof.visibleAfter, value: proof.before.value, visibleText: proof.before.visibleText },
    visibleRestored: { ...proof.visibleAfter },
    visibleRenderRevision: null,
  }
}

function forceVisibleEditorReadbackSource(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  return {
    ...proof,
    before: { ...proof.before, source: 'visible-formula-bar' },
    after: { ...proof.after, source: 'visible-formula-bar' },
    restored: { ...proof.restored, source: 'visible-formula-bar' },
  }
}

function driftVisibleTargetReadback(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (proof.sampleIndex !== 0) {
    return proof
  }
  return {
    ...proof,
    visibleAfter: {
      ...proof.visibleAfter,
      value: 'editor-only-ghost',
      visibleText: 'editor-only-ghost',
    },
  }
}

function driftIntendedFormulaPayload(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (proof.sampleIndex !== 0) {
    return proof
  }
  return {
    ...proof,
    intendedPayload: { kind: 'formula', formula: '=999+1' },
  }
}

function removeRenderedPostMutationFill(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (proof.sampleIndex !== 0) {
    return proof
  }
  return {
    ...proof,
    before: { ...proof.before, fillColor: '#34a853' },
    restored: { ...proof.before, fillColor: '#34a853' },
    visibleRestored: { ...proof.visibleRestored, fillColor: '#34a853' },
    visibleAfter: { ...proof.visibleAfter, fillColor: null },
  }
}

function driftIntendedFillColor(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (proof.sampleIndex !== 0) {
    return proof
  }
  return {
    ...proof,
    after: { ...proof.after, fillColor: '#34a853' },
    visibleAfter: { ...proof.visibleAfter, fillColor: '#34a853' },
  }
}

function moveMutationTargetAwayFromRenderedSelection(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  return {
    ...proof,
    targetRange: 'B2',
  }
}

function driftMutationTargetProofProduct(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  return {
    ...proof,
    product: 'google-sheets',
  }
}

function driftMutationTargetScreenshotIdentity(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (proof.sampleIndex !== 0 || !proof.targetScreenshots) {
    return proof
  }
  return {
    ...proof,
    targetScreenshots: {
      ...proof.targetScreenshots,
      before: {
        ...proof.targetScreenshots.before,
        product: 'google-sheets',
        sampleIndex: proof.sampleIndex + 1,
        sheetId: 'gid:drifted',
        sheetName: 'DriftedSheet',
        workload: 'formula-edit',
      },
    },
  }
}

function fakeGoogleSheetsNameBoxPage(calls: string[], selectedRange: string): SameCorpusNameBoxPage & SameCorpusNameBoxReaderPage {
  const locator = {
    fill: async (value: string) => {
      calls.push(`fill:${value}`)
    },
    first: () => locator,
    inputValue: async () => {
      calls.push('inputValue')
      return selectedRange
    },
    press: async (key: string) => {
      calls.push(`press:${key}`)
    },
  }
  return {
    keyboard: {
      press: async (key: string) => {
        calls.push(key === 'Escape' ? 'keyboard:Escape' : key)
      },
    },
    locator: (selector: string) => {
      calls.push(`locator:${selector}`)
      return locator
    },
  }
}

function testPrimaryShortcut(): 'Meta' | 'Control' {
  return process.platform === 'darwin' ? 'Meta' : 'Control'
}

function sameCorpusVisibleMutationReadback(
  product: SameCorpusProductVisualProof['product'],
  workload: UiResponsivenessSameCorpusWorkload,
  phase: 'before' | 'after',
  sampleIndex: number,
) {
  const source = product === 'bilig' && workload === 'fill-format-change' ? 'visible-grid-cell' : 'visible-formula-bar'
  return {
    ...sameCorpusMutationReadback(product, workload, phase, sampleIndex),
    source,
  }
}

function sameCorpusMutationReadback(
  product: SameCorpusProductVisualProof['product'],
  workload: UiResponsivenessSameCorpusWorkload,
  phase: 'before' | 'after',
  sampleIndex: number,
) {
  const after = phase === 'after'
  const source = product === 'bilig' ? ('bilig-authoritative-range' as const) : ('visible-formula-bar' as const)
  if (workload === 'formula-edit') {
    return {
      value: after ? String(sampleIndex + 2) : 'metric-1',
      formula: after ? `=${String(sampleIndex + 1)}+1` : null,
      fillColor: null,
      visibleText: after ? String(sampleIndex + 2) : 'metric-1',
      source,
      ...sameCorpusBiligRevisionReadbackFields(product, phase, sampleIndex),
    }
  }
  if (workload === 'fill-format-change') {
    return {
      value: 'metric-1',
      formula: null,
      fillColor: after ? sameCorpusExpectedFillColor(sampleIndex) : null,
      visibleText: 'metric-1',
      source,
      ...sameCorpusBiligRevisionReadbackFields(product, phase, sampleIndex),
    }
  }
  return {
    value: after ? sameCorpusEditVisibleCellValue(sampleIndex) : 'metric-1',
    formula: null,
    fillColor: null,
    visibleText: after ? sameCorpusEditVisibleCellValue(sampleIndex) : 'metric-1',
    source,
    ...sameCorpusBiligRevisionReadbackFields(product, phase, sampleIndex),
  }
}

function sameCorpusAuthoritativeReadbackRevision(product: SameCorpusProductVisualProof['product'], sampleIndex: number): string {
  return product === 'bilig' ? sameCorpusBiligCapturedRevision('after', sampleIndex) : `authoritative-readback-${sampleIndex + 1}`
}

function sameCorpusVisibleRenderRevision(product: SameCorpusProductVisualProof['product'], sampleIndex: number): string {
  return product === 'bilig'
    ? `bilig-visible-scene-sha256:${sameCorpusBiligVisibleSceneSha256(sampleIndex)}`
    : `visible-render-${sampleIndex + 1}`
}

function sameCorpusBiligRevisionReadbackFields(
  product: SameCorpusProductVisualProof['product'],
  phase: 'before' | 'after',
  sampleIndex: number,
): { readonly capturedRevision?: string; readonly visibleSceneProofSha256?: string } {
  if (product !== 'bilig') {
    return {}
  }
  return {
    capturedRevision: sameCorpusBiligCapturedRevision(phase, sampleIndex),
    ...(phase === 'after' ? { visibleSceneProofSha256: sameCorpusBiligVisibleSceneSha256(sampleIndex) } : {}),
  }
}

function sameCorpusBiligCapturedRevision(phase: 'before' | 'after', sampleIndex: number): string {
  return `${phase}-readback-${sampleIndex + 1}`
}

function sameCorpusBiligVisibleSceneSha256(sampleIndex: number): string {
  return String(sampleIndex + 1)
    .repeat(64)
    .slice(0, 64)
}

function strictPixelGridEvidence(): string[] {
  return [
    'pixelGridProofVersion=grid-pixels-v1',
    'pixelSampleSource=screenshot',
    'screenshotPixelWidth=1440',
    'screenshotPixelHeight=900',
    'nonBlankPixels=10000',
    'visibleGridLinePixels=4000',
    'verticalLineRuns=8',
    'horizontalLineRuns=16',
    'verticalLineCoverageBands=6',
    'horizontalLineCoverageBands=6',
    'largestVerticalLineGapPx=120',
    'largestHorizontalLineGapPx=24',
  ]
}

function evidenceWithoutKey(evidence: readonly string[], key: string): string[] {
  return evidence.filter((entry) => !entry.startsWith(`${key}=`))
}

function replaceEvidence(evidence: readonly string[], key: string, value: string): string[] {
  return evidence.map((entry) => (entry.startsWith(`${key}=`) ? `${key}=${value}` : entry))
}
