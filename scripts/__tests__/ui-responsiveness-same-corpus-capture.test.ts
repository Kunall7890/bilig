import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { exportXlsx } from '../../packages/excel-import/src/index.js'
import { buildWorkbookBenchmarkCorpus } from '../../packages/benchmarks/src/workbook-corpus.js'
import {
  assertSameCorpusBrowserRunAllowed,
  buildSameCorpusFingerprint,
  collectSameCorpusProductMeasurements,
  parseCaptureArgs,
  parseEmitXlsxArgs,
  parsePreflightArgs,
  parseSaveStorageStateArgs,
  verifyXlsxCorpusFingerprint,
} from '../capture-ui-responsiveness-same-corpus.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from '../ui-responsiveness-same-corpus-workloads.ts'
import {
  buildCaptureScenarioProof,
  isSameCorpusProductPixelGridProofComplete,
  type SameCorpusProductVisualProof,
} from '../ui-responsiveness-same-corpus-proof.ts'
import { sameCorpusChromiumLaunchOptions } from '../ui-responsiveness-same-corpus-page-utils.ts'
import { sameCorpusScrollProbeSelectorsForProduct } from '../ui-responsiveness-same-corpus-scroll-page.ts'
import { incumbentEditableWorkloadBlocker, sameCorpusKeyboardOperations } from '../ui-responsiveness-same-corpus-workload-runner.ts'

const sameCorpusFixtureCheckedCells = [
  { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
  { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
  { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
] as const

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
      biligUrl: 'http://127.0.0.1:5173/?benchmarkCorpus=dense-mixed-250k',
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
      biligUrl: 'http://127.0.0.1:4173/?benchmarkCorpus=wide-mixed-250k',
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

  it('rejects missing incumbent URLs because the generated proof must be comparable', () => {
    expect(() => parseCaptureArgs(['--output', 'tmp/ui-capture.json'])).toThrow('Missing required arguments.')
  })

  it('rejects blank capture argument values', () => {
    expect(() =>
      parseCaptureArgs(['--output', '   ', '--google-sheets-url', 'https://docs.google.com/spreadsheets/d/sheet-id/edit']),
    ).toThrow('Missing value after --output')
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
    expect(sameCorpusKeyboardOperations('bilig', 'select-cell', 0, 'darwin')).toEqual([{ kind: 'press', key: 'ArrowRight' }])
    expect(sameCorpusKeyboardOperations('bilig', 'jump-deep-row', 0, 'darwin')).toEqual([{ kind: 'press', key: 'Meta+ArrowDown' }])
    expect(sameCorpusKeyboardOperations('bilig', 'fill-format-change', 0, 'linux')).toEqual([{ kind: 'press', key: 'Control+B' }])
    expect(sameCorpusKeyboardOperations('bilig', 'formula-edit', 1, 'darwin')).toEqual([
      { kind: 'type', text: '=2+1' },
      { kind: 'press', key: 'Enter' },
    ])
    expect(sameCorpusKeyboardOperations('google-sheets', 'edit-visible-cell', 2, 'darwin')).toEqual([
      { kind: 'type', text: 'google-sheets-same-corpus-3' },
      { kind: 'press', key: 'Enter' },
    ])
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
        'mode',
        'contractVersion',
        'backendStatus',
        'frameProofStatus',
        'frameProofSignature',
        'hasPresentedFrame',
        'hasPresentedVisibleFrame',
        'presentedFrameProofSignature',
        'currentSceneOwnershipSignature',
        'presentedSceneOwnershipSignature',
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
      { label: 'wrong renderer mode', evidence: replaceEvidence(baseEvidence, 'mode', 'canvas2d') },
      { label: 'wrong render proof contract', evidence: replaceEvidence(baseEvidence, 'contractVersion', 'same-corpus-ui-v1') },
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
) {
  return {
    product,
    source: product === 'bilig' ? 'http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k' : 'https://example.com/sheet',
    operationResponseMsSamples: [10, 11, 12],
    postOperationFrameMsSamples: [8, 9, 10],
    corpusVerification: {
      verified: true,
      method,
      sheetName: 'WideGrid',
      materializedCells: 250_000,
      checkedCells: sameCorpusFixtureCheckedCells,
    },
    limitations: [],
  }
}

function sameCorpusVisualProof(
  product: SameCorpusProductVisualProof['product'],
  method: SameCorpusProductVisualProof['pixelGridProof']['method'],
): SameCorpusProductVisualProof {
  return {
    product,
    screenshotPath: `tmp/${product}-sample-1.png`,
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
              'contractVersion=same-corpus-ui-v2',
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
              'currentSceneOwnershipSignature=scene-current',
              'presentedSceneOwnershipSignature=scene-current',
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
  }
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
  ]
}

function evidenceWithoutKey(evidence: readonly string[], key: string): string[] {
  return evidence.filter((entry) => !entry.startsWith(`${key}=`))
}

function replaceEvidence(evidence: readonly string[], key: string, value: string): string[] {
  return evidence.map((entry) => (entry.startsWith(`${key}=`) ? `${key}=${value}` : entry))
}
