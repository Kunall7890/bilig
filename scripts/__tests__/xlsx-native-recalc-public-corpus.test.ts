import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { ValueTag } from '../../packages/protocol/src/enums.js'
import { writeSimpleXlsxWorkbook } from '../../packages/xlsx/src/simple-workbook-writer.js'
import { describe, expect, it } from 'vitest'

import {
  buildNativeRecalcPublicCorpusCliArgs,
  buildNativeRecalcPublicCorpusTarget,
  cellValueFromCliRead,
  cellValueFromFormulaCacheLiteral,
  discoverNativeRecalcPublicCorpusTargetsForCorpora,
  formulaOracleReadTarget,
  formulaOraclesFromNativeFormulaCacheCells,
  quoteSheetNameForTarget,
  readNativeRecalcPublicCorpusInputs,
  summarizeNativeRecalcPublicCorpusResults,
  type NativeRecalcPublicCorpusResult,
} from '../xlsx-native-recalc-public-corpus.ts'
import { asRecord } from '../public-workbook-corpus-json.ts'
import type { PublicWorkbookArtifact } from '../public-workbook-corpus-types.ts'

const artifact: PublicWorkbookArtifact = {
  id: 'workbook-public-formula',
  sourceId: 'source-public-formula',
  sourceUrl: 'https://example.com/public-formula-workbook',
  downloadUrl: 'https://example.com/public-formula-workbook.xlsx',
  fileName: 'public-formula-workbook.xlsx',
  cachePath: 'files/public-formula-workbook.xlsx',
  sha256: 'a'.repeat(64),
  byteSize: 1024,
  workbookFingerprint: 'fingerprint',
  fetchedAt: new Date(0).toISOString(),
  license: {
    spdxId: 'CC-BY-4.0',
    title: 'Creative Commons Attribution 4.0 International',
    evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
  },
}

describe('xlsx native recalc public corpus runner', () => {
  it('quotes formula oracle read targets for workbook sheet names', () => {
    expect(quoteSheetNameForTarget('Data')).toBe('Data')
    expect(quoteSheetNameForTarget('Real gross state product')).toBe("'Real gross state product'")
    expect(quoteSheetNameForTarget("Bob's Sheet")).toBe("'Bob''s Sheet'")

    expect(
      formulaOracleReadTarget({
        sheetName: 'Real gross state product',
        address: 'C7',
        expected: { tag: ValueTag.Number, value: 123 },
      }),
    ).toBe("'Real gross state product'!C7")
  })

  it('builds streaming-native CLI args with WorkPaper fallback disabled', () => {
    const target = buildNativeRecalcPublicCorpusTarget({
      artifact,
      inputPath: '/cache/public-formula-workbook.xlsx',
      outputDir: '/outputs',
      maxRssBytes: 350 * 1024 * 1024,
      maxFormulaReadsPerWorkbook: 1,
      oracles: [
        {
          sheetName: 'Summary',
          address: 'B9',
          expected: { tag: ValueTag.Number, value: 42 },
        },
        {
          sheetName: 'Summary',
          address: 'B10',
          expected: { tag: ValueTag.Number, value: 43 },
        },
      ],
    })

    expect(target.formulaCellCount).toBe(2)
    expect(target.selectedFormulaCellCount).toBe(1)
    expect(target.reads).toEqual(['Summary!B9'])
    expect(buildNativeRecalcPublicCorpusCliArgs(target)).toEqual([
      '/cache/public-formula-workbook.xlsx',
      '--out',
      '/outputs/workbook-public-formula-aaaaaaaaaaaa.native.xlsx',
      '--engine',
      'streaming-native',
      '--fallback-policy',
      'error',
      '--max-rss-bytes',
      String(350 * 1024 * 1024),
      '--json',
      '--read',
      'Summary!B9',
    ])
  })

  it('converts CLI read cells back to protocol values', () => {
    expect(cellValueFromCliRead({ tag: 0 })).toEqual({ tag: ValueTag.Empty })
    expect(cellValueFromCliRead({ tag: 1, value: 12.5 })).toEqual({ tag: ValueTag.Number, value: 12.5 })
    expect(cellValueFromCliRead({ tag: 2, value: true })).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(cellValueFromCliRead({ tag: 3, value: 'done' })).toEqual({ tag: ValueTag.String, value: 'done', stringId: 0 })
    expect(cellValueFromCliRead({ tag: 4, code: 5 })).toEqual({ tag: ValueTag.Error, code: 5 })
    expect(cellValueFromCliRead({ tag: 1, value: Number.NaN })).toBeNull()
  })

  it('converts native formula cache literals to protocol oracle values', () => {
    expect(cellValueFromFormulaCacheLiteral(null)).toEqual({ tag: ValueTag.Empty })
    expect(cellValueFromFormulaCacheLiteral(12.5)).toEqual({ tag: ValueTag.Number, value: 12.5 })
    expect(cellValueFromFormulaCacheLiteral(true)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(cellValueFromFormulaCacheLiteral('done')).toEqual({ tag: ValueTag.String, value: 'done', stringId: 0 })
    expect(cellValueFromFormulaCacheLiteral('#REF!')).toBeNull()
    expect(cellValueFromFormulaCacheLiteral(Number.NaN)).toBeNull()
  })

  it('builds formula oracles from @bilig/xlsx native formula cache cells', () => {
    expect(
      formulaOraclesFromNativeFormulaCacheCells([
        {
          target: "'Revenue & Ops'!B1",
          formula: '=A1*10',
          cachedValue: 20,
        },
        {
          target: "'Bob''s Sheet'!C7",
          formula: '=A7&" units"',
          cachedValue: 'old units',
        },
        {
          target: 'Data!D8',
          formula: '=1/0',
          cachedValue: '#DIV/0!',
        },
        {
          target: 'Data!E9',
          formula: '=NOW()',
        },
        {
          target: "'Broken!F10",
          formula: '=1',
          cachedValue: 1,
        },
      ]),
    ).toEqual([
      {
        sheetName: 'Revenue & Ops',
        address: 'B1',
        expected: { tag: ValueTag.Number, value: 20 },
      },
      {
        sheetName: "Bob's Sheet",
        address: 'C7',
        expected: { tag: ValueTag.String, value: 'old units', stringId: 0 },
      },
    ])
  })

  it('parses repeated native corpus manifest/cache inputs', () => {
    expect(
      readNativeRecalcPublicCorpusInputs(['--corpus', './one/manifest.json', './one', '--corpus', './two/manifest.json', './two']),
    ).toEqual([
      {
        manifestPath: join(process.cwd(), 'one/manifest.json'),
        cacheDir: join(process.cwd(), 'one'),
      },
      {
        manifestPath: join(process.cwd(), 'two/manifest.json'),
        cacheDir: join(process.cwd(), 'two'),
      },
    ])

    expect(() => readNativeRecalcPublicCorpusInputs(['--corpus', './manifest.json'])).toThrow(
      'Expected --corpus <manifest-path> <cache-dir>',
    )
  })

  it('dedupes formula workbook targets by SHA across combined public corpora', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-public-corpus-combined-'))
    try {
      const firstCacheDir = join(tempDir, 'first')
      const secondCacheDir = join(tempDir, 'second')
      mkdirSync(join(firstCacheDir, 'files'), { recursive: true })
      mkdirSync(join(secondCacheDir, 'files'), { recursive: true })
      const firstWorkbook = writeSimpleXlsxWorkbook({
        sheets: [
          {
            name: 'Summary',
            cells: [
              { address: 'A1', row: 0, col: 0, value: 2 },
              { address: 'B1', row: 0, col: 1, formula: 'A1*2', value: 4 },
            ],
          },
        ],
      })
      const secondWorkbook = writeSimpleXlsxWorkbook({
        sheets: [
          {
            name: 'Summary',
            cells: [
              { address: 'A1', row: 0, col: 0, value: 3 },
              { address: 'B1', row: 0, col: 1, formula: 'A1*2', value: 6 },
            ],
          },
        ],
      })
      writeFileSync(join(firstCacheDir, artifact.cachePath), firstWorkbook)
      writeFileSync(join(secondCacheDir, artifact.cachePath), firstWorkbook)
      writeFileSync(join(secondCacheDir, 'files/second-formula-workbook.xlsx'), secondWorkbook)

      const duplicateArtifact = { ...artifact, id: 'workbook-public-formula-duplicate' }
      const secondArtifact = {
        ...artifact,
        id: 'workbook-public-formula-second',
        sourceId: 'source-public-formula-second',
        downloadUrl: 'https://example.com/second-formula-workbook.xlsx',
        fileName: 'second-formula-workbook.xlsx',
        cachePath: 'files/second-formula-workbook.xlsx',
        sha256: 'b'.repeat(64),
      }
      const discovered = discoverNativeRecalcPublicCorpusTargetsForCorpora({
        corpora: [
          {
            cacheDir: firstCacheDir,
            manifest: {
              schemaVersion: 1,
              targetWorkbookCount: 2,
              generatedAt: '2026-06-07T00:00:00.000Z',
              sources: [],
              artifacts: [artifact],
            },
          },
          {
            cacheDir: secondCacheDir,
            manifest: {
              schemaVersion: 1,
              targetWorkbookCount: 2,
              generatedAt: '2026-06-07T00:00:00.000Z',
              sources: [],
              artifacts: [duplicateArtifact, secondArtifact],
            },
          },
        ],
        outputDir: join(tempDir, 'outputs'),
        maxRssBytes: 350 * 1024 * 1024,
        limit: 1,
        maxFormulaReadsPerWorkbook: 10,
        minFormulaCellsPerWorkbook: 1,
      })

      expect(discovered.discoveredFormulaWorkbookCount).toBe(2)
      expect(discovered.targets).toHaveLength(1)
      expect(discovered.targets[0]?.sha256).toBe(artifact.sha256)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('summarizes public native recalc results without treating unsupported formulas as passes', () => {
    const base = {
      sourceUrl: artifact.sourceUrl,
      sha256: artifact.sha256,
      formulaCellCount: 2,
      selectedFormulaCellCount: 2,
      maxRssBytes: 350 * 1024 * 1024,
      peakRssBytes: 120 * 1024 * 1024,
    }
    const results: NativeRecalcPublicCorpusResult[] = [
      {
        ...base,
        id: 'passed',
        label: 'passed.xlsx',
        status: 'passed',
        nativeKernelFormulaCellCount: 2,
        patchedCacheCount: 2,
      },
      {
        ...base,
        id: 'unsupported',
        label: 'unsupported.xlsx',
        status: 'unsupported',
        unsupportedReason: 'unsupported function: XLOOKUP',
      },
    ]

    expect(
      summarizeNativeRecalcPublicCorpusResults({
        discoveredFormulaWorkbookCount: 2,
        attemptedFormulaWorkbookCount: 2,
        requireFormulaWorkbookCount: 50,
        requirePassedFormulaWorkbookCount: 50,
        results,
      }),
    ).toEqual({
      discoveredFormulaWorkbookCount: 2,
      attemptedFormulaWorkbookCount: 2,
      requiredFormulaWorkbookCount: 50,
      requiredPassedFormulaWorkbookCount: 50,
      formulaWorkbookCoverageGap: 48,
      passedFormulaWorkbookCoverageGap: 49,
      passedWorkbookCount: 1,
      unsupportedWorkbookCount: 1,
      failedWorkbookCount: 0,
      selectedFormulaCellCount: 4,
      nativeKernelFormulaCellCount: 2,
      patchedCacheCount: 2,
      maxPeakRssBytes: 120 * 1024 * 1024,
    })
  })

  it('exposes the public corpus runner as a package script', () => {
    const packageJson = asRecord(JSON.parse(readFileSync(join(process.cwd(), 'package.json'), 'utf8')))
    const scripts = asRecord(packageJson['scripts'])
    const script = String(scripts['xlsx-native-recalc:public-corpus'])

    expect(script).toContain('bun scripts/xlsx-native-recalc-public-corpus.ts')
    expect(script).toContain('--limit 50')
    expect(script).toContain('--max-rss-mb 350')
    expect(script).toContain('--require-formula-workbook-count 50')
    expect(script).toContain('--require-passed-formula-workbook-count 50')
    expect(script).toContain('--require-passed')
    expect(script).toContain('--corpus .cache/public-workbook-corpus/manifest.json .cache/public-workbook-corpus')
    expect(script).toContain('--corpus .cache/public-workbook-corpus-financial/manifest.json .cache/public-workbook-corpus-financial')
  })
})
