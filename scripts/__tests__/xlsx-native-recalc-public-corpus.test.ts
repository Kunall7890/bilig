import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { ValueTag } from '../../packages/protocol/src/enums.js'
import { describe, expect, it } from 'vitest'

import {
  buildNativeRecalcPublicCorpusCliArgs,
  buildNativeRecalcPublicCorpusTarget,
  cellValueFromCliRead,
  formulaOracleReadTarget,
  quoteSheetNameForTarget,
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
        results,
      }),
    ).toEqual({
      discoveredFormulaWorkbookCount: 2,
      attemptedFormulaWorkbookCount: 2,
      requiredFormulaWorkbookCount: 50,
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

    expect(scripts['xlsx-native-recalc:public-corpus']).toBe('bun scripts/xlsx-native-recalc-public-corpus.ts')
  })
})
