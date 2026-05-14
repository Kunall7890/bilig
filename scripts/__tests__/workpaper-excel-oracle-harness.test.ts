import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import {
  classifyFormulaComparison,
  runCacheDiagnostic,
  runExcelOracleEvaluation,
  writeSummary,
  type NormalizedFormulaValue,
} from '../workpaper-excel-oracle-harness.ts'

const numberValue = (value: number): NormalizedFormulaValue => ({ kind: 'number', value })

describe('WorkPaper Excel oracle harness classifier', () => {
  it('classifies a workbook where cache equals Excel and Bilig matches', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(3),
        embeddedCacheValue: numberValue(3),
        excelOracleValue: numberValue(3),
        formula: 'A1+B1',
      }),
    ).toBe('bilig_matches_excel')
  })

  it('classifies a workbook where the cache is stale but Bilig matches fresh Excel', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(3),
        embeddedCacheValue: numberValue(2),
        excelOracleValue: numberValue(3),
        formula: 'A1+B1',
      }),
    ).toBe('cache_stale_bilig_matches_excel')
  })

  it('classifies a workbook where the cache is stale and Bilig mismatches fresh Excel', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(4),
        embeddedCacheValue: numberValue(2),
        excelOracleValue: numberValue(3),
        formula: 'A1+B1',
      }),
    ).toBe('cache_stale_bilig_mismatches_excel')
  })

  it('skips volatile formulas before comparing cache or Bilig output', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(45100),
        embeddedCacheValue: numberValue(45000),
        excelOracleValue: numberValue(45100),
        formula: 'TODAY()',
      }),
    ).toBe('volatile_skipped')
  })

  it('marks a workbook without an Excel oracle as missing oracle instead of an accuracy failure', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(3),
        embeddedCacheValue: numberValue(2),
        formula: 'A1+B1',
      }),
    ).toBe('missing_excel_oracle')
  })

  it('rejects an Excel oracle value when Excel rewrites the formula as an unsupported UDF', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: { kind: 'string', value: '2026-04-01' },
        embeddedCacheValue: { kind: 'string', value: '2026-04-01' },
        excelOracleFormula: 'IFERROR(_xludf.XLOOKUP(C14,Bank!$D$2:$D$31,Bank!$B$2:$B$31,"",0),"")',
        excelOracleValue: { kind: 'string', value: '' },
        formula: 'IFERROR(XLOOKUP(C14,Bank!$D$2:$D$31,Bank!$B$2:$B$31,"",0),"")',
      }),
    ).toBe('missing_excel_oracle')
  })

  it('accepts Excel compatibility prefixes when comparing oracle formulas', () => {
    expect(
      classifyFormulaComparison({
        actualBiligValue: numberValue(20),
        embeddedCacheValue: numberValue(20),
        excelOracleFormula: '_xlfn.XLOOKUP(2,A2:A4,B2:B4)',
        excelOracleValue: numberValue(20),
        formula: 'XLOOKUP(2,A2:A4,B2:B4)',
      }),
    ).toBe('bilig_matches_excel')
  })

  it('classifies synthetic workbooks without promoting cache-only mismatches to accuracy bugs', () => {
    withTempDirs((originalDir, recalculatedDir, outputDir) => {
      writeWorkbookPair(
        originalDir,
        recalculatedDir,
        'fresh.xlsx',
        formulaWorkbook({ cachedValue: 2, marker: 'fresh' }),
        formulaWorkbook({ cachedValue: 2, marker: 'fresh' }),
      )
      writeWorkbookPair(
        originalDir,
        recalculatedDir,
        'stale-bilig-ok.xlsx',
        formulaWorkbook({ cachedValue: 99, marker: 'stale-bilig-ok' }),
        formulaWorkbook({ cachedValue: 2, marker: 'stale-bilig-ok' }),
      )
      writeWorkbookPair(
        originalDir,
        recalculatedDir,
        'stale-bilig-bad.xlsx',
        formulaWorkbook({ cachedValue: 99, marker: 'stale-bilig-bad' }),
        formulaWorkbook({ cachedValue: 3, inputValue: 2, marker: 'stale-bilig-bad' }),
      )
      writeWorkbookPair(originalDir, recalculatedDir, 'volatile.xlsx', volatileWorkbook(46_127), volatileWorkbook(46_128))
      writeOriginalWorkbook(originalDir, 'missing-oracle.xlsx', formulaWorkbook({ cachedValue: 2, marker: 'missing-oracle' }))

      const cacheDiagnostic = runCacheDiagnostic(originalDir, outputDir, { sampleLimit: 25, timeoutMs: 30_000 })
      const oracle = runExcelOracleEvaluation(originalDir, recalculatedDir, outputDir, { sampleLimit: 25, timeoutMs: 30_000 })
      const summary = writeSummary(outputDir)

      expect(cacheDiagnostic.mode).toBe('cache-diagnostic')
      expect(cacheDiagnostic.summary.biligVsFreshExcelMatchRate).toBeNull()
      expect(cacheDiagnostic.summary.cacheOnlyMismatches).toBe(2)
      expect(cacheDiagnostic.workbooks.flatMap((workbook) => workbook.comparisons).map((entry) => entry.classification)).toEqual(
        expect.arrayContaining(['missing_excel_oracle', 'volatile_skipped']),
      )

      expect(oracle.summary).toMatchObject({
        totalWorkbooksEvaluated: 5,
        importParserFailures: 0,
        timeoutFailures: 0,
        missingExcelOracleWorkbooks: 1,
        totalFormulaCells: 5,
        comparableFormulaCells: 3,
        biligVsFreshExcelMatchRate: 0.666667,
        embeddedCacheFreshnessRate: 0.333333,
        staleCacheFalsePositives: 1,
        realBiligMismatches: 1,
        cacheOnlyMismatches: 0,
      })
      expect(oracle.workbooks.flatMap((workbook) => workbook.comparisons).map((entry) => entry.classification)).toEqual(
        expect.arrayContaining([
          'bilig_matches_excel',
          'cache_stale_bilig_matches_excel',
          'cache_stale_bilig_mismatches_excel',
          'volatile_skipped',
          'missing_excel_oracle',
        ]),
      )
      const trueMismatch = oracle.workbooks
        .flatMap((workbook) => workbook.comparisons)
        .find((entry) => entry.classification.includes('mismatches_excel'))
      expect(trueMismatch).toMatchObject({
        classification: 'cache_stale_bilig_mismatches_excel',
        expectedExcelValue: { kind: 'number', value: 3 },
        actualBiligValue: { kind: 'number', value: 2 },
        formula: 'A1+1',
      })
      expect(existsSync(join(outputDir, 'cache-diagnostic.json'))).toBe(true)
      expect(existsSync(join(outputDir, 'excel-oracle-report.json'))).toBe(true)
      expect(existsSync(join(outputDir, 'summary.md'))).toBe(true)
      expect(existsSync(join(outputDir, 'github-issues', 'formula-mismatch-01.md'))).toBe(true)
      expect(readFileSync(join(outputDir, 'summary.md'), 'utf8')).toContain('Embedded XLSX cached values are diagnostic only')
      expect(summary.report?.summary.realBiligMismatches).toBe(1)
    })
  })
})

function withTempDirs(run: (originalDir: string, recalculatedDir: string, outputDir: string) => void): void {
  const root = mkdtempSync(join(tmpdir(), 'bilig-workpaper-excel-oracle-harness-'))
  try {
    run(join(root, 'original'), join(root, 'recalculated'), join(root, 'output'))
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
}

function writeWorkbookPair(
  originalDir: string,
  recalculatedDir: string,
  fileName: string,
  originalWorkbook: XLSX.WorkBook,
  recalculatedWorkbook: XLSX.WorkBook,
): void {
  const originalBytes = writeOriginalWorkbook(originalDir, fileName, originalWorkbook)
  const workbookId = `workbook-${createHash('sha256').update(originalBytes).digest('hex').slice(0, 16)}`
  mkdirSync(recalculatedDir, { recursive: true })
  writeFileSync(join(recalculatedDir, `${workbookId}.xlsx`), XLSX.write(recalculatedWorkbook, { bookType: 'xlsx', type: 'buffer' }))
}

function writeOriginalWorkbook(originalDir: string, fileName: string, workbook: XLSX.WorkBook): Buffer {
  mkdirSync(originalDir, { recursive: true })
  const bytes: unknown = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
  if (!Buffer.isBuffer(bytes)) {
    throw new Error('Expected XLSX.write to return a Buffer for xlsx buffer output')
  }
  writeFileSync(join(originalDir, fileName), bytes)
  return bytes
}

function formulaWorkbook(args: { readonly cachedValue: number; readonly inputValue?: number; readonly marker: string }): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[args.inputValue ?? 1, null, args.marker]])
  sheet.B1 = { t: 'n', f: 'A1+1', v: args.cachedValue }
  sheet['!ref'] = 'A1:C1'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
  return workbook
}

function volatileWorkbook(cachedValue: number): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[null]])
  sheet.A1 = { t: 'n', f: 'NOW()', v: cachedValue }
  sheet['!ref'] = 'A1:A1'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
  return workbook
}
