import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import {
  createMacosExcelRecalculationAppleScript,
  isMacosExcelInstalled,
  parseMacosExcelRecalculationOutput,
  runMacosExcelRecalculationOracle,
} from '../macos-excel-oracle.js'

describe('macOS Desktop Excel oracle harness', () => {
  it('builds an AppleScript runner that opens, recalculates, reads, and closes a workbook', () => {
    const script = createMacosExcelRecalculationAppleScript({
      worksheetName: 'Cases',
      formulaCells: [{ address: 'C1', formula: '=A1+B1' }],
      valueCells: ['C1'],
    })

    expect(script).toContain('tell application "Microsoft Excel"')
    expect(script).toContain('open workbook workbook file name workbookPath')
    expect(script).toContain('set formula of range "C1"')
    expect(script).toContain('calculate full rebuild')
    expect(script).toContain('close targetWorkbook saving no')
    expect(script).not.toContain('active workbook')
  })

  it('can save the opened workbook when a caller needs fresh Excel caches persisted', () => {
    const script = createMacosExcelRecalculationAppleScript({
      worksheetName: 'Cases',
      formulaCells: [],
      valueCells: ['C1'],
      saveWorkbook: true,
    })

    expect(script).toContain('close targetWorkbook saving yes')
  })

  it('parses typed Excel oracle values into normalized formula values', () => {
    expect(
      parseMacosExcelRecalculationOutput(['version=16.96', 'number\t42', 'boolean\ttrue', 'string\tBilig', 'blank\t'].join('\n'), 4),
    ).toEqual({
      excelVersion: '16.96',
      rawValues: ['number\t42', 'boolean\ttrue', 'string\tBilig', 'blank\t'],
      values: [{ kind: 'number', value: 42 }, { kind: 'boolean', value: true }, { kind: 'string', value: 'Bilig' }, { kind: 'blank' }],
    })
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'uses real Microsoft Excel for Mac as the recalculation oracle',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-excel-fixtures-live-'))
      try {
        const workbookPath = join(tempDir, 'oracle.xlsx')
        const worksheet = XLSX.utils.aoa_to_sheet([[10, 3, null]])
        const workbook = XLSX.utils.book_new()
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Cases')
        writeFileSync(workbookPath, XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }))

        const result = runMacosExcelRecalculationOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [{ address: 'C1', formula: '=A1+B1*2' }],
          valueCells: ['C1'],
        })

        expect(result.excelVersion).toMatch(/^\d+\./u)
        expect(result.values).toEqual([{ kind: 'number', value: 16 }])
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )
})
