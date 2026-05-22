import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { buildSharedFormulaWorkbookBytes, isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }
const inspectedCells = ['A2', 'B2', 'C2', 'A3', 'B3', 'C3', 'A4', 'B4', 'C4', 'A5', 'B5', 'C5'] as const

describe('macOS Desktop Excel shared-formula structural oracle', () => {
  it('exports canonical formulas after structural row inserts inside imported shared-formula runs', () => {
    const workbook = WorkPaper.buildFromSnapshot(
      importXlsx(buildSharedFormulaWorkbookBytes(), 'shared-formula-model.xlsx').snapshot,
      workbookConfig,
    )
    try {
      const sheetId = workbook.getSheetId('Model')
      if (sheetId === undefined) {
        throw new Error('Model sheet is missing')
      }
      workbook.addRows(sheetId, 2, 1)

      expect(workbook.getCellFormula(addressToCell('B2'))).toBe('=A2*2')
      expect(workbook.getCellFormula(addressToCell('B3'))).toBeUndefined()
      expect(workbook.getCellFormula(addressToCell('B4'))).toBe('=A4*2')
      expect(workbook.getCellFormula(addressToCell('C5'))).toBe('=B5+1')

      const exported = exportXlsx(workbook.exportSnapshot())
      const exportedSheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(exportedSheetXml).not.toContain('t="shared"')
      expect(cellXml(exportedSheetXml, 'B4')).toContain('<f>A4*2</f>')
      expect(cellXml(exportedSheetXml, 'C5')).toContain('<f>B5+1</f>')

      const roundTrip = WorkPaper.buildFromSnapshot(
        importXlsx(exported, 'shared-formula-row-insert-roundtrip.xlsx').snapshot,
        workbookConfig,
      )
      try {
        expect(roundTrip.getCellFormula(addressToCell('B3'))).toBeUndefined()
        expect(roundTrip.getCellFormula(addressToCell('B4'))).toBe('=A4*2')
        expect(roundTrip.getCellFormula(addressToCell('C5'))).toBe('=B5+1')
      } finally {
        roundTrip.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('does not reuse untouched imported XLSX bytes after explicit recalculation', () => {
    const imported = importXlsx(buildSharedFormulaWorkbookBytes({ calculationMode: 'manual' }), 'manual-shared-formula-model.xlsx')
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      const untouchedXml = strFromU8(unzipSync(exportXlsx(workbook.exportSnapshot()))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(untouchedXml).toContain('t="shared"')

      workbook.rebuildAndRecalculate()

      const recalculatedXml = strFromU8(unzipSync(exportXlsx(workbook.exportSnapshot()))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(recalculatedXml).not.toContain('t="shared"')
      expect(cellXml(recalculatedXml, 'B3')).toContain('<f>A3*2</f>')
      expect(cellXml(recalculatedXml, 'C4')).toContain('<f>B4+1</f>')
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel after inserting a row inside a native shared-formula range',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-shared-formula-structure-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-shared-formula-row-insert-oracle.xlsx')
        writeFileSync(excelWorkbookPath, buildSharedFormulaWorkbookBytes())
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Model',
          operations: [{ kind: 'insertRows', range: '3:3' }],
          inspectCells: inspectedCells,
          saveWorkbook: true,
        })

        const workbook = WorkPaper.buildFromSnapshot(
          importXlsx(buildSharedFormulaWorkbookBytes(), 'headless-shared-formula-source.xlsx').snapshot,
          workbookConfig,
        )
        try {
          const sheetId = workbook.getSheetId('Model')
          if (sheetId === undefined) {
            throw new Error('Model sheet is missing')
          }
          workbook.addRows(sheetId, 2, 1)

          const headlessWorkbookPath = join(tempDir, 'headless-shared-formula-row-insert-oracle.xlsx')
          writeFileSync(headlessWorkbookPath, exportXlsx(workbook.exportSnapshot()))
          const headlessExcel = runMacosExcelStructuralOperationOracle({
            workbookPath: headlessWorkbookPath,
            worksheetName: 'Model',
            operations: [{ kind: 'setCellValue', address: 'A3', value: '' }],
            inspectCells: inspectedCells,
            saveWorkbook: true,
          })

          expect(headlessExcel.cells.map(({ address, formula, value }) => ({ address, formula, value }))).toEqual(
            excelResult.cells.map(({ address, formula, value }) => ({ address, formula, value })),
          )
        } finally {
          workbook.dispose()
        }

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-shared-formula-row-insert-oracle.xlsx')
        expect(excelTruth.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'B4')).toMatchObject({ formula: 'A4*2' })
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    60_000,
  )
})

function addressToCell(address: string): WorkPaperCellAddress {
  const [, columnLetters, rowDigits] = /^([A-Z]+)(\d+)$/u.exec(address) ?? []
  if (!columnLetters || !rowDigits) {
    throw new Error(`Unexpected oracle address: ${address}`)
  }
  let col = 0
  for (const letter of columnLetters) {
    col = col * 26 + letter.charCodeAt(0) - 64
  }
  return { sheet: 1, row: Number(rowDigits) - 1, col: col - 1 }
}

function cellXml(sheetXml: string, address: string): string {
  return sheetXml.match(new RegExp(`<c[^>]* r="${address}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0] ?? ''
}
