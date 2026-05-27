import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { readRuntimeImage } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { WorkPaper } from '../index.js'

function sourceWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    [1, null],
    [2, null],
  ])
  sheet['B1'] = { t: 'n', f: 'A1+1', v: 2 }
  sheet['B2'] = { t: 'n', f: 'A2+1', v: 3 }
  sheet['!ref'] = 'A1:B2'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Data')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['customXml/item1.xml'] = strToU8('<keep source="true"/>')
  return zipSync(zip)
}

function attachSourceBytesForTest(snapshot: object, bytes: Uint8Array): void {
  Object.defineProperty(snapshot, Symbol.for('bilig.importedXlsxSourceBytes'), {
    configurable: true,
    enumerable: false,
    value: bytes,
  })
}

describe('WorkPaper source-preserving XLSX export', () => {
  it('exports imported XLSX scalar edits by patching the source package without a runtime image snapshot', () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving.xlsx')
    attachSourceBytesForTest(imported.snapshot, sourceBytes)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      workbook.setCellContents({ sheet: 1, row: 0, col: 0 }, 5)

      const exportedSnapshot = workbook.exportSnapshot()
      expect(readRuntimeImage(exportedSnapshot)).toBeUndefined()

      const exportedZip = unzipSync(exportXlsx(exportedSnapshot))
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
      expect(exportedZip['xl/calcChain.xml']).toBeUndefined()

      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('r="A1"')
      expect(sheetXml).toContain('<v>5</v>')
      expect(sheetXml).toContain('<f>A1+1</f><v>2</v>')

      const workbookXml = strFromU8(exportedZip['xl/workbook.xml'] ?? new Uint8Array())
      expect(workbookXml).toContain('fullCalcOnLoad="1"')
      expect(workbookXml).toContain('forceFullCalc="1"')
    } finally {
      workbook.dispose()
    }
  })
})
