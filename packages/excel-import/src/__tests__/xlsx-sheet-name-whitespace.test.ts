import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'
import { readXlsxTestZipText } from './xlsx-test-helpers.js'

describe('XLSX sheet-name whitespace export', () => {
  it('preserves trailing spaces in raw workbook sheet names and formulas', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'country-erp-whitespace' },
      sheets: [
        {
          id: 'country-erp',
          name: 'Country ERP ',
          order: 0,
          cells: [
            { address: 'A5', value: 'United States' },
            { address: 'F196', value: 4.25 },
          ],
          metadata: {
            merges: [{ sheetName: 'Country ERP ', startAddress: 'A226', endAddress: 'C226' }],
          },
        },
        {
          id: 'inputs',
          name: 'Inputs',
          order: 1,
          cells: [{ address: 'B15', formula: "VLOOKUP(B8,'Country ERP '!A5:F196,6,FALSE)" }],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const workbookXml = readXlsxTestZipText(exported, 'xl/workbook.xml')
    const roundTripped = importXlsx(exported, 'country-erp-whitespace.xlsx')

    expect(workbookXml).toContain('<sheet name="Country ERP "')
    expect(workbookXml).not.toContain('<sheet name="Country ERP" ')
    expect(roundTripped.snapshot.sheets.map((sheet) => sheet.name)).toContain('Country ERP ')
    expect(roundTripped.snapshot.sheets.map((sheet) => sheet.name)).not.toContain('Country ERP')

    const inputSheet = roundTripped.snapshot.sheets.find((sheet) => sheet.name === 'Inputs')
    expect(inputSheet?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          address: 'B15',
          formula: "VLOOKUP(B8,'Country ERP '!A5:F196,6,FALSE)",
        }),
      ]),
    )

    const countrySheet = roundTripped.snapshot.sheets.find((sheet) => sheet.name === 'Country ERP ')
    expect(countrySheet?.metadata?.merges).toEqual([{ sheetName: 'Country ERP ', startAddress: 'A226', endAddress: 'C226' }])
  })

  it('preserves whitespace-only sheet names across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'whitespace-only-sheet' },
      sheets: [
        {
          id: 'blank-named-sheet',
          name: ' ',
          order: 0,
          cells: [{ address: 'A1', value: 'Visible data' }],
          metadata: {
            merges: [{ sheetName: ' ', startAddress: 'A3', endAddress: 'B3' }],
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const workbookXml = readXlsxTestZipText(exported, 'xl/workbook.xml')
    const roundTripped = importXlsx(exported, 'whitespace-only-sheet.xlsx')

    expect(workbookXml).toContain('<sheet name=" "')
    expect(workbookXml).not.toContain('<sheet name="Sheet1"')
    expect(roundTripped.snapshot.sheets[0]?.name).toBe(' ')
    expect(roundTripped.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', row: 0, col: 0, value: 'Visible data' }])
    expect(roundTripped.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: ' ', startAddress: 'A3', endAddress: 'B3' }])
  })
})
