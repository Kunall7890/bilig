import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { importCsv, importWorkbookFile, importXlsx, readImportedXlsxCellStyle } from '../index.js'
import { CSV_CONTENT_TYPE } from '@bilig/agent-api'

function buildWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()

  const sheet1 = XLSX.utils.aoa_to_sheet([
    [1, 2],
    [3, null],
  ])
  sheet1['C1'] = { t: 'n', f: 'A1+B1', z: '0.00' }
  sheet1['!ref'] = 'A1:C2'
  sheet1['!cols'] = [{ wpx: 120 }, { wch: 10 }, { wpx: 80 }]
  sheet1['!rows'] = [{ hpx: 30 }, { hpt: 18 }]
  sheet1['!merges'] = [{ s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }]

  const sheet2 = XLSX.utils.aoa_to_sheet([['hello'], [true]])
  sheet2['A1'] = {
    ...sheet2['A1'],
    c: [{ a: 'Greg', t: 'comment' }],
  }

  XLSX.utils.book_append_sheet(workbook, sheet1, 'Sheet1')
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Sheet2')
  workbook.Workbook = {
    Names: [{ Name: 'IgnoredName', Ref: 'Sheet1!$A$1' }],
  }

  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildPrepaidWorkbookFixture(shape: 'tracking' | 'daily'): Uint8Array {
  const workbook = XLSX.utils.book_new()
  if (shape === 'tracking') {
    const tracking = XLSX.utils.aoa_to_sheet([
      ['PREPAID EXPENSE TRACKING', null, null, null, null, null, null, null, null, null, null, null],
      [],
      [
        'ID',
        'Date Paid',
        'Vendor',
        'Description',
        'Category',
        'Total Amount',
        'Start Date',
        'End Date',
        'Life Months',
        'Monthly Amount',
        'Remaining Balance',
        'Status',
      ],
      ['PE001', 45292, 'Acme Insurance', 'Annual insurance premium', 'Insurance', 12000, 45292, 45657, null, null, null, 'Active'],
    ])
    tracking.I4 = { t: 'n', f: 'DATEDIF(G4,H4,"M")+1' }
    tracking.J4 = { t: 'n', f: 'F4/I4' }
    tracking.K4 = { t: 'n', f: "F4-SUMIF('Amortization Schedule'!$B:$B,A4,'Amortization Schedule'!$E:$E)" }
    tracking['!ref'] = 'A1:L4'
    tracking['!cols'] = [{ wpx: 132 }, { wpx: 96 }, { wpx: 142 }, { wpx: 210 }]
    tracking['!rows'] = [{ hpx: 30 }, {}, { hpx: 24 }]
    tracking['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }]

    const amortization = XLSX.utils.aoa_to_sheet([
      ['AMORTIZATION SCHEDULE', null, null, null, null, null],
      [],
      ['Month', 'Prepaid ID', 'Description', 'Monthly Amount', 'Cumulative Amortized', 'Remaining Balance'],
      ['Jan 2024', 'PE001', 'Annual insurance premium'],
    ])
    amortization.D4 = { t: 'n', f: "VLOOKUP(B4,'Prepaid Tracking'!A:J,10,FALSE())" }
    amortization.E4 = { t: 'n', f: 'D4' }
    amortization.F4 = { t: 'n', f: "VLOOKUP(B4,'Prepaid Tracking'!A:F,6,FALSE())-E4" }
    amortization['!ref'] = 'A1:F4'
    amortization['!cols'] = [{ wpx: 112 }, { wpx: 96 }, { wpx: 210 }, { wpx: 126 }, { wpx: 148 }, { wpx: 138 }]
    amortization['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }]
    XLSX.utils.book_append_sheet(workbook, tracking, 'Prepaid Tracking')
    XLSX.utils.book_append_sheet(workbook, amortization, 'Amortization Schedule')
    return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
  }

  const daily = XLSX.utils.aoa_to_sheet([
    ['Daily Prepaid Schedule', null, null, null, null, null, null, null, null],
    ['Vendor', 'Description', 'Start Date', 'End Date', 'Total Amount', 'Jan 2026', 'Feb 2026', '2026 Amortized', 'Remaining Balance'],
    ['TenantWorks', 'Facilities platform', 46054, 46234, 6600],
  ])
  daily.F3 = { t: 'n', f: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)' }
  daily.G3 = { t: 'n', f: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,2,1),0))-MAX($C3,DATE(2026,2,1))+1)/($D3-$C3+1),0),2)' }
  daily.H3 = { t: 'n', f: 'ROUND(SUM(F3:G3),2)' }
  daily.I3 = { t: 'n', f: 'ROUND(E3-H3,2)' }
  daily['!ref'] = 'A1:I3'
  daily['!cols'] = [
    { wpx: 168 },
    { wpx: 190 },
    { wpx: 104 },
    { wpx: 104 },
    { wpx: 118 },
    { wpx: 96 },
    { wpx: 96 },
    { wpx: 134 },
    { wpx: 138 },
  ]
  daily['!rows'] = [{ hpx: 30 }, { hpx: 24 }]
  daily['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }]
  XLSX.utils.book_append_sheet(workbook, daily, 'Daily Prepaids')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

describe('excel import', () => {
  it('imports sheets, formulas, dimensions, and warnings from xlsx bytes', () => {
    const imported = importXlsx(buildWorkbook(), 'Quarterly Report.xlsx')

    expect(imported.workbookName).toBe('Quarterly Report')
    expect(imported.sheetNames).toEqual(['Sheet1', 'Sheet2'])
    expect(imported.snapshot.workbook.name).toBe('Quarterly Report')
    expect(imported.snapshot.sheets).toHaveLength(2)

    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: 'Sheet1',
      metadata: {
        columns: [
          { index: 0, size: 120 },
          { index: 1, size: 65 },
          { index: 2, size: 80 },
        ],
        rows: [
          { index: 0, size: 30 },
          { index: 1, size: 18 },
        ],
        merges: [{ sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'B4' }],
      },
    })
    expect(imported.snapshot.sheets[0]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A1', value: 1 })]))
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: 'C1', formula: 'A1+B1', format: '0.00' })]),
    )
    expect(imported.snapshot.sheets[1]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A1', value: 'hello' })]))
    expect(imported.snapshot.sheets[1]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A2', value: true })]))

    expect(imported.warnings).toEqual(['Defined names were ignored during XLSX import.', 'Cell comments were ignored during XLSX import.'])
    expect(imported.preview.workbookName).toBe('Quarterly Report')
    expect(imported.preview.sheetCount).toBe(2)
    expect(imported.preview.sheets).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: 'Sheet1',
          rowCount: 2,
          columnCount: 3,
          nonEmptyCellCount: 4,
          previewRows: [
            ['1', '2', '=A1+B1'],
            ['3', '', ''],
          ],
        }),
      ]),
    )
  })

  it('maps imported xlsx styles into Bilig style records', () => {
    expect(
      readImportedXlsxCellStyle({
        patternType: 'solid',
        fgColor: { rgb: '1D3989' },
        font: {
          name: 'Aptos',
          sz: 12,
          bold: true,
          italic: true,
          underline: true,
          color: { rgb: 'FFFFFFFF' },
        },
        alignment: {
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
          indent: 1,
        },
        border: {
          bottom: {
            style: 'thin',
            color: { rgb: 'FF000000' },
          },
        },
      }),
    ).toEqual({
      fill: { backgroundColor: '#1d3989' },
      font: {
        family: 'Aptos',
        size: 12,
        bold: true,
        italic: true,
        underline: true,
        color: '#ffffff',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrap: true,
        indent: 1,
      },
      borders: {
        bottom: {
          style: 'solid',
          weight: 'thin',
          color: '#000000',
        },
      },
    })
  })

  it('imports multiple prepaid-amortization workbook shapes without file-specific dispatch', () => {
    const tracking = importXlsx(buildPrepaidWorkbookFixture('tracking'), 'tracking-prepaids.xlsx')
    expect(tracking.sheetNames).toEqual(['Prepaid Tracking', 'Amortization Schedule'])
    expect(tracking.snapshot.sheets[0]).toMatchObject({
      name: 'Prepaid Tracking',
      metadata: {
        columns: expect.arrayContaining([{ id: 'col:0', index: 0, size: 132 }]),
        rows: expect.arrayContaining([{ id: 'row:0', index: 0, size: 30 }]),
        merges: [{ sheetName: 'Prepaid Tracking', startAddress: 'A1', endAddress: 'L1' }],
      },
    })
    expect(tracking.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A4', value: 'PE001' }),
        expect.objectContaining({
          address: 'K4',
          formula: "F4-SUMIF('Amortization Schedule'!$B:$B,A4,'Amortization Schedule'!$E:$E)",
        }),
      ]),
    )

    const daily = importXlsx(buildPrepaidWorkbookFixture('daily'), 'daily-prepaids.xlsx')
    expect(daily.sheetNames).toEqual(['Daily Prepaids'])
    expect(daily.snapshot.sheets[0]).toMatchObject({
      name: 'Daily Prepaids',
      metadata: {
        columns: expect.arrayContaining([{ id: 'col:0', index: 0, size: 168 }]),
        rows: expect.arrayContaining([{ id: 'row:0', index: 0, size: 30 }]),
        merges: [{ sheetName: 'Daily Prepaids', startAddress: 'A1', endAddress: 'I1' }],
      },
    })
    expect(daily.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A3', value: 'TenantWorks' }),
        expect.objectContaining({
          address: 'F3',
          formula: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)',
        }),
      ]),
    )
  })

  it('imports csv files into a single-sheet workbook preview', () => {
    const imported = importCsv('Name,Value\nalpha,12\nbeta,=A2', 'metrics.csv')

    expect(imported.workbookName).toBe('metrics')
    expect(imported.sheetNames).toEqual(['metrics'])
    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: 'metrics',
      cells: [
        { address: 'A1', value: 'Name' },
        { address: 'B1', value: 'Value' },
        { address: 'A2', value: 'alpha' },
        { address: 'B2', value: 12 },
        { address: 'A3', value: 'beta' },
        { address: 'B3', formula: 'A2' },
      ],
    })
    expect(imported.preview).toMatchObject({
      workbookName: 'metrics',
      sheetCount: 1,
      sheets: [
        {
          name: 'metrics',
          rowCount: 3,
          columnCount: 2,
          nonEmptyCellCount: 6,
          previewRows: [
            ['Name', 'Value'],
            ['alpha', '12'],
            ['beta', '=A2'],
          ],
        },
      ],
    })
  })

  it('dispatches workbook import by content type', () => {
    const imported = importWorkbookFile(new TextEncoder().encode('A,B\n1,2'), 'dispatch.csv', CSV_CONTENT_TYPE)

    expect(imported.workbookName).toBe('dispatch')
    expect(imported.sheetNames).toEqual(['dispatch'])
  })
})
