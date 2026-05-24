import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('MS-OI29500 calcPr and calcChain import', () => {
  it('imports calcPr force-full-calc/calcId and calcChain provenance into formula audit metadata', () => {
    const imported = importXlsx(buildCalcChainWorkbookBytes(), 'ms-oi29500-calc-chain.xlsx')
    const formulaAudit = imported.snapshot.workbook.metadata?.formulaAudit

    expect(imported.snapshot.workbook.metadata?.calculationSettings).toEqual({
      mode: 'manual',
      compatibilityMode: 'excel-modern',
      fullCalcOnLoad: true,
      forceFullCalc: true,
      calcCompleted: false,
      calcId: 191029,
    })
    expect(formulaAudit).toMatchObject({
      calcChain: {
        packagePath: 'xl/calcChain.xml',
        cells: [
          { sheetIndex: 1, sheetName: 'Model', address: 'B2' },
          { sheetIndex: 1, sheetName: 'Model', address: 'C2' },
        ],
      },
      diagnostics: expect.arrayContaining([
        expect.objectContaining({
          code: 'manual-calc-mode',
          clause: '18.2.2',
        }),
        expect.objectContaining({
          code: 'force-full-calc',
          clause: '18.2.2',
        }),
        expect.objectContaining({
          code: 'calc-not-completed',
          clause: '18.2.2',
        }),
      ]),
    })

    const exportedWorkbookXml = workbookXml(exportXlsx(imported.snapshot))
    expect(exportedWorkbookXml).toContain('calcId="191029"')
    expect(exportedWorkbookXml).toContain('forceFullCalc="1"')
    expect(exportedWorkbookXml).toContain('calcCompleted="0"')
  })

  it('resolves calcChain cell sheet names from workbook sheet ids instead of tab order', () => {
    const imported = importXlsx(buildCalcChainWorkbookBytesWithSheetIdGap(), 'ms-oi29500-calc-chain-sheet-id-gap.xlsx')

    expect(imported.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report'])
    expect(imported.snapshot.workbook.metadata?.formulaAudit?.calcChain?.cells).toEqual([
      { sheetIndex: 2, sheetName: 'Inputs', address: 'A1' },
      { sheetIndex: 3, sheetName: 'Report', address: 'A1' },
    ])
  })
})

function buildCalcChainWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Input', 'Gross', 'Net'],
    [10, { f: 'A2*2', v: 20 }, { f: 'B2-1', v: 19 }],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  const calcPrXml = '<calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" forceFullCalc="1" calcCompleted="0"/>'
  zip['xl/workbook.xml'] = strToU8(
    /<calcPr\b/u.test(sourceWorkbookXml)
      ? sourceWorkbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/u, calcPrXml)
      : sourceWorkbookXml.replace('</workbook>', `${calcPrXml}</workbook>`),
  )
  zip['xl/calcChain.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<c r="B2" i="1"/>',
      '<c r="C2"/>',
      '</calcChain>',
    ].join(''),
  )
  return zipSync(zip)
}

function buildCalcChainWorkbookBytesWithSheetIdGap(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[{ f: '10+1', v: 11 }]]), 'Inputs')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[{ f: 'Inputs!A1+1', v: 12 }]]), 'Report')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  zip['xl/workbook.xml'] = strToU8(
    sourceWorkbookXml
      .replace(/name="Inputs" sheetId="\d+"/u, 'name="Inputs" sheetId="2"')
      .replace(/name="Report" sheetId="\d+"/u, 'name="Report" sheetId="3"'),
  )
  zip['xl/calcChain.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<c r="A1" i="2"/>',
      '<c r="A1" i="3"/>',
      '</calcChain>',
    ].join(''),
  )
  return zipSync(zip)
}

function workbookXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array())
}
