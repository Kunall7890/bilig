import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { dataTableFormulasWarning, exportXlsx, importXlsx } from '../index.js'

describe('xlsx data table formulas roundtrip', () => {
  it('preserves what-if data-table formula metadata in worksheet cells', () => {
    const source = buildWorkbookWithDataTableFormula()

    const imported = importXlsx(source, 'what-if-data-table.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.warnings).toEqual([dataTableFormulasWarning])
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [
        {
          address: 'B2',
          formulaXml: '<f t="dataTable" ref="B2:D4" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/>',
        },
      ],
    })
    expect(dataTableFormulaXml(exported)).toEqual(dataTableFormulaXml(source))
    expect(cellXml(exported, 'B2')).toContain('<v>42</v>')
  })
})

function buildWorkbookWithDataTableFormula(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sheetXml.replace(
      /<c\b[^>]*\br=(["'])B2\1[^>]*>[\s\S]*?<\/c>/u,
      '<c r="B2"><f t="dataTable" ref="B2:D4" dt2D="1" dtr="1" r1="A1" r2="A2" ca="1"/><v>42</v></c>',
    ),
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'What-if table' },
    sheets: [
      {
        id: 1,
        name: 'Sensitivity',
        order: 0,
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'B2', value: 42 },
          { address: 'C2', value: 43 },
          { address: 'D4', value: 44 },
        ],
      },
    ],
  }
}

function dataTableFormulaXml(bytes: Uint8Array): string[] {
  return [...readZipText(bytes, 'xl/worksheets/sheet1.xml').matchAll(/<f\b[^>]*\bt=(["'])dataTable\1[^>]*(?:\/>|>[\s\S]*?<\/f>)/gu)].map(
    (match) => match[0],
  )
}

function cellXml(bytes: Uint8Array, address: string): string {
  return readZipText(bytes, 'xl/worksheets/sheet1.xml').match(new RegExp(`<c[^>]* r="${address}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0] ?? ''
}

function readZipText(bytes: Uint8Array, path: string): string {
  return readZipTextFromZip(unzipSync(bytes), path)
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}
