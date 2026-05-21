import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { dataTableFormulasWarning, exportXlsx, importXlsx } from '../index.js'

describe('xlsx data table formulas roundtrip', () => {
  it('lowers two-variable data-table outputs to calculable formulas while preserving native metadata', () => {
    const source = buildWorkbookWithTwoVariableDataTableFormula()

    const imported = importXlsx(source, 'what-if-data-table.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.warnings).not.toContain(dataTableFormulasWarning)
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [
        {
          address: 'C3',
          formulaXml: '<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2"/>',
        },
      ],
    })
    expect(snapshotCell(imported.snapshot, 'C3')).toMatchObject({
      address: 'C3',
      formula: 'MULTIPLE.OPERATIONS(B2,A1,C2,A2,B3)',
      value: 40,
    })
    expect(snapshotCell(imported.snapshot, 'D3')).toMatchObject({
      address: 'D3',
      formula: 'MULTIPLE.OPERATIONS(B2,A1,D2,A2,B3)',
      value: 60,
    })
    expect(snapshotCell(imported.snapshot, 'C4')).toMatchObject({
      address: 'C4',
      formula: 'MULTIPLE.OPERATIONS(B2,A1,C2,A2,B4)',
      value: 60,
    })
    expect(snapshotCell(imported.snapshot, 'D4')).toMatchObject({
      address: 'D4',
      formula: 'MULTIPLE.OPERATIONS(B2,A1,D2,A2,B4)',
      value: 90,
    })

    expect(dataTableFormulaXml(exported)).toEqual(dataTableFormulaXml(source))
    expect(cellXml(exported, 'C3')).toContain('<v>40</v>')
    for (const address of ['D3', 'C4', 'D4']) {
      expect(cellXml(exported, address)).not.toContain('<f')
      expect(cellXml(exported, address)).not.toContain('MULTIPLE.OPERATIONS')
    }
  })

  it('lowers one-variable row-input and column-input data-table outputs to calculable formulas', () => {
    const source = buildWorkbookWithOneVariableDataTableFormulas()

    const imported = importXlsx(source, 'one-variable-data-table.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.warnings).not.toContain(dataTableFormulasWarning)
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [
        {
          address: 'C2',
          formulaXml: '<f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1"/>',
        },
        {
          address: 'B6',
          formulaXml: '<f t="dataTable" ref="B6:B8" dt2D="0" dtr="0" r1="A1"/>',
        },
      ],
    })
    expect(snapshotCell(imported.snapshot, 'C2')).toMatchObject({
      address: 'C2',
      formula: 'MULTIPLE.OPERATIONS(B2,A1,C1)',
      value: 30,
    })
    expect(snapshotCell(imported.snapshot, 'D2')).toMatchObject({
      address: 'D2',
      formula: 'MULTIPLE.OPERATIONS(B2,A1,D1)',
      value: 40,
    })
    expect(snapshotCell(imported.snapshot, 'B6')).toMatchObject({
      address: 'B6',
      formula: 'MULTIPLE.OPERATIONS(B5,A1,A6)',
      value: 20,
    })
    expect(snapshotCell(imported.snapshot, 'B7')).toMatchObject({
      address: 'B7',
      formula: 'MULTIPLE.OPERATIONS(B5,A1,A7)',
      value: 30,
    })
    expect(snapshotCell(imported.snapshot, 'B8')).toMatchObject({
      address: 'B8',
      formula: 'MULTIPLE.OPERATIONS(B5,A1,A8)',
      value: 40,
    })

    expect(dataTableFormulaXml(exported)).toEqual(dataTableFormulaXml(source))
    for (const address of ['D2', 'B7', 'B8']) {
      expect(cellXml(exported, address)).not.toContain('<f')
      expect(cellXml(exported, address)).not.toContain('MULTIPLE.OPERATIONS')
    }
  })

  it('keeps malformed data-table formulas as preserved metadata with a warning', () => {
    const source = buildWorkbookWithMalformedDataTableFormula()

    const imported = importXlsx(source, 'malformed-data-table.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.warnings).toContain(dataTableFormulasWarning)
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [
        {
          address: 'B2',
          formulaXml: '<f t="dataTable" ref="B2:D4" ca="1"/>',
        },
      ],
    })
    expect(snapshotCell(imported.snapshot, 'B2')).toMatchObject({ address: 'B2', value: 42 })
    expect(snapshotCell(imported.snapshot, 'B2')?.formula).toBeUndefined()
    expect(dataTableFormulaXml(exported)).toEqual(dataTableFormulaXml(source))
  })
})

function buildWorkbookWithTwoVariableDataTableFormula(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildTwoVariableWorkbook()))
  const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sheetXml.replace(
      /<c\b[^>]*\br=(["'])C3\1[^>]*>[\s\S]*?<\/c>/u,
      '<c r="C3"><f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2"/><v>40</v></c>',
    ),
  )
  return zipSync(zip)
}

function buildWorkbookWithOneVariableDataTableFormulas(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildOneVariableWorkbook()))
  const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sheetXml
      .replace(
        /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
        '<c r="C2"><f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1"/><v>30</v></c>',
      )
      .replace(
        /<c\b[^>]*\br=(["'])B6\1[^>]*>[\s\S]*?<\/c>/u,
        '<c r="B6"><f t="dataTable" ref="B6:B8" dt2D="0" dtr="0" r1="A1"/><v>20</v></c>',
      ),
  )
  return zipSync(zip)
}

function buildWorkbookWithMalformedDataTableFormula(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildMalformedWorkbook()))
  const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sheetXml.replace(/<c\b[^>]*\br=(["'])B2\1[^>]*>[\s\S]*?<\/c>/u, '<c r="B2"><f t="dataTable" ref="B2:D4" ca="1"/><v>42</v></c>'),
  )
  return zipSync(zip)
}

function buildTwoVariableWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'What-if table' },
    sheets: [
      {
        id: 1,
        name: 'Sensitivity',
        order: 0,
        cells: [
          { address: 'A1', value: 1 },
          { address: 'A2', value: 10 },
          { address: 'A3', formula: 'A1*A2', value: 10 },
          { address: 'B2', formula: 'A3', value: 10 },
          { address: 'C2', value: 2 },
          { address: 'D2', value: 3 },
          { address: 'B3', value: 20 },
          { address: 'C3', value: 40 },
          { address: 'D3', value: 60 },
          { address: 'B4', value: 30 },
          { address: 'C4', value: 60 },
          { address: 'D4', value: 90 },
        ],
      },
    ],
  }
}

function buildOneVariableWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'One-variable what-if tables' },
    sheets: [
      {
        id: 1,
        name: 'Sensitivity',
        order: 0,
        cells: [
          { address: 'A1', value: 1 },
          { address: 'A2', formula: 'A1*10', value: 10 },
          { address: 'B1', value: 2 },
          { address: 'C1', value: 3 },
          { address: 'D1', value: 4 },
          { address: 'B2', formula: 'A2', value: 10 },
          { address: 'C2', value: 30 },
          { address: 'D2', value: 40 },
          { address: 'A5', value: 1 },
          { address: 'B5', formula: 'A1*10', value: 10 },
          { address: 'A6', value: 2 },
          { address: 'B6', value: 20 },
          { address: 'A7', value: 3 },
          { address: 'B7', value: 30 },
          { address: 'A8', value: 4 },
          { address: 'B8', value: 40 },
        ],
      },
    ],
  }
}

function buildMalformedWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Malformed what-if table' },
    sheets: [
      {
        id: 1,
        name: 'Sensitivity',
        order: 0,
        cells: [
          { address: 'A1', value: 10 },
          { address: 'B2', value: 42 },
          { address: 'C2', value: 43 },
          { address: 'D4', value: 44 },
        ],
      },
    ],
  }
}

function snapshotCell(snapshot: WorkbookSnapshot, address: string): WorkbookSnapshot['sheets'][number]['cells'][number] | undefined {
  return snapshot.sheets[0]?.cells.find((cell) => cell.address === address)
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
