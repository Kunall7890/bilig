import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'

import { exportXlsx, importXlsx } from '../index.js'

const sharedRichStringXml = [
  '<si>',
  '<r><rPr><b/><sz val="10"/><color rgb="FF1F4E79"/><rFont val="Helv"/></rPr><t>Important:</t></r>',
  '<r><rPr><i/><sz val="10"/><color rgb="FFC00000"/><rFont val="Helv"/></rPr><t xml:space="preserve"> Before signing off</t></r>',
  '</si>',
].join('')

const inlineRichStringXml = [
  '<is>',
  '<r><rPr><u/><sz val="11"/><color rgb="FF008000"/><rFont val="Calibri"/></rPr><t>Revenue</t></r>',
  '<r><rPr><sz val="11"/><rFont val="Calibri"/></rPr><t xml:space="preserve"> sensitivity</t></r>',
  '</is>',
].join('')

describe('cell rich text roundtrip', () => {
  it('preserves shared and inline rich text runs across XLSX round trips', () => {
    const source = buildRichTextCellWorkbookBytes()

    const imported = importXlsx(source, 'cell-rich-text.xlsx')
    const exported = exportXlsx(imported.snapshot)
    const exportedSummary = readRichTextSummary(exported)

    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A1', value: 'Important: Before signing off' }),
        expect.objectContaining({ address: 'B1', value: 'Revenue sensitivity' }),
      ]),
    )
    expect(exportedSummary.richSharedStringCount).toBe(1)
    expect(exportedSummary.richInlineStringCount).toBe(1)
    expect(exportedSummary.runCount).toBe(4)
    expect(exportedSummary.runPropertyCount).toBe(4)
    expect(exportedSummary.sharedStringsXml).toContain(sharedRichStringXml)
    expect(exportedSummary.worksheetXml).toContain(inlineRichStringXml)
    expect(exportedSummary.worksheetXml).toContain('<c r="A1" t="s"><v>0</v></c>')
    expect(exportedSummary.worksheetXml).toContain(`<c r="B1" t="inlineStr">${inlineRichStringXml}</c>`)
  })

  it('does not restore imported rich text runs after cell text changes', () => {
    const imported = importXlsx(buildRichTextCellWorkbookBytes(), 'cell-rich-text.xlsx')
    const sheet = imported.snapshot.sheets[0]
    const sharedCell = sheet?.cells.find((cell) => cell.address === 'A1')
    const inlineCell = sheet?.cells.find((cell) => cell.address === 'B1')
    if (!sharedCell || !inlineCell) {
      throw new Error('Fixture import did not produce the expected rich text cells.')
    }
    sharedCell.value = 'Changed shared label'
    inlineCell.value = 'Changed inline label'

    const exportedSummary = readRichTextSummary(exportXlsx(imported.snapshot))

    expect(exportedSummary.richSharedStringCount).toBe(0)
    expect(exportedSummary.richInlineStringCount).toBe(0)
    expect(exportedSummary.runCount).toBe(0)
    expect(exportedSummary.worksheetXml).toContain('Changed shared label')
    expect(exportedSummary.worksheetXml).toContain('Changed inline label')
  })
})

interface RichTextSummary {
  readonly worksheetXml: string
  readonly sharedStringsXml: string
  readonly richSharedStringCount: number
  readonly richInlineStringCount: number
  readonly runCount: number
  readonly runPropertyCount: number
}

function buildRichTextCellWorkbookBytes(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Labels',
        cells: [
          { address: 'A1', row: 0, col: 0, sharedStringIndex: 0 },
          { address: 'B1', row: 0, col: 1, inlineStringXml: inlineRichStringXml },
        ],
      },
    ],
    sharedStringsXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">',
      sharedRichStringXml,
      '</sst>',
    ].join(''),
  })
}

function readRichTextSummary(bytes: Uint8Array): RichTextSummary {
  const zip = unzipSync(bytes)
  const worksheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  const sharedStringsXml = strFromU8(zip['xl/sharedStrings.xml'] ?? new Uint8Array())
  const sharedStringItems = sharedStringsXml.match(/<si\b[^>]*>[\s\S]*?<\/si>/gu) ?? []
  const inlineStringItems = worksheetXml.match(/<is\b[^>]*>[\s\S]*?<\/is>/gu) ?? []
  return {
    worksheetXml,
    sharedStringsXml,
    richSharedStringCount: sharedStringItems.filter((entry) => entry.includes('<r>')).length,
    richInlineStringCount: inlineStringItems.filter((entry) => entry.includes('<r>')).length,
    runCount: (sharedStringsXml.match(/<r>/gu)?.length ?? 0) + (worksheetXml.match(/<r>/gu)?.length ?? 0),
    runPropertyCount: (sharedStringsXml.match(/<rPr>/gu)?.length ?? 0) + (worksheetXml.match(/<rPr>/gu)?.length ?? 0),
  }
}
