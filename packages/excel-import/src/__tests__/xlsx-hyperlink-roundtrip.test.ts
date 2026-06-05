import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('hyperlink roundtrip', () => {
  it('preserves external and internal cell hyperlinks through import and export', () => {
    const imported = importXlsx(buildHyperlinkWorkbookBytes(), 'hyperlinks.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.hyperlinks).toEqual([
      {
        sheetName: 'Inputs',
        address: 'A1',
        target: 'https://example.com/report',
        tooltip: 'Open report',
        display: 'Open report',
      },
      {
        sheetName: 'Inputs',
        address: 'B2',
        target: '#Summary!A1',
        tooltip: 'Jump to summary',
        display: 'Summary',
      },
    ])

    const reimported = importXlsx(exportXlsx(imported.snapshot), 'hyperlinks-roundtrip.xlsx')

    expect(reimported.snapshot.sheets[0]?.metadata?.hyperlinks).toEqual(imported.snapshot.sheets[0]?.metadata?.hyperlinks)
  })
})

function buildHyperlinkWorkbookBytes(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Inputs',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Open report' },
            { address: 'B2', row: 1, col: 1, value: 'Summary' },
          ],
        },
        {
          name: 'Summary',
          cells: [{ address: 'A1', row: 0, col: 0, value: 'Destination' }],
        },
      ],
    }),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array()).replace(
      '</worksheet>',
      [
        '<hyperlinks>',
        '<hyperlink ref="A1" r:id="rIdHyperlink1" tooltip="Open report" display="Open report"/>',
        '<hyperlink ref="B2" location="Summary!A1" tooltip="Jump to summary" display="Summary"/>',
        '</hyperlinks>',
        '</worksheet>',
      ].join(''),
    ),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rIdHyperlink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/report" TargetMode="External"/>',
      '</Relationships>',
    ].join(''),
  )
  return zipSync(zip)
}
