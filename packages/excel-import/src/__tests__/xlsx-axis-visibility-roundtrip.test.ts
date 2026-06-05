import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { describe, expect, it } from 'vitest'
import { strFromU8, unzipSync } from 'fflate'

import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX hidden row and column roundtrip', () => {
  it('preserves hidden row and column state in metadata and exported worksheet XML', () => {
    const imported = importXlsx(buildHiddenAxisWorkbookBytes(), 'hidden-axis.xlsx')
    const metadata = imported.snapshot.sheets[0]?.metadata

    expect(metadata?.rows).toContainEqual(
      expect.objectContaining({
        id: 'row:2',
        index: 2,
        hidden: true,
      }),
    )
    expect(metadata?.columns).toContainEqual(
      expect.objectContaining({
        id: 'col:2',
        index: 2,
        hidden: true,
      }),
    )
    expect(metadata?.rowMetadata).toContainEqual(
      expect.objectContaining({
        start: 2,
        count: 1,
        hidden: true,
      }),
    )
    expect(metadata?.columnMetadata).toContainEqual(
      expect.objectContaining({
        start: 2,
        count: 1,
        hidden: true,
      }),
    )

    const exportedSheetXml = strFromU8(unzipSync(exportXlsx(imported.snapshot))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    expect(exportedSheetXml).toMatch(/<col\b(?=[^>]*\bmin="3")(?=[^>]*\bmax="3")(?=[^>]*\bhidden="1")[^>]*\/>/u)
    expect(exportedSheetXml).toMatch(/<row\b(?=[^>]*\br="3")(?=[^>]*\bhidden="1")[^>]*(?:\/|>)/u)
  })
})

function buildHiddenAxisWorkbookBytes(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Visibility',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'visible a' },
          { address: 'B1', row: 0, col: 1, value: 'visible b' },
          { address: 'C1', row: 0, col: 2, value: 'hidden c' },
          { address: 'D1', row: 0, col: 3, value: 'visible d' },
          { address: 'A2', row: 1, col: 0, value: 1 },
          { address: 'B2', row: 1, col: 1, value: 2 },
          { address: 'C2', row: 1, col: 2, value: 3 },
          { address: 'D2', row: 1, col: 3, value: 4 },
          { address: 'A3', row: 2, col: 0, value: 'hidden row' },
          { address: 'B3', row: 2, col: 1, value: 'hidden row' },
          { address: 'C3', row: 2, col: 2, value: 'hidden row' },
          { address: 'D3', row: 2, col: 3, value: 'hidden row' },
        ],
        rows: [{ index: 2, hidden: true, size: 18 }],
        columns: [{ index: 2, hidden: true, size: 96 }],
      },
    ],
  })
}
