import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('large simple XLSX export', () => {
  it('round-trips large value-heavy sheets without the style writer hot path', () => {
    const exported = exportXlsx(buildLargeSimpleSnapshot())
    const imported = importXlsx(exported, 'large-simple.xlsx')
    const sheet = imported.snapshot.sheets[0]
    const styleRange = sheet?.metadata?.styleRanges?.find((entry) => entry.range.startAddress === 'A1')
    const style = imported.snapshot.workbook.metadata?.styles?.find((entry) => entry.id === styleRange?.styleId)

    expect(sheet?.cells).toHaveLength(100_000)
    expect(sheet?.cells.find((cell) => cell.address === 'A2')).toMatchObject({ value: 50, format: '0.00' })
    expect(sheet?.metadata?.merges).toEqual([{ sheetName: 'Large', startAddress: 'A1', endAddress: 'B1' }])
    expect(style?.fill?.backgroundColor).toBe('#ffcc00')
    expect(style?.font?.bold).toBe(true)
  }, 15_000)
})

function buildLargeSimpleSnapshot(): WorkbookSnapshot {
  const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
  for (let row = 0; row < 2_000; row += 1) {
    for (let column = 0; column < 50; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column })
      cells.push({
        address,
        value: row * 50 + column,
        ...(address === 'A2' ? { format: '0.00' } : {}),
      })
    }
  }
  return {
    version: 1,
    workbook: {
      name: 'large-simple',
      metadata: {
        styles: [
          {
            id: 'header',
            fill: { backgroundColor: '#ffcc00' },
            font: { bold: true },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Large',
        order: 0,
        cells,
        metadata: {
          columns: [{ id: 'column:0', index: 0, size: 120 }],
          rows: [{ id: 'row:0', index: 0, size: 28 }],
          merges: [{ sheetName: 'Large', startAddress: 'A1', endAddress: 'B1' }],
          styleRanges: [{ range: { sheetName: 'Large', startAddress: 'A1', endAddress: 'B1' }, styleId: 'header' }],
        },
      },
    ],
  }
}
