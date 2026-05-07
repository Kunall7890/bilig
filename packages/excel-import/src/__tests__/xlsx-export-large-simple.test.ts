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

  it('does not let broad column metadata expand the exported worksheet scan range', () => {
    const start = performance.now()
    const exported = exportXlsx(buildBroadColumnMetadataSnapshot())
    const durationMs = performance.now() - start
    const workbook = XLSX.read(exported, { type: 'array', cellFormula: true, cellText: false, cellDates: false })

    expect(durationMs).toBeLessThan(1_500 * readBenchmarkTolerance())
    expect(workbook.Sheets['Wide']?.['!ref']).toBe('A3040')
  }, 15_000)
})

function readBenchmarkTolerance(): number {
  const raw = process.env.BILIG_BENCH_TOLERANCE
  if (!raw) {
    return 1
  }
  const tolerance = Number(raw)
  return Number.isFinite(tolerance) && tolerance > 0 ? tolerance : 1
}

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

function buildBroadColumnMetadataSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'wide-column-metadata' },
    sheets: [
      {
        id: 1,
        name: 'Wide',
        order: 0,
        cells: [{ address: 'A3040', value: 1 }],
        metadata: {
          columns: Array.from({ length: 16_384 }, (_entry, index) => ({
            id: `column:${String(index)}`,
            index,
            size: 64,
          })),
        },
      },
    ],
  }
}
