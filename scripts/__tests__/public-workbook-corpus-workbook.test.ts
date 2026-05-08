import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx } from '../../packages/excel-import/src/index.js'
import { ValueTag } from '../../packages/protocol/src/enums.js'
import type { WorkbookSnapshot } from '../../packages/protocol/src/types.js'
import { extractFormulaOracles, inspectWorkbookFootprint } from '../public-workbook-corpus-workbook.ts'

describe('public workbook corpus workbook helpers', () => {
  it('extracts formula oracles from broad sparse worksheet refs', () => {
    const oracles = extractFormulaOracles(buildBroadSparseWorkbookBytes())

    expect(oracles).toEqual([
      {
        sheetName: 'Sparse',
        address: 'XFD512',
        expected: { tag: ValueTag.Number, value: 42 },
      },
    ])
  }, 15_000)

  it('records explicit used ranges from actual populated cells instead of broad worksheet refs', () => {
    const footprint = inspectWorkbookFootprint(buildBroadSparseWorkbookBytes(), 'sparse.xlsx')

    expect(footprint.workbookMetadata.dimensions).toEqual([
      {
        sheetName: 'Sparse',
        rowCount: 512,
        columnCount: 16_384,
        nonEmptyCellCount: 1,
        usedRange: { startRow: 511, startColumn: 16_383, endRow: 511, endColumn: 16_383 },
      },
    ])
  }, 15_000)

  it('counts raw XLSX pivot table parts even when semantic pivot import is unavailable', () => {
    const footprint = inspectWorkbookFootprint(exportXlsx(buildPivotWorkbookSnapshot()), 'raw-pivot.xlsx')

    expect(footprint.featureCounts.pivotCount).toBe(1)
  }, 15_000)
})

function buildBroadSparseWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet: XLSX.WorkSheet = {
    XFD512: { t: 'n', f: '40+2', v: 42 },
    '!ref': 'A1:XFD512',
  }
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sparse')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildPivotWorkbookSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'raw-pivot',
      metadata: {
        pivots: [
          {
            name: 'RevenuePivot',
            sheetName: 'Pivot',
            address: 'A1',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Revenue', summarizeBy: 'sum' }],
            rows: 3,
            cols: 2,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Revenue' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 12 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 8 },
        ],
      },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Sum of Revenue' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 12 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 8 },
        ],
      },
    ],
  }
}
