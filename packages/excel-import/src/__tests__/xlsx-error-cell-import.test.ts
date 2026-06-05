import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { describe, expect, it } from 'vitest'

import { importXlsx } from '../index.js'

describe('github issue #42 xlsx literal error cell import', () => {
  it('preserves literal Excel error cells as display text instead of numeric error codes', () => {
    const imported = importXlsx(buildLiteralErrorWorkbookBytes(), 'literal-error-cells.xlsx')
    const cells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell.value]) ?? [])

    expect(cells.get('A2')).toBe('#N/A')
    expect(cells.get('B2')).toBe('#DIV/0!')
    expect(cells.get('C2')).toBe('#REF!')
    expect(cells.get('D2')).toBe('#VALUE!')
    expect(cells.get('E2')).toBe('#NAME?')
    expect(cells.get('F2')).toBe('#NUM!')

    expect(imported.warnings).toEqual([])
    expect(imported.preview.sheets[0]?.previewRows[1]).toEqual(['#N/A', '#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#NUM!'])
  })
})

function buildLiteralErrorWorkbookBytes(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Errors',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'NA' },
          { address: 'B1', row: 0, col: 1, value: 'DIV' },
          { address: 'C1', row: 0, col: 2, value: 'REF' },
          { address: 'D1', row: 0, col: 3, value: 'VALUE' },
          { address: 'E1', row: 0, col: 4, value: 'NAME' },
          { address: 'F1', row: 0, col: 5, value: 'NUM' },
          { address: 'A2', row: 1, col: 0, error: '#N/A' },
          { address: 'B2', row: 1, col: 1, error: '#DIV/0!' },
          { address: 'C2', row: 1, col: 2, error: '#REF!' },
          { address: 'D2', row: 1, col: 3, error: '#VALUE!' },
          { address: 'E2', row: 1, col: 4, error: '#NAME?' },
          { address: 'F2', row: 1, col: 5, error: '#NUM!' },
        ],
      },
    ],
  })
}
