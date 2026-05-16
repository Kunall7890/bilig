import { describe, expect, it } from 'vitest'
import { isWorkbookChangeRange, normalizeWorkbookChangeRange } from '../workbook-change-range.js'

describe('workbook change range guards', () => {
  it('normalizes legacy cell ranges without explicit scope', () => {
    expect(
      normalizeWorkbookChangeRange({
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B2',
    })
  })

  it('preserves valid structural scope', () => {
    expect(
      normalizeWorkbookChangeRange({
        sheetName: 'Sheet1',
        startAddress: 'A3',
        endAddress: 'A4',
        scope: 'rows',
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A3',
      endAddress: 'A4',
      scope: 'rows',
    })
  })

  it('rejects malformed scope instead of silently downgrading to a cell range', () => {
    const malformed = {
      sheetName: 'Sheet1',
      startAddress: 'A3',
      endAddress: 'A4',
      scope: 'row-band',
    }

    expect(normalizeWorkbookChangeRange(malformed)).toBeNull()
    expect(isWorkbookChangeRange(malformed)).toBe(false)
  })
})
