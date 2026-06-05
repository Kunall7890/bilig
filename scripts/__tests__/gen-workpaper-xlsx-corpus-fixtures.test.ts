import {
  encodeCellAddress,
  readXlsxZipEntries,
  writeSimpleXlsxWorkbook,
  zipSourcePreservingEntries,
  type SimpleXlsxWorkbook,
} from '@bilig/xlsx'
import { describe, expect, it } from 'vitest'

import { xlsxZipEntryContentsEqual } from '../gen-workpaper-xlsx-corpus-fixtures.ts'

function workbookWithValue(value: string): SimpleXlsxWorkbook {
  return {
    sheets: [
      {
        name: 'Sheet1',
        cells: [
          {
            address: encodeCellAddress({ r: 0, c: 0 }),
            row: 0,
            col: 0,
            value,
          },
        ],
      },
    ],
  }
}

describe('WorkPaper XLSX corpus fixture generation', () => {
  it('compares generated XLSX fixtures by inflated entry contents instead of ZIP bytes', () => {
    const source = writeSimpleXlsxWorkbook(workbookWithValue('same payload'))
    const repacked = zipSourcePreservingEntries({ ...readXlsxZipEntries(source) }, new Map(), { dosTime: { time: 1, date: 33 } })

    expect(Buffer.from(repacked).equals(Buffer.from(source))).toBe(false)
    expect(xlsxZipEntryContentsEqual(source, repacked)).toBe(true)
  })

  it('rejects generated XLSX fixtures when workbook entry contents change', () => {
    const source = writeSimpleXlsxWorkbook(workbookWithValue('original payload'))
    const changed = writeSimpleXlsxWorkbook(workbookWithValue('changed payload'))

    expect(xlsxZipEntryContentsEqual(source, changed)).toBe(false)
  })
})
