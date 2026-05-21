import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { addExportCalculationSettingsToXlsxBytes } from '../xlsx-calculation-settings.js'
import { manualCalculationModeWarning, precisionAsDisplayedCalculationWarning } from '../index.js'
import { borrowXlsxZipByteSource, importXlsxFromZipByteSource } from '../xlsx-byte-source-import.js'

describe('XLSX byte-source import', () => {
  it('preserves reusable readRangeInto support on borrowed ZIP byte sources', () => {
    const source = new InstrumentedByteSource(new Uint8Array([1, 2, 3, 4]))
    const borrowed = borrowXlsxZipByteSource(source)
    const scratch = new Uint8Array(2)

    expect(Array.from(borrowed.readRangeInto?.(1, 3, scratch) ?? [])).toEqual([2, 3])

    expect(source.readIntoCount).toBe(1)
    expect(source.rangeCount).toBe(0)
  })

  it('preserves calculation warnings when byte-source import falls back to materialized SheetJS parsing', () => {
    const manual = importXlsxFromZipByteSource(
      new InstrumentedByteSource(buildCalculationWorkbookBytes({ mode: 'manual', compatibilityMode: 'excel-modern' })),
      'manual.xlsx',
      { attachSourceReaderForUntouchedExport: false },
    )
    const precisionAsDisplayed = importXlsxFromZipByteSource(
      new InstrumentedByteSource(
        buildCalculationWorkbookBytes({ mode: 'automatic', compatibilityMode: 'excel-modern', fullPrecision: false }),
      ),
      'precision-as-displayed.xlsx',
      { attachSourceReaderForUntouchedExport: false },
    )

    expect(manual.warnings).toContain(manualCalculationModeWarning)
    expect(precisionAsDisplayed.warnings).toContain(precisionAsDisplayedCalculationWarning)
  })
})

function buildCalculationWorkbookBytes(settings: {
  readonly mode: 'automatic' | 'manual'
  readonly compatibilityMode: 'excel-modern'
  readonly fullPrecision?: false
}): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Input', 'Value'],
    ['A', 1],
    ['B', 2],
    ['Total', null],
  ])
  sheet.B4 = { t: 'n', f: 'B2+B3', v: settings.mode === 'manual' ? 99 : 3 }
  sheet['!ref'] = 'A1:B4'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  return addExportCalculationSettingsToXlsxBytes(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), {
    version: 1,
    workbook: {
      name: 'calculation-warning-byte-source',
      metadata: { calculationSettings: settings },
    },
    sheets: [],
  })
}

class InstrumentedByteSource {
  readonly byteLength: number
  rangeCount = 0
  readIntoCount = 0

  constructor(private readonly bytes: Uint8Array) {
    this.byteLength = bytes.byteLength
  }

  readRange(start: number, end: number): Uint8Array {
    this.rangeCount += 1
    return this.bytes.subarray(start, end)
  }

  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array {
    this.readIntoCount += 1
    target.set(this.bytes.subarray(start, end), 0)
    return target.subarray(0, end - start)
  }
}
