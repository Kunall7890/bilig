import { describe, expect, it } from 'vitest'

import { shouldUseDenseSheetJsParse } from '../xlsx-dense-sheetjs-parse.js'

const encoder = new TextEncoder()

function zipEntry(xml: string): Uint8Array {
  return encoder.encode(xml)
}

describe('dense SheetJS parse routing', () => {
  it('uses native XLSX range decoding to gate wide worksheet dimensions', () => {
    const workbookZip = {
      'xl/worksheets/sheet1.xml': zipEntry('<worksheet><dimension ref="A1:DX10"/></worksheet>'),
    }

    expect(
      shouldUseDenseSheetJsParse(new Uint8Array(1024), workbookZip, {
        maxColumnCount: 128,
        minByteLength: 1024,
      }),
    ).toBe(true)

    expect(
      shouldUseDenseSheetJsParse(new Uint8Array(1024), workbookZip, {
        maxColumnCount: 127,
        minByteLength: 1024,
      }),
    ).toBe(false)
  })

  it('keeps single-cell dimensions eligible and ignores malformed dimension refs', () => {
    expect(
      shouldUseDenseSheetJsParse(
        new Uint8Array(1024),
        {
          'xl/worksheets/sheet1.xml': zipEntry('<worksheet><dimension ref="C7"/></worksheet>'),
        },
        {
          maxColumnCount: 3,
          minByteLength: 1024,
        },
      ),
    ).toBe(true)

    expect(
      shouldUseDenseSheetJsParse(
        new Uint8Array(1024),
        {
          'xl/worksheets/sheet1.xml': zipEntry('<worksheet><dimension ref="not-a-range"/></worksheet>'),
        },
        {
          maxColumnCount: 128,
          minByteLength: 1024,
        },
      ),
    ).toBe(false)
  })
})
