import { describe, expect, it } from 'vitest'

import {
  appendLargeSimpleColumnMetadataFromBytes,
  appendLargeSimpleRowMetadataTagFromBytes,
  readLargeSimpleDrawingRelationshipIdTagFromBytes,
  readLargeSimpleSheetFormatPrTagFromBytes,
  readLargeSimpleTableRelationshipIdsFromBytes,
} from '../xlsx-large-simple-metadata-byte-scan.js'

const encoder = new TextEncoder()

describe('large simple metadata byte scan', () => {
  it('parses column metadata without decoding the cols XML span', () => {
    const bytes = encoder.encode(
      '<cols><col min="1" max="2" width="12.5" style="3" hidden="1" customWidth="1" bestFit="0" outlineLevel="1"/></cols>',
    )
    const entries: Parameters<typeof appendLargeSimpleColumnMetadataFromBytes>[0] = []
    const metadata: Parameters<typeof appendLargeSimpleColumnMetadataFromBytes>[1] = []

    appendLargeSimpleColumnMetadataFromBytes(entries, metadata, bytes, 0, bytes.byteLength)

    expect(entries).toEqual([
      { id: 'col:0', index: 0, size: 75, hidden: true },
      { id: 'col:1', index: 1, size: 75, hidden: true },
    ])
    expect(metadata).toEqual([
      {
        start: 0,
        count: 2,
        size: 75,
        xlsxWidth: 12.5,
        styleIndex: 3,
        hidden: true,
        customWidth: true,
        bestFit: false,
        outlineLevel: 1,
      },
    ])
  })

  it('coalesces contiguous repeated row and column metadata while scanning bytes', () => {
    const rowMetadata: Parameters<typeof appendLargeSimpleRowMetadataTagFromBytes>[1] = []
    const rowEntries: Parameters<typeof appendLargeSimpleRowMetadataTagFromBytes>[0] = []
    for (const row of [
      '<row r="1" s="2" customFormat="1"/>',
      '<row r="2" s="2" customFormat="1"/>',
      '<row r="3" s="3" customFormat="1"/>',
    ]) {
      const bytes = encoder.encode(row)
      appendLargeSimpleRowMetadataTagFromBytes(rowEntries, rowMetadata, bytes, '<row'.length, bytes.byteLength - 1)
    }

    expect(rowEntries).toEqual([])
    expect(rowMetadata).toEqual([
      { start: 0, count: 2, styleIndex: 2, customFormat: true },
      { start: 2, count: 1, styleIndex: 3, customFormat: true },
    ])

    const columnBytes = encoder.encode(
      '<cols><col min="1" max="1" style="4" customFormat="1"/><col min="2" max="3" style="4" customFormat="1"/></cols>',
    )
    const columnEntries: Parameters<typeof appendLargeSimpleColumnMetadataFromBytes>[0] = []
    const columnMetadata: Parameters<typeof appendLargeSimpleColumnMetadataFromBytes>[1] = []

    appendLargeSimpleColumnMetadataFromBytes(columnEntries, columnMetadata, columnBytes, 0, columnBytes.byteLength)

    expect(columnEntries).toEqual([])
    expect(columnMetadata).toEqual([{ start: 0, count: 3, styleIndex: 4, customFormat: true }])
  })

  it('parses sheet format, drawing, and table relationship metadata from bytes', () => {
    const sheetFormatPr = encoder.encode('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>')
    const drawing = encoder.encode('<drawing r:id="rIdDrawing1"/>')
    const tableParts = encoder.encode('<tableParts><tablePart r:id="rIdTable1"/></tableParts>')

    expect(readLargeSimpleSheetFormatPrTagFromBytes(sheetFormatPr, 0, sheetFormatPr.byteLength)).toEqual({
      defaultRowHeight: 15,
      outlineLevelRow: 1,
    })
    expect(readLargeSimpleDrawingRelationshipIdTagFromBytes(drawing, 0, drawing.byteLength)).toBe('rIdDrawing1')
    expect(readLargeSimpleTableRelationshipIdsFromBytes(tableParts, 0, tableParts.byteLength)).toEqual(['rIdTable1'])
  })
})
