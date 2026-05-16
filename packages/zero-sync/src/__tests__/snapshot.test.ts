import { describe, expect, it } from 'vitest'
import { projectWorkbookToSnapshot } from '../snapshot'

describe('projectWorkbookToSnapshot', () => {
  it('reconstructs workbook metadata and sheet formatting from normalized Zero rows', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        compatibilityMode: 'excel-modern',
        recalcEpoch: 7,
        snapshot: {
          version: 1,
          workbook: {
            name: 'Warm Snapshot',
            metadata: {
              tables: [
                {
                  name: 'FallbackTable',
                  sheetName: 'Sheet1',
                  startAddress: 'A1',
                  endAddress: 'A2',
                  columnNames: ['A'],
                  headerRow: true,
                  totalsRow: false,
                },
              ],
            },
          },
          sheets: [
            {
              name: 'Sheet1',
              order: 0,
              cells: [],
            },
          ],
        },
        calculationSettings: {
          mode: 'automatic',
          recalcEpoch: 7,
        },
        workbookMetadataEntries: [{ key: 'locale', value: 'en-US' }],
        definedNames: [
          {
            name: 'Sales',
            value: {
              kind: 'range-ref',
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'A2',
            },
          },
        ],
        styles: [
          {
            id: 'style-1',
            recordJSON: {
              fill: {
                backgroundColor: '#ffee00',
              },
            },
          },
        ],
        numberFormats: [
          {
            id: 'fmt-1',
            code: '0.00',
            kind: 'number',
          },
        ],
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
            freezeRows: 1,
            freezeCols: 2,
            cells: [
              {
                address: 'A1',
                inputValue: 42,
                explicitFormatId: 'fmt-1',
              },
            ],
            rowMetadata: [{ startIndex: 0, count: 1, size: 28 }],
            columnMetadata: [{ startIndex: 0, count: 1, size: 144 }],
            styleRanges: [
              {
                startRow: 0,
                endRow: 0,
                startCol: 0,
                endCol: 0,
                styleId: 'style-1',
              },
            ],
            formatRanges: [
              {
                startRow: 0,
                endRow: 0,
                startCol: 0,
                endCol: 0,
                formatId: 'fmt-1',
              },
            ],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.workbook.name).toBe('Projected Book')
    expect(projected?.workbook.metadata?.properties).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(projected?.workbook.metadata?.definedNames).toEqual([
      {
        name: 'Sales',
        value: {
          kind: 'range-ref',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A2',
        },
      },
    ])
    expect(projected?.workbook.metadata?.styles).toEqual([
      {
        id: 'style-1',
        fill: {
          backgroundColor: '#ffee00',
        },
      },
    ])
    expect(projected?.workbook.metadata?.formats).toEqual([
      {
        id: 'fmt-1',
        code: '0.00',
        kind: 'number',
      },
    ])
    expect(projected?.workbook.metadata?.tables?.[0]?.name).toBe('FallbackTable')
    expect(projected?.sheets[0]?.cells[0]).toEqual({
      address: 'A1',
      value: 42,
      format: '0.00',
    })
    expect(projected?.sheets[0]?.metadata?.freezePane).toEqual({ rows: 1, cols: 2 })
    expect(projected?.sheets[0]?.metadata?.styleRanges).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        styleId: 'style-1',
      },
    ])
    expect(projected?.sheets[0]?.metadata?.formatRanges).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
        formatId: 'fmt-1',
      },
    ])
  })

  it('preserves fallback axis style metadata when normalized rows omit it', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        snapshot: {
          version: 1,
          workbook: { name: 'Warm Snapshot' },
          sheets: [
            {
              name: 'Sheet1',
              order: 0,
              cells: [],
              metadata: {
                rowMetadata: [{ start: 0, count: 1, size: 28, styleIndex: 7, customFormat: true }],
                columnMetadata: [{ start: 0, count: 1, size: 144, styleIndex: 9, customFormat: true }],
              },
            },
          ],
        },
        calculationSettings: {
          mode: 'automatic',
          recalcEpoch: 0,
        },
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
            freezeRows: 0,
            freezeCols: 0,
            cells: [],
            rowMetadata: [{ startIndex: 0, count: 1, size: 28 }],
            columnMetadata: [{ startIndex: 0, count: 1, size: 144 }],
            styleRanges: [],
            formatRanges: [],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.sheets[0]?.metadata?.rowMetadata).toEqual([{ start: 0, count: 1, size: 28, styleIndex: 7, customFormat: true }])
    expect(projected?.sheets[0]?.metadata?.columnMetadata).toEqual([{ start: 0, count: 1, size: 144, styleIndex: 9, customFormat: true }])
  })

  it('drops unsafe projected sheet metadata bounds', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        sheets: [
          {
            name: 'Sheet1',
            sortOrder: 0,
            freezeRows: 1.5,
            freezeCols: Number.MAX_SAFE_INTEGER + 1,
            cells: [],
            rowMetadata: [
              { startIndex: 0, count: 1, size: 28 },
              { startIndex: -1, count: 1, size: 30 },
              { startIndex: 2, count: 0, size: 30 },
            ],
            columnMetadata: [{ startIndex: 0, count: Number.MAX_SAFE_INTEGER + 1, size: 144 }],
            styleRanges: [
              { startRow: 2, endRow: 1, startCol: 0, endCol: 0, styleId: 'style-1' },
              { startRow: 0, endRow: 0, startCol: 1.5, endCol: 2, styleId: 'style-1' },
            ],
            formatRanges: [{ startRow: 0, endRow: 0, startCol: 2, endCol: 1, formatId: 'fmt-1' }],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.sheets[0]?.metadata?.freezePane).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.rowMetadata).toEqual([{ start: 0, count: 1, size: 28 }])
    expect(projected?.sheets[0]?.metadata?.columnMetadata).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.styleRanges).toBeUndefined()
    expect(projected?.sheets[0]?.metadata?.formatRanges).toBeUndefined()
  })

  it('drops malformed defined-name values', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        definedNames: [
          { name: 'Rate', value: 0.12 },
          { name: 'Scalar', value: { kind: 'scalar', value: 'ok' } },
          { name: 'Cell', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' } },
          { name: 'BadNumber', value: Number.NaN },
          { name: 'BadScalar', value: { kind: 'scalar', value: Number.POSITIVE_INFINITY } },
          { name: 'BadObject', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1' } },
          { name: 'ArbitraryObject', value: { formula: 'A1' } },
        ],
        sheets: [],
      },
      'doc-1',
    )

    expect(projected?.workbook.metadata?.definedNames).toEqual([
      { name: 'Rate', value: 0.12 },
      { name: 'Scalar', value: { kind: 'scalar', value: 'ok' } },
      { name: 'Cell', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' } },
    ])
  })

  it('drops unsafe projected sheet and cell coordinates', () => {
    const projected = projectWorkbookToSnapshot(
      {
        id: 'doc-1',
        name: 'Projected Book',
        recalcEpoch: 2.5,
        sheets: [
          {
            id: Number.MAX_SAFE_INTEGER + 1,
            name: 'BadOrder',
            sortOrder: 1.5,
            cells: [{ rowNum: 0, colNum: 0, inputValue: 'ignored' }],
          },
          {
            id: 2,
            name: 'Sheet1',
            sortOrder: 0,
            cells: [
              { rowNum: 1, colNum: 2, inputValue: 'kept' },
              { rowNum: -1, colNum: 0, inputValue: 'dropped' },
              { rowNum: 0.5, colNum: 0, inputValue: 'dropped' },
            ],
            rowMetadata: [{ startIndex: 0, count: 1, size: -10 }],
          },
        ],
      },
      'doc-1',
    )

    expect(projected?.workbook.metadata?.volatileContext).toBeUndefined()
    expect(projected?.sheets).toEqual([
      {
        id: 2,
        name: 'Sheet1',
        order: 0,
        cells: [{ address: 'C2', value: 'kept' }],
        metadata: { rowMetadata: [{ start: 0, count: 1 }] },
      },
    ])
  })
})
