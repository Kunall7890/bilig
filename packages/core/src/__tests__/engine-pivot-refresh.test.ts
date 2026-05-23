import { describe, expect, it } from 'vitest'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('engine pivot source refresh', () => {
  it('uses live worksheet source rows instead of stale imported pivot cache records', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'imported-pivot-cache' })
    await engine.ready()
    engine.importSnapshot(importedSourceBackedPivotSnapshot())
    engine.setCellFormula('Summary', 'A1', 'GETPIVOTDATA("Sales",Pivot!B2,"Region","East")')

    expect(engine.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 15 })

    engine.setCellValue('Data', 'D2', 20)

    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 25 })
    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 25 })
    expect(engine.exportSnapshot().workbook.metadata?.pivots?.[0]).toMatchObject({
      cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
      cachedRecords: [
        ['East', 'priority', 'Widget', 20],
        ['West', 'priority', 'Widget', 7],
        ['East', 'priority', 'Gizmo', 5],
      ],
      rows: 3,
      cols: 2,
    })
  })

  it('evaluates GETPIVOTDATA with the same aggregate variants as materialized pivots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'pivot-aggregate-getpivotdata' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C4' }, [
      ['Region', 'Sales', 'Units'],
      ['East', 10, 2],
      ['East', 20, 3],
      ['West', 7, 5],
    ])
    engine.setPivotTable('Pivot', 'B2', {
      name: 'AggregatePivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'C4' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'average', outputLabel: 'Average Sales' },
        { sourceColumn: 'Sales', summarizeBy: 'min', outputLabel: 'Min Sales' },
        { sourceColumn: 'Sales', summarizeBy: 'max', outputLabel: 'Max Sales' },
        { sourceColumn: 'Units', summarizeBy: 'countNums', outputLabel: 'Numeric Units' },
        { sourceColumn: 'Units', summarizeBy: 'product', outputLabel: 'Product Units' },
      ],
    })

    engine.setCellFormula('Summary', 'A1', 'GETPIVOTDATA("Average Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Summary', 'A2', 'GETPIVOTDATA("Min Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Summary', 'A3', 'GETPIVOTDATA("Max Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Summary', 'A4', 'GETPIVOTDATA("Numeric Units",Pivot!B2,"Region","East")')
    engine.setCellFormula('Summary', 'A5', 'GETPIVOTDATA("Product Units",Pivot!B2,"Region","East")')

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Summary', 'A4')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Summary', 'A5')).toEqual({ tag: ValueTag.Number, value: 6 })
  })
})

function importedSourceBackedPivotSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Imported stale pivot cache',
      metadata: {
        pivots: [
          {
            name: 'SalesByRegion',
            sheetName: 'Pivot',
            address: 'B2',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
            sourceKind: 'worksheet',
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
            cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
            cachedRecords: [
              ['East', 'priority', 'Widget', 1],
              ['West', 'priority', 'Widget', 2],
            ],
            rows: 1,
            cols: 1,
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
          { address: 'B1', value: 'Notes' },
          { address: 'C1', value: 'Product' },
          { address: 'D1', value: 'Sales' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 'priority' },
          { address: 'C2', value: 'Widget' },
          { address: 'D2', value: 10 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 'priority' },
          { address: 'C3', value: 'Widget' },
          { address: 'D3', value: 7 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 'priority' },
          { address: 'C4', value: 'Gizmo' },
          { address: 'D4', value: 5 },
        ],
      },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        cells: [],
      },
      {
        id: 3,
        name: 'Summary',
        order: 2,
        cells: [],
      },
    ],
  }
}
