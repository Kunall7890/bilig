import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { EnginePivotError } from '../engine/errors.js'
import type { EnginePivotService } from '../engine/services/pivot-service.js'

function isEnginePivotService(value: unknown): value is EnginePivotService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'materializePivot') === 'function' &&
    typeof Reflect.get(value, 'resolvePivotData') === 'function' &&
    typeof Reflect.get(value, 'clearOwnedPivot') === 'function' &&
    typeof Reflect.get(value, 'clearPivotForCell') === 'function'
  )
}

function getPivotService(engine: SpreadsheetEngine): EnginePivotService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null || !('pivot' in runtime)) {
    throw new TypeError('Expected engine runtime to expose a pivot service')
  }
  const pivot = Reflect.get(runtime, 'pivot')
  if (!isEnginePivotService(pivot)) {
    throw new TypeError('Expected engine runtime pivot service')
  }
  return pivot
}

async function buildPivotEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'spec' })
  await engine.ready()
  engine.createSheet('Data')
  engine.createSheet('Pivot')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' }, [
    ['Region', 'Notes', 'Product', 'Sales'],
    ['East', 'priority', 'Widget', 10],
    ['West', 'priority', 'Widget', 7],
    ['East', 'priority', 'Gizmo', 5],
  ])
  engine.setPivotTable('Pivot', 'B2', {
    name: 'SalesByRegion',
    source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
    groupBy: ['Region'],
    values: [
      { sourceColumn: 'Sales', summarizeBy: 'sum' },
      { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
    ],
  })
  return engine
}

describe('EnginePivotService', () => {
  it('clears owned pivot output cells without leaving stale ownership behind', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivot = engine.getPivotTables()[0]
    if (!pivot) {
      throw new TypeError('Expected pivot table')
    }

    const changed = Effect.runSync(service.clearOwnedPivot(pivot))

    expect(changed.length).toBeGreaterThan(0)
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Empty })
  })

  it('resolves pivot aggregates through the extracted service boundary', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)

    const resolved = Effect.runSync(
      service.resolvePivotData('Pivot', 'C3', 'Sales', [{ field: 'Region', item: engine.getCellValue('Pivot', 'B3') }]),
    )

    expect(resolved).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('preserves full semantic pivot fields when source edits resize materialized output', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'pivot-semantic-preservation' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'E5' }, [
      ['Region', 'Quarter', 'Status', 'Segment', 'Sales'],
      ['East', 'Q1', 'Closed', 'Enterprise', 10],
      ['East', 'Q2', 'Closed', 'Enterprise', 5],
      ['West', 'Q2', 'Closed', 'SMB', 7],
      ['West', 'Q3', 'Open', 'Enterprise', 100],
    ])

    engine.setPivotTable('Pivot', 'B2', {
      name: 'ClosedEnterpriseSalesByQuarter',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'E6' },
      groupBy: ['Region'],
      columnFields: ['Quarter'],
      pageFields: [{ sourceColumn: 'Segment', selectedValue: 'Enterprise' }],
      filters: [{ sourceColumn: 'Status', includedValues: ['Closed'] }],
      hiddenItems: [{ sourceColumn: 'Quarter', values: ['Q2'] }],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Closed Sales' }],
    })

    expect(engine.getCellValue('Pivot', 'C2')).toMatchObject({ tag: ValueTag.String, value: 'Q1 Closed Sales' })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A6', endAddress: 'E6' }, [['East', 'Q3', 'Closed', 'Enterprise', 9]])

    expect(engine.getCellValue('Pivot', 'D2')).toMatchObject({ tag: ValueTag.String, value: 'Q3 Closed Sales' })
    expect(engine.getCellValue('Pivot', 'D3')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getPivotTables()[0]).toMatchObject({
      name: 'ClosedEnterpriseSalesByQuarter',
      columnFields: ['Quarter'],
      pageFields: [{ sourceColumn: 'Segment', selectedValue: 'Enterprise' }],
      filters: [{ sourceColumn: 'Status', includedValues: ['Closed'] }],
      hiddenItems: [{ sourceColumn: 'Quarter', values: ['Q2'] }],
      rows: 2,
      cols: 3,
    })
  })

  it('resolves GETPIVOTDATA aggregate variants through the materialized pivot reducer', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'pivot-getpivotdata-aggregates' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.createSheet('Report')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C4' }, [
      ['Region', 'Sales', 'Units'],
      ['East', 10, 2],
      ['East', 20, 3],
      ['West', 7, 5],
    ])
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesStats',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'C4' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'average', outputLabel: 'Avg Sales' },
        { sourceColumn: 'Sales', summarizeBy: 'min', outputLabel: 'Min Sales' },
        { sourceColumn: 'Sales', summarizeBy: 'max', outputLabel: 'Max Sales' },
        { sourceColumn: 'Sales', summarizeBy: 'countNums', outputLabel: 'Numeric Sales' },
        { sourceColumn: 'Units', summarizeBy: 'product', outputLabel: 'Unit Product' },
      ],
    })

    engine.setCellFormula('Report', 'A1', 'GETPIVOTDATA("Avg Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Report', 'A2', 'GETPIVOTDATA("Min Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Report', 'A3', 'GETPIVOTDATA("Max Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Report', 'A4', 'GETPIVOTDATA("Numeric Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Report', 'A5', 'GETPIVOTDATA("Unit Product",Pivot!B2,"Region","East")')

    expect(engine.getCellValue('Report', 'A1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Report', 'A2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Report', 'A3')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Report', 'A4')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Report', 'A5')).toEqual({ tag: ValueTag.Number, value: 6 })
  })

  it('uses worksheet source instead of stale imported cache records for source-backed pivots', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'stale-source-backed-pivot-cache',
        metadata: {
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Pivot',
              address: 'B2',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
              sourceKind: 'worksheet',
              cacheFields: ['Region', 'Sales'],
              cachedRecords: [
                ['East', 999],
                ['West', 7],
                ['East', 1],
              ],
              groupBy: ['Region'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
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
            { address: 'B1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 7 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 5 },
          ],
        },
        { id: 2, name: 'Pivot', order: 1, cells: [] },
        { id: 3, name: 'Report', order: 2, cells: [] },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'stale-source-backed-pivot-cache' })
    await engine.ready()
    engine.importSnapshot(snapshot)
    engine.setCellFormula('Report', 'A1', 'GETPIVOTDATA("Sales Total",Pivot!B2,"Region","East")')

    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Report', 'A1')).toEqual({ tag: ValueTag.Number, value: 15 })

    engine.setCellValue('Data', 'B2', 100)

    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(engine.getCellValue('Report', 'A1')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(engine.exportSnapshot().workbook.metadata?.pivots?.[0]).toMatchObject({ rows: 3, cols: 2 })
  })

  it('claims imported pivot output cells before refreshing source-backed pivots', async () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'excel-authored-source-backed-pivot-output',
        metadata: {
          styles: [{ id: 'xlsx-style-1', font: { bold: true } }],
          pivots: [
            {
              name: 'SalesByRegionQuarter',
              sheetName: 'Pivot',
              address: 'B2',
              source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'C5' },
              sourceKind: 'worksheet',
              cacheFields: ['Region', 'Quarter', 'Sales'],
              cachedRecords: [
                ['East', 'Q1', 10],
                ['East', 'Q2', 5],
                ['West', 'Q2', 7],
                ['East', 'Q3', 9],
              ],
              groupBy: ['Region'],
              columnFields: ['Quarter'],
              values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
              rows: 4,
              cols: 4,
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
            { address: 'B1', value: 'Quarter' },
            { address: 'C1', value: 'Sales' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 'Q1' },
            { address: 'C2', value: 10 },
            { address: 'A3', value: 'East' },
            { address: 'B3', value: 'Q2' },
            { address: 'C3', value: 5 },
            { address: 'A4', value: 'West' },
            { address: 'B4', value: 'Q2' },
            { address: 'C4', value: 7 },
            { address: 'A5', value: 'East' },
            { address: 'B5', value: 'Q3' },
            { address: 'C5', value: 9 },
          ],
        },
        {
          id: 2,
          name: 'Pivot',
          order: 1,
          metadata: {
            styleRanges: [{ range: { sheetName: 'Pivot', startAddress: 'B4', endAddress: 'B5' }, styleId: 'xlsx-style-1' }],
          },
          cells: [
            { address: 'B2', value: 'Region' },
            { address: 'C2', value: 'Q1' },
            { address: 'D2', value: 'Q2' },
            { address: 'E2', value: 'Q3' },
            { address: 'B3', value: 'East' },
            { address: 'C3', value: 10 },
            { address: 'D3', value: 5 },
            { address: 'E3', value: 9 },
            { address: 'B4', value: 'West' },
            { address: 'D4', value: 7 },
            { address: 'B5', value: 'Grand Total' },
            { address: 'C5', value: 10 },
            { address: 'D5', value: 12 },
            { address: 'E5', value: 9 },
          ],
        },
      ],
    }

    const engine = new SpreadsheetEngine({ workbookName: 'excel-authored-source-backed-pivot-output' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    expect(engine.getCellValue('Pivot', 'B2').tag).not.toBe(ValueTag.Error)
    expect(engine.getCellValue('Pivot', 'B5')).toEqual({ tag: ValueTag.Empty })
    expect(engine.exportSnapshot().sheets.find((sheet) => sheet.name === 'Pivot')?.metadata?.styleRanges).toBeUndefined()
    expect(engine.exportSnapshot().workbook.metadata?.pivots?.[0]).toMatchObject({ rows: 3, cols: 4 })

    engine.setCellFormula('Pivot', 'H1', 'GETPIVOTDATA("Sales Total",B2,"Region","East","Quarter","Q1")')
    engine.setCellFormula('Pivot', 'H2', 'GETPIVOTDATA("Sales Total",B2,"Region","East","Quarter","Q3")')
    engine.setCellValue('Data', 'C2', 100)

    expect(engine.getCellValue('Pivot', 'H1')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Pivot', 'H2')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.exportSnapshot().workbook.metadata?.pivots?.[0]).toMatchObject({ rows: 3, cols: 4 })
  })

  it('clears pivot ownership for one cell and rematerializes through the service boundary', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivot = engine.getPivotTables()[0]
    if (!pivot) {
      throw new TypeError('Expected pivot table')
    }
    const pivotCellIndex = engine.workbook.getCellIndex('Pivot', 'B2')
    if (pivotCellIndex === undefined) {
      throw new TypeError('Expected pivot output cell')
    }

    const cleared = Effect.runSync(service.clearPivotForCell(pivotCellIndex))
    expect(cleared.length).toBeGreaterThan(0)
    expect(engine.getPivotTables()).toEqual([])
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })

    const rematerialized = Effect.runSync(service.materializePivot(pivot))
    expect(rematerialized.length).toBeGreaterThan(0)
    expect(engine.getCellValue('Pivot', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'Region',
    })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('drops orphaned pivot ownership when the pivot metadata is already gone', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivotCellIndex = engine.workbook.getCellIndex('Pivot', 'B2')
    if (pivotCellIndex === undefined) {
      throw new TypeError('Expected pivot output cell')
    }

    expect(engine.deletePivotTable('Pivot', 'B2')).toBe(true)
    const pivotOutputOwners = Reflect.get(engine, 'pivotOutputOwners')
    if (!(pivotOutputOwners instanceof Map)) {
      throw new TypeError('Expected pivot output owners')
    }
    pivotOutputOwners.set(pivotCellIndex, 'Pivot!B2')

    expect(Effect.runSync(service.clearPivotForCell(pivotCellIndex))).toEqual([])
    expect(pivotOutputOwners.has(pivotCellIndex)).toBe(false)
  })

  it('wraps materialize, resolve, clear-owned, and clear-for-cell failures with EnginePivotError', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivot = engine.getPivotTables()[0]
    if (!pivot) {
      throw new TypeError('Expected pivot table')
    }

    const getValueSpy = vi.spyOn(engine.workbook.cellStore, 'getValue').mockImplementation(() => {
      throw new Error('pivot cell explode')
    })
    const materialize = Effect.runSync(Effect.either(service.materializePivot(pivot)))
    expect(materialize._tag).toBe('Left')
    expect(materialize.left).toBeInstanceOf(EnginePivotError)
    expect(materialize.left.message).toContain(`Failed to materialize pivot ${pivot.name}`)
    const clearedOwned = Effect.runSync(Effect.either(service.clearOwnedPivot(pivot)))
    expect(clearedOwned._tag).toBe('Left')
    expect(clearedOwned.left).toBeInstanceOf(EnginePivotError)
    expect(clearedOwned.left.message).toContain(`Failed to clear pivot output ownership for ${pivot.name}`)
    getValueSpy.mockRestore()

    const listPivotsSpy = vi.spyOn(engine.workbook, 'listPivots').mockImplementation(() => {
      throw new Error('resolve explode')
    })
    const resolved = Effect.runSync(Effect.either(service.resolvePivotData('Pivot', 'B2', 'Sales', [])))
    expect(resolved._tag).toBe('Left')
    expect(resolved.left).toBeInstanceOf(EnginePivotError)
    expect(resolved.left.message).toContain('Failed to resolve pivot data for Pivot!B2')
    listPivotsSpy.mockRestore()

    const pivotCellIndex = engine.workbook.getCellIndex('Pivot', 'B2')
    if (pivotCellIndex === undefined) {
      throw new TypeError('Expected pivot output cell')
    }
    const getPivotByKeySpy = vi.spyOn(engine.workbook, 'getPivotByKey').mockImplementation(() => {
      throw new Error('clear cell explode')
    })
    const clearedForCell = Effect.runSync(Effect.either(service.clearPivotForCell(pivotCellIndex)))
    expect(clearedForCell._tag).toBe('Left')
    expect(clearedForCell.left).toBeInstanceOf(EnginePivotError)
    expect(clearedForCell.left.message).toContain(`Failed to clear pivot ownership for cell ${pivotCellIndex}`)
    getPivotByKeySpy.mockRestore()
  })
})
