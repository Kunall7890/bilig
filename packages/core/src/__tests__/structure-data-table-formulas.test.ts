import { ValueTag, type WorkbookSheetDataTableFormulasSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { SpreadsheetEngine } from '../index.js'

describe('structural data-table formula metadata', () => {
  it('preserves native data-table metadata through engine import/export', async () => {
    const engine = await buildNativeDataTableEngine()

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toEqual(twoVariableDataTableMetadata())
  })

  it('retargets native data-table metadata when inserting rows above the table', async () => {
    const engine = await buildNativeDataTableEngine()

    engine.insertRows('DataTable', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [{ address: 'C4', formulaXml: '<f t="dataTable" ref="C4:D5" dt2D="1" dtr="1" r1="A2" r2="A3"/>' }],
    })
    expect(engine.getCell('DataTable', 'C4').formula).toBe('MULTIPLE.OPERATIONS(B3,A2,C3,A3,B4)')
    expect(engine.getCell('DataTable', 'D5').formula).toBe('MULTIPLE.OPERATIONS(B3,A2,D3,A3,B5)')
    expect(engine.getCellValue('DataTable', 'C4')).toEqual({ tag: ValueTag.Number, value: 40 })
    expect(engine.getCellValue('DataTable', 'D5')).toEqual({ tag: ValueTag.Number, value: 90 })
  })

  it('retargets native data-table metadata when inserting columns left of the table', async () => {
    const engine = await buildNativeDataTableEngine()

    engine.insertColumns('DataTable', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [{ address: 'D3', formulaXml: '<f t="dataTable" ref="D3:E4" dt2D="1" dtr="1" r1="B1" r2="B2"/>' }],
    })
    expect(engine.getCell('DataTable', 'D3').formula).toBe('MULTIPLE.OPERATIONS(C2,B1,D2,B2,C3)')
    expect(engine.getCell('DataTable', 'E4').formula).toBe('MULTIPLE.OPERATIONS(C2,B1,E2,B2,C4)')
    expect(engine.getCellValue('DataTable', 'D3')).toEqual({ tag: ValueTag.Number, value: 40 })
    expect(engine.getCellValue('DataTable', 'E4')).toEqual({ tag: ValueTag.Number, value: 90 })
  })

  it('retargets native one-variable data-table metadata when inserting rows above the tables', async () => {
    const engine = await buildNativeDataTableEngine(oneVariableDataTableSnapshot())

    engine.insertRows('DataTable', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [
        { address: 'C3', formulaXml: '<f t="dataTable" ref="C3:D3" dt2D="0" dtr="1" r1="A2" ca="1"/>' },
        { address: 'B7', formulaXml: '<f t="dataTable" ref="B7:B9" dt2D="0" dtr="0" r1="A2"/>' },
      ],
    })
    expect(engine.getCell('DataTable', 'C3').formula).toBe('MULTIPLE.OPERATIONS(B3,A2,C2)')
    expect(engine.getCell('DataTable', 'B9').formula).toBe('MULTIPLE.OPERATIONS(B6,A2,A9)')
    expect(engine.getCellValue('DataTable', 'C3')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('DataTable', 'B9')).toEqual({ tag: ValueTag.Number, value: 40 })
  })

  it('retargets native one-variable data-table metadata when inserting columns left of the tables', async () => {
    const engine = await buildNativeDataTableEngine(oneVariableDataTableSnapshot())

    engine.insertColumns('DataTable', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toEqual({
      formulas: [
        { address: 'D2', formulaXml: '<f t="dataTable" ref="D2:E2" dt2D="0" dtr="1" r1="B1" ca="1"/>' },
        { address: 'C6', formulaXml: '<f t="dataTable" ref="C6:C8" dt2D="0" dtr="0" r1="B1"/>' },
      ],
    })
    expect(engine.getCell('DataTable', 'D2').formula).toBe('MULTIPLE.OPERATIONS(C2,B1,D1)')
    expect(engine.getCell('DataTable', 'C8').formula).toBe('MULTIPLE.OPERATIONS(C5,B1,B8)')
    expect(engine.getCellValue('DataTable', 'D2')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('DataTable', 'C8')).toEqual({ tag: ValueTag.Number, value: 40 })
  })

  it('drops native data-table metadata when a required address is deleted', async () => {
    const engine = await buildNativeDataTableEngine()

    engine.deleteRows('DataTable', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toBeUndefined()
  })

  it('drops native data-table metadata when inserting through the data-table rectangle', async () => {
    const engine = await buildNativeDataTableEngine()

    engine.insertRows('DataTable', 2, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toBeUndefined()
  })

  it('preserves native data-table metadata for no-op inserts after the table and input refs', async () => {
    const engine = await buildNativeDataTableEngine()

    engine.insertRows('DataTable', 8, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toEqual(twoVariableDataTableMetadata())
  })

  it('drops native data-table metadata for unproved moves through the table geometry', async () => {
    const engine = await buildNativeDataTableEngine()

    engine.moveRows('DataTable', 0, 1, 6)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toBeUndefined()
  })

  it('drops native data-table metadata when the anchor diverges from the rewritten ref top-left', async () => {
    const snapshot = twoVariableDataTableSnapshot()
    const sheet = snapshot.sheets[0]
    if (!sheet) {
      throw new Error('Missing test sheet')
    }
    const engine = await buildNativeDataTableEngine({
      ...snapshot,
      sheets: [
        {
          ...sheet,
          metadata: {
            dataTableFormulas: {
              formulas: [{ address: 'C3', formulaXml: '<f t="dataTable" ref="D3:D4" dt2D="0" dtr="0" r1="A1"/>' }],
            },
          },
        },
      ],
    })

    engine.insertRows('DataTable', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.dataTableFormulas).toBeUndefined()
  })
})

async function buildNativeDataTableEngine(snapshot: WorkbookSnapshot = twoVariableDataTableSnapshot()): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'native-data-table-structural' })
  await engine.ready()
  engine.importSnapshot(snapshot)
  return engine
}

function twoVariableDataTableMetadata(): WorkbookSheetDataTableFormulasSnapshot {
  return {
    formulas: [{ address: 'C3', formulaXml: '<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2"/>' }],
  }
}

function twoVariableDataTableSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'native-data-table-structural' },
    sheets: [
      {
        id: 1,
        name: 'DataTable',
        order: 0,
        metadata: { dataTableFormulas: twoVariableDataTableMetadata() },
        cells: [
          { address: 'A1', value: 1 },
          { address: 'A2', value: 10 },
          { address: 'B2', formula: 'A3' },
          { address: 'C2', value: 2 },
          { address: 'D2', value: 3 },
          { address: 'A3', formula: 'A1*A2' },
          { address: 'B3', value: 20 },
          { address: 'C3', formula: 'MULTIPLE.OPERATIONS(B2,A1,C2,A2,B3)' },
          { address: 'D3', formula: 'MULTIPLE.OPERATIONS(B2,A1,D2,A2,B3)' },
          { address: 'B4', value: 30 },
          { address: 'C4', formula: 'MULTIPLE.OPERATIONS(B2,A1,C2,A2,B4)' },
          { address: 'D4', formula: 'MULTIPLE.OPERATIONS(B2,A1,D2,A2,B4)' },
        ],
      },
    ],
  }
}

function oneVariableDataTableMetadata(): WorkbookSheetDataTableFormulasSnapshot {
  return {
    formulas: [
      { address: 'C2', formulaXml: '<f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1" ca="1"/>' },
      { address: 'B6', formulaXml: '<f t="dataTable" ref="B6:B8" dt2D="0" dtr="0" r1="A1"/>' },
    ],
  }
}

function oneVariableDataTableSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'native-one-variable-data-table-structural' },
    sheets: [
      {
        id: 1,
        name: 'DataTable',
        order: 0,
        metadata: { dataTableFormulas: oneVariableDataTableMetadata() },
        cells: [
          { address: 'A1', value: 1 },
          { address: 'A2', formula: 'A1*10' },
          { address: 'B1', value: 2 },
          { address: 'C1', value: 3 },
          { address: 'D1', value: 4 },
          { address: 'B2', formula: 'A2' },
          { address: 'C2', formula: 'MULTIPLE.OPERATIONS(B2,A1,C1)' },
          { address: 'D2', formula: 'MULTIPLE.OPERATIONS(B2,A1,D1)' },
          { address: 'A5', value: 1 },
          { address: 'B5', formula: 'A1*10' },
          { address: 'A6', value: 2 },
          { address: 'B6', formula: 'MULTIPLE.OPERATIONS(B5,A1,A6)' },
          { address: 'A7', value: 3 },
          { address: 'B7', formula: 'MULTIPLE.OPERATIONS(B5,A1,A7)' },
          { address: 'A8', value: 4 },
          { address: 'B8', formula: 'MULTIPLE.OPERATIONS(B5,A1,A8)' },
        ],
      },
    ],
  }
}
