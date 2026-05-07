import { describe, expect, it } from 'vitest'
import { formatAddress } from '@bilig/formula'
import { ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

const sheetName = 'Sheet1'

type SheetCellInput = LiteralInput

// Local HyperFormula smoke tests used as the behavior source at revision
// 9a510a2acb97c3d3490f9e3b9e961a1c4a98b9ad.
// The assertions below are bilig-owned parity scenarios, not copied test bodies.
describe('HyperFormula smoke behavior parity', () => {
  it('builds an array-backed sheet and evaluates range sums', async () => {
    const engine = await createEngineFromGrid([
      [1, 2, 3],
      [4, 5, 6],
      ['=SUM(A1:C1)', '=SUM(A2:C2)', '=SUM(A1:C2)'],
    ])

    expectNumber(engine.getCellValue(sheetName, 'A3'), 6)
    expectNumber(engine.getCellValue(sheetName, 'B3'), 15)
    expectNumber(engine.getCellValue(sheetName, 'C3'), 21)
  })

  it('evaluates arithmetic and logical formulas', async () => {
    const engine = await createEngineFromGrid([
      [10, 20, 30],
      ['=A1+B1+C1', '=A1*B1', '=C1/A1'],
      ['=IF(A1>5, "big", "small")', '=AND(A1>0, B1>0)', '=OR(A1<0, B1>0)'],
    ])

    expectNumber(engine.getCellValue(sheetName, 'A2'), 60)
    expectNumber(engine.getCellValue(sheetName, 'B2'), 200)
    expectNumber(engine.getCellValue(sheetName, 'C2'), 3)
    expectString(engine.getCellValue(sheetName, 'A3'), 'big')
    expectBoolean(engine.getCellValue(sheetName, 'B3'), true)
    expectBoolean(engine.getCellValue(sheetName, 'C3'), true)
  })

  it('evaluates common spreadsheet functions', async () => {
    const engine = await createEngineFromGrid([
      [1, 2, 3, 4, 5],
      ['=SUM(A1:E1)', '=AVERAGE(A1:E1)', '=MIN(A1:E1)', '=MAX(A1:E1)', '=COUNT(A1:E1)'],
      ['=CONCATENATE("Hello", " ", "World")', '=LEN("Test")', '=UPPER("hello")', '=LOWER("HELLO")', '=ABS(-5)'],
    ])

    expectNumber(engine.getCellValue(sheetName, 'A2'), 15)
    expectNumber(engine.getCellValue(sheetName, 'B2'), 3)
    expectNumber(engine.getCellValue(sheetName, 'C2'), 1)
    expectNumber(engine.getCellValue(sheetName, 'D2'), 5)
    expectNumber(engine.getCellValue(sheetName, 'E2'), 5)
    expectString(engine.getCellValue(sheetName, 'A3'), 'Hello World')
    expectNumber(engine.getCellValue(sheetName, 'B3'), 4)
    expectString(engine.getCellValue(sheetName, 'C3'), 'HELLO')
    expectString(engine.getCellValue(sheetName, 'D3'), 'hello')
    expectNumber(engine.getCellValue(sheetName, 'E3'), 5)
  })

  it('updates formulas through row insertion and removal', async () => {
    const engine = await createEngineFromGrid([[1], [2], [3], ['=SUM(A1:A3)']])

    expectNumber(engine.getCellValue(sheetName, 'A4'), 6)

    engine.insertRows(sheetName, 1, 1)
    engine.setCellValue(sheetName, 'A2', 10)

    expectNumber(engine.getCellValue(sheetName, 'A5'), 16)

    engine.deleteRows(sheetName, 1, 1)

    expectNumber(engine.getCellValue(sheetName, 'A4'), 6)
  })
})

async function createEngineFromGrid(rows: readonly (readonly SheetCellInput[])[]): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'hyperformula-smoke-parity' })
  await engine.ready()
  engine.createSheet(sheetName)

  rows.forEach((rowValues, row) => {
    rowValues.forEach((input, col) => {
      const address = formatAddress(row, col)
      if (typeof input === 'string' && input.startsWith('=')) {
        engine.setCellFormula(sheetName, address, input.slice(1))
        return
      }
      engine.setCellValue(sheetName, address, input)
    })
  })

  return engine
}

function expectNumber(actual: CellValue, value: number): void {
  expect(actual).toEqual({ tag: ValueTag.Number, value })
}

function expectBoolean(actual: CellValue, value: boolean): void {
  expect(actual).toEqual({ tag: ValueTag.Boolean, value })
}

function expectString(actual: CellValue, value: string): void {
  expect(actual).toEqual({ tag: ValueTag.String, value, stringId: expect.any(Number) })
}
