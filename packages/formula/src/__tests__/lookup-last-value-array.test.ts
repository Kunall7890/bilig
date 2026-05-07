import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

import { evaluatePlanResult, getLookupBuiltin, lowerToPlan, parseFormula, type RangeBuiltinArgument } from '../index.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const empty = (): CellValue => ({ tag: ValueTag.Empty })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

const cells: Record<string, CellValue> = {
  A1: num(10),
  A2: text('x'),
  A3: num(20),
  A4: empty(),
  A5: num(30),
  A6: text('skip'),
  B1: empty(),
  B2: text('Seed'),
  B3: text(''),
  B4: text('Series A'),
  B5: text('Series C'),
  B6: text(''),
}

function resolveVerticalRange(start: string, end: string): CellValue[] {
  const col = start[0]
  const startRow = Number(start.slice(1))
  const endRow = Number(end.slice(1))
  const values: CellValue[] = []
  for (let row = startRow; row <= endRow; row += 1) {
    values.push(cells[`${col}${row}`] ?? empty())
  }
  return values
}

function evaluate(formula: string) {
  return evaluatePlanResult(lowerToPlan(parseFormula(formula)), {
    sheetName: 'Sheet1',
    resolveCell: (_sheetName: string, address: string) => cells[address] ?? empty(),
    resolveRange: (_sheetName: string, start: string, end: string) => resolveVerticalRange(start, end),
  })
}

describe('LOOKUP last-value array idioms', () => {
  it('vectorizes scalar information functions over range arguments', () => {
    expect(evaluate('ISNUMBER(A1:A6)')).toEqual({
      kind: 'array',
      rows: 6,
      cols: 1,
      values: [bool(true), bool(false), bool(true), bool(false), bool(true), bool(false)],
    })
    expect(evaluate('1/(ISNUMBER(A1:A6))')).toEqual({
      kind: 'array',
      rows: 6,
      cols: 1,
      values: [num(1), err(ErrorCode.Div0), num(1), err(ErrorCode.Div0), num(1), err(ErrorCode.Div0)],
    })
  })

  it('ignores lookup-vector error sentinels during approximate LOOKUP', () => {
    const LOOKUP = getLookupBuiltin('LOOKUP')!
    expect(
      LOOKUP(
        num(2),
        cellRange([err(ErrorCode.Div0), num(1), err(ErrorCode.Div0), num(1), num(1), err(ErrorCode.Div0)], 6, 1),
        cellRange([empty(), text('Seed'), text(''), text('Series A'), text('Series C'), text('')], 6, 1),
      ),
    ).toEqual(text('Series C'))
  })

  it('resolves common finance latest-value LOOKUP array formulas', () => {
    expect(evaluate('LOOKUP(2,1/(ISNUMBER(A1:A6)),A1:A6)')).toEqual(num(30))
    expect(evaluate('LOOKUP(2,1/(B1:B6<>""),B1:B6)')).toEqual(text('Series C'))
  })
})
