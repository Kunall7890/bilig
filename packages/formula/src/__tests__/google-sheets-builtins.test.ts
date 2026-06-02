import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'
import { compileFormula, evaluatePlan, evaluatePlanResult, lowerToPlan, parseFormula } from '../index.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const empty = (): CellValue => ({ tag: ValueTag.Empty })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

const queryTable = cellRange(
  [
    text('Region'),
    text('Leads'),
    text('ARR'),
    text('North'),
    num(8),
    num(96_000),
    text('West'),
    num(12),
    num(144_000),
    text('South'),
    num(5),
    num(60_000),
  ],
  4,
  3,
)

const queryGroupTable = cellRange(
  [
    text('Region'),
    text('Segment'),
    text('ARR'),
    text('North'),
    text('SMB'),
    num(96_000),
    text('West'),
    text('Enterprise'),
    num(144_000),
    text('North'),
    text('Enterprise'),
    num(60_000),
    text('West'),
    text('SMB'),
    num(24_000),
  ],
  5,
  3,
)

const context = {
  sheetName: 'Sheet1',
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case 'A1':
        return text('north')
      case 'B1':
        return text('south')
      case 'A2':
        return text('north')
      case 'B2':
        return empty()
      default:
        return empty()
    }
  },
  resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
    if (start === 'A1' && end === 'A2') {
      return [text('north'), text('north')]
    }
    if (start === 'B1' && end === 'B2') {
      return [text('south'), empty()]
    }
    if (start === 'A1' && end === 'B1') {
      return [text('north'), text('south')]
    }
    if (start === 'A1' && end === 'B2') {
      return [text('north'), text('south'), text('north'), empty()]
    }
    if (start === 'A1' && end === 'C4') {
      return queryTable.values
    }
    if (start === 'A2' && end === 'C4') {
      return queryTable.values.slice(3)
    }
    return []
  },
}

describe('Google Sheets compatibility builtins', () => {
  it('supports scalar Google Sheets helpers', () => {
    expect(getBuiltin('JOIN')?.(text('|'), text('north'), num(42), empty(), bool(true))).toEqual(text('north|42||TRUE'))
    expect(getBuiltin('COUNTUNIQUE')?.(text('north'), text('north'), num(42), text(''), empty(), bool(true))).toEqual(num(3))
    expect(getBuiltin('TO_TEXT')?.(num(42.5))).toEqual(text('42.5'))
    expect(getBuiltin('ISEMAIL')?.(text('agent@example.com'))).toEqual(bool(true))
    expect(getBuiltin('ISEMAIL')?.(text('not an email'))).toEqual(bool(false))
    expect(getBuiltin('ISURL')?.(text('https://example.com/workbook'))).toEqual(bool(true))
    expect(getBuiltin('ISURL')?.(text('ftp://example.com/workbook'))).toEqual(bool(false))
    expect(getBuiltin('JOIN')?.(text(','), err(ErrorCode.Ref), text('x'))).toEqual(err(ErrorCode.Ref))
  })

  it('supports COUNTUNIQUEIFS with Google Sheets criteria semantics', () => {
    const COUNTUNIQUEIFS = getLookupBuiltin('COUNTUNIQUEIFS')!
    const uniqueRange = cellRange([text('alice'), text('bob'), text('alice'), text(''), empty(), num(42), num(42)], 7, 1)
    const regionRange = cellRange([text('west'), text('west'), text('east'), text('west'), text('west'), text('west'), text('west')], 7, 1)
    const amountRange = cellRange([num(10), num(20), num(30), num(40), num(50), num(60), num(70)], 7, 1)

    expect(COUNTUNIQUEIFS(uniqueRange, regionRange, text('west'))).toEqual(num(3))
    expect(COUNTUNIQUEIFS(uniqueRange, regionRange, text('west'), amountRange, text('>=50'))).toEqual(num(1))
    expect(COUNTUNIQUEIFS(uniqueRange, regionRange, text('north'))).toEqual(num(0))
    expect(COUNTUNIQUEIFS(uniqueRange, regionRange, text('west'), amountRange, text('<>'))).toEqual(num(3))
    expect(COUNTUNIQUEIFS(uniqueRange, regionRange, text('west'), cellRange([text('west')], 1, 1), text('west'))).toEqual(
      err(ErrorCode.Value),
    )
    expect(COUNTUNIQUEIFS(uniqueRange, regionRange, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(COUNTUNIQUEIFS(cellRange([err(ErrorCode.Ref)], 1, 1), cellRange([text('west')], 1, 1), text('west'))).toEqual(err(ErrorCode.Ref))
  })

  it('supports Google Sheets array helpers', () => {
    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2)
    const sortnTable = cellRange(
      [text('North'), num(96_000), text('West'), num(144_000), text('North'), num(96_000), text('East'), num(60_000)],
      4,
      2,
    )

    expect(getLookupBuiltin('ARRAY_CONSTRAIN')?.(matrix, num(2), num(1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(1), num(3)],
    })
    expect(getLookupBuiltin('FLATTEN')?.(matrix, text('tail'))).toEqual({
      kind: 'array',
      rows: 5,
      cols: 1,
      values: [num(1), num(2), num(3), num(4), text('tail')],
    })
    expect(getLookupBuiltin('SORTN')?.(sortnTable, num(2), num(0), num(2), bool(false))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('West'), num(144_000), text('North'), num(96_000)],
    })
    expect(getLookupBuiltin('SORTN')?.(sortnTable, num(2), num(1), num(2), bool(false))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 2,
      values: [text('West'), num(144_000), text('North'), num(96_000), text('North'), num(96_000)],
    })
    expect(getLookupBuiltin('ARRAY_CONSTRAIN')?.(matrix, num(0), num(1))).toEqual(err(ErrorCode.Value))
  })

  it('supports SORTN tie modes on sort-key values and external sort vectors', () => {
    const rows = cellRange(
      [
        text('Alice'),
        num(100),
        num(90),
        text('Devon'),
        num(100),
        num(95),
        text('Carol'),
        num(80),
        num(85),
        text('Eloise'),
        num(80),
        num(90),
        text('Frank'),
        num(80),
        num(90),
        text('Bob'),
        num(75),
        num(85),
      ],
      6,
      3,
    )
    const scores = cellRange([num(3), num(1), num(2), num(4), num(5), num(6)], 6, 1)
    const externalTiers = cellRange([num(10), num(10), num(8), num(8), num(8), num(7)], 6, 1)

    expect(getLookupBuiltin('SORTN')?.(rows, num(3), num(1), num(2), bool(false))).toEqual({
      kind: 'array',
      rows: 5,
      cols: 3,
      values: [
        text('Alice'),
        num(100),
        num(90),
        text('Devon'),
        num(100),
        num(95),
        text('Carol'),
        num(80),
        num(85),
        text('Eloise'),
        num(80),
        num(90),
        text('Frank'),
        num(80),
        num(90),
      ],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(3), num(2), num(2), bool(false))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [text('Alice'), num(100), num(90), text('Carol'), num(80), num(85), text('Bob'), num(75), num(85)],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(3), num(3), num(2), bool(false))).toEqual({
      kind: 'array',
      rows: 6,
      cols: 3,
      values: rows.values,
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(2), num(2), num(2), bool(false), num(3), bool(false))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [text('Devon'), num(100), num(95), text('Alice'), num(100), num(90)],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(3), num(3), num(2), bool(false), num(3), bool(false))).toEqual({
      kind: 'array',
      rows: 4,
      cols: 3,
      values: [
        text('Devon'),
        num(100),
        num(95),
        text('Alice'),
        num(100),
        num(90),
        text('Eloise'),
        num(80),
        num(90),
        text('Frank'),
        num(80),
        num(90),
      ],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(2), num(0), scores, bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [text('Devon'), num(100), num(95), text('Carol'), num(80), num(85)],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(1), num(1), externalTiers, bool(false))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [text('Alice'), num(100), num(90), text('Devon'), num(100), num(95)],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(2), num(2), externalTiers, bool(false))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [text('Alice'), num(100), num(90), text('Carol'), num(80), num(85)],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(2), num(3), externalTiers, bool(false))).toEqual({
      kind: 'array',
      rows: 5,
      cols: 3,
      values: [
        text('Alice'),
        num(100),
        num(90),
        text('Devon'),
        num(100),
        num(95),
        text('Carol'),
        num(80),
        num(85),
        text('Eloise'),
        num(80),
        num(90),
        text('Frank'),
        num(80),
        num(90),
      ],
    })
    expect(getLookupBuiltin('SORTN')?.(rows, num(1), num(4))).toEqual(err(ErrorCode.Value))
    expect(getLookupBuiltin('SORTN')?.(rows, num(1), num(0), cellRange([num(1), num(2)], 2, 1), bool(true))).toEqual(err(ErrorCode.Value))
  })

  it('supports a local Google Sheets QUERY subset for workbook ranges', () => {
    expect(getLookupBuiltin('QUERY')?.(queryTable, text('select A, C where B >= 8 order by C desc limit 2'), num(1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 2,
      values: [text('Region'), text('ARR'), text('West'), num(144_000), text('North'), num(96_000)],
    })
    expect(getLookupBuiltin('QUERY')?.(queryTable, text('select Col1, Col3 where Col2 < 10 offset 1'), num(1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('Region'), text('ARR'), text('South'), num(60_000)],
    })
    expect(
      getLookupBuiltin('QUERY')?.(
        queryGroupTable,
        text('select A, sum(C), count(C) where C >= 50000 group by A order by sum(C) desc'),
        num(1),
      ),
    ).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [text('Region'), text('sum ARR'), text('count ARR'), text('North'), num(156_000), num(2), text('West'), num(144_000), num(1)],
    })
    expect(getLookupBuiltin('QUERY')?.(queryGroupTable, text('select A, B, sum(C) group by A, B order by A limit 2'), num(1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [
        text('Region'),
        text('Segment'),
        text('sum ARR'),
        text('North'),
        text('SMB'),
        num(96_000),
        text('North'),
        text('Enterprise'),
        num(60_000),
      ],
    })
    expect(getLookupBuiltin('QUERY')?.(queryTable, text('select A group by A'), num(1))).toEqual(err(ErrorCode.Value))
    expect(getLookupBuiltin('QUERY')?.(queryTable, text('select A, avg(C) group by A'), num(1))).toEqual(err(ErrorCode.Value))
    expect(getLookupBuiltin('QUERY')?.(queryTable, text('select A, sum(C) group by B'), num(1))).toEqual(err(ErrorCode.Value))
    expect(getLookupBuiltin('QUERY')?.(queryTable, text('select D'), num(1))).toEqual(err(ErrorCode.Value))
  })

  it('evaluates Google Sheets helpers through the JS plan path', () => {
    expect(compileFormula('ARRAYFORMULA(A1:B2)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('ARRAYFORMULA(A1:A2&"-lead")')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('ARRAY_CONSTRAIN(A1:B2,1,2)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('FLATTEN(A1:B2)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('QUERY(A1:C4,"select A,C where B >= 8",1)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('SORTN(A1:C4,2,0,2,FALSE)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('COUNTUNIQUEIFS(A1:A2,B1:B2,"south")')).toMatchObject({ mode: 0, producesSpill: false })
    expect(compileFormula('JOIN("|",A1:B1)')).toMatchObject({ mode: 0, producesSpill: false })
    expect(compileFormula('GOOGLEFINANCE("GOOG","price")')).toMatchObject({
      mode: 1,
      optimizedAst: { kind: 'ErrorLiteral', code: ErrorCode.Blocked },
      producesSpill: false,
      symbolicNames: [],
    })
    expect(compileFormula('IMPORTDATA("https://example.com/data.csv")')).toMatchObject({
      mode: 1,
      optimizedAst: { kind: 'ErrorLiteral', code: ErrorCode.Blocked },
      symbolicNames: [],
    })
    expect(compileFormula('IMPORTHTML("https://example.com","table",1)')).toMatchObject({
      mode: 1,
      optimizedAst: { kind: 'ErrorLiteral', code: ErrorCode.Blocked },
      symbolicNames: [],
    })
    expect(compileFormula('IMPORTXML("https://example.com","//h1")')).toMatchObject({
      mode: 1,
      optimizedAst: { kind: 'ErrorLiteral', code: ErrorCode.Blocked },
      symbolicNames: [],
    })
    expect(compileFormula('IMPORTFEED("https://example.com/feed.xml")')).toMatchObject({
      mode: 1,
      optimizedAst: { kind: 'ErrorLiteral', code: ErrorCode.Blocked },
      symbolicNames: [],
    })

    expect(evaluatePlan(lowerToPlan(parseFormula('JOIN("|",A1:B1)')), context)).toEqual(text('north|south'))
    expect(evaluatePlan(lowerToPlan(parseFormula('COUNTUNIQUE(A1:B2)')), context)).toEqual(num(2))
    expect(evaluatePlan(lowerToPlan(parseFormula('COUNTUNIQUEIFS(A1:A2,B1:B2,"south")')), context)).toEqual(num(1))
    expect(evaluatePlan(lowerToPlan(parseFormula('GOOGLEFINANCE("GOOG","price")')), context)).toEqual(err(ErrorCode.Blocked))
    expect(evaluatePlan(lowerToPlan(parseFormula('IMPORTDATA("https://example.com/data.csv")')), context)).toEqual(err(ErrorCode.Blocked))
    expect(evaluatePlan(lowerToPlan(parseFormula('IMPORTHTML("https://example.com","table",1)')), context)).toEqual(err(ErrorCode.Blocked))
    expect(evaluatePlan(lowerToPlan(parseFormula('IMPORTXML("https://example.com","//h1")')), context)).toEqual(err(ErrorCode.Blocked))
    expect(evaluatePlan(lowerToPlan(parseFormula('IMPORTFEED("https://example.com/feed.xml")')), context)).toEqual(err(ErrorCode.Blocked))
    expect(evaluatePlanResult(lowerToPlan(parseFormula('ARRAYFORMULA(A1:B2)')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('north'), text('south'), text('north'), empty()],
    })
    expect(evaluatePlanResult(lowerToPlan(parseFormula('ARRAYFORMULA(A1:A2&"-lead")')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [text('north-lead'), text('north-lead')],
    })
    expect(evaluatePlan(lowerToPlan(parseFormula('ARRAYFORMULA()')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('ARRAYFORMULA(A1,A2)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlanResult(lowerToPlan(parseFormula('ARRAY_CONSTRAIN(A1:B2,1,2)')), context)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      sourceRange: { refKind: 'cells', sheetName: 'Sheet1', start: 'A1', end: 'B1' },
      values: [text('north'), text('south')],
    })
    expect(evaluatePlanResult(lowerToPlan(parseFormula('FLATTEN(A1:B2,"done")')), context)).toEqual({
      kind: 'array',
      rows: 5,
      cols: 1,
      values: [text('north'), text('south'), text('north'), empty(), text('done')],
    })
    expect(
      evaluatePlanResult(lowerToPlan(parseFormula('QUERY(A1:C4,"select A,C where B >= 8 order by C desc limit 1",1)')), context),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('Region'), text('ARR'), text('West'), num(144_000)],
    })
    expect(
      evaluatePlanResult(
        lowerToPlan(parseFormula('QUERY(A1:C4,"select A,sum(C) where B >= 8 group by A order by sum(C) desc",1)')),
        context,
      ),
    ).toEqual({
      kind: 'array',
      rows: 3,
      cols: 2,
      values: [text('Region'), text('sum ARR'), text('West'), num(144_000), text('North'), num(96_000)],
    })
    expect(evaluatePlanResult(lowerToPlan(parseFormula('SORTN(A2:C4,2,0,3,FALSE)')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [text('West'), num(12), num(144_000), text('North'), num(8), num(96_000)],
    })
  })
})
