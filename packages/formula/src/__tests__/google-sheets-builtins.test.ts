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
    if (start === 'A1' && end === 'B1') {
      return [text('north'), text('south')]
    }
    if (start === 'A1' && end === 'B2') {
      return [text('north'), text('south'), text('north'), empty()]
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

  it('supports Google Sheets array helpers', () => {
    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2)

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
    expect(getLookupBuiltin('ARRAY_CONSTRAIN')?.(matrix, num(0), num(1))).toEqual(err(ErrorCode.Value))
  })

  it('evaluates Google Sheets helpers through the JS plan path', () => {
    expect(compileFormula('ARRAY_CONSTRAIN(A1:B2,1,2)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('FLATTEN(A1:B2)')).toMatchObject({ mode: 0, producesSpill: true })
    expect(compileFormula('JOIN("|",A1:B1)')).toMatchObject({ mode: 0, producesSpill: false })

    expect(evaluatePlan(lowerToPlan(parseFormula('JOIN("|",A1:B1)')), context)).toEqual(text('north|south'))
    expect(evaluatePlan(lowerToPlan(parseFormula('COUNTUNIQUE(A1:B2)')), context)).toEqual(num(2))
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
  })
})
