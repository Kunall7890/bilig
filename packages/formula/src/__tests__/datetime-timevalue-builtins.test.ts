import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { datetimeBuiltins } from '../builtins/datetime.js'
import { evaluateAst } from '../js-evaluator.js'
import { parseFormula } from '../parser.js'

const context = {
  sheetName: 'Sheet1',
  resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
  resolveRange: (): CellValue[] => [],
}

describe('TIMEVALUE date text semantics', () => {
  it('ignores date text and returns the time fraction', () => {
    const expected = (6 * 3600 + 35 * 60) / 86_400

    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '22-Aug-2011 6:35 AM', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: expected,
    })
    expect(evaluateAst(parseFormula('TIMEVALUE("22-Aug-2011 6:35 AM")'), context)).toEqual({
      tag: ValueTag.Number,
      value: expected,
    })
  })

  it('normalizes overflow time text and returns the time fraction', () => {
    const expected = 1 / 24

    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '25:00:00', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: expected,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '1/1/2024 25:00:00', stringId: 2 })).toEqual({
      tag: ValueTag.Number,
      value: expected,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '1:75 PM', stringId: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 14.25 / 24,
    })
  })

  it('preserves fractional seconds in time text for time-part extraction', () => {
    const parsed = datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:00:00.6', stringId: 1 })
    if (parsed.tag !== ValueTag.Number) {
      throw new Error('Expected TIMEVALUE fractional seconds to parse as a number')
    }
    expect(parsed.value).toBeCloseTo((12 * 3600 + 0.6) / 86_400, 15)
    expect(evaluateAst(parseFormula('SECOND("12:00:00.6")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(evaluateAst(parseFormula('SECOND(TIMEVALUE("12:00:00.6"))'), context)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
  })

  it('still rejects text with no time component', () => {
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '22-Aug-2011', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })
})
