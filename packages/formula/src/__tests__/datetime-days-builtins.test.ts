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

describe('DAYS text date semantics', () => {
  it('treats text date arguments as DATEVALUE dates', () => {
    expect(
      datetimeBuiltins.DAYS(
        { tag: ValueTag.String, value: '15-MAR-2021', stringId: 1 },
        { tag: ValueTag.String, value: '1-FEB-2021', stringId: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })
    expect(evaluateAst(parseFormula('DAYS("15-MAR-2021","1-FEB-2021")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })
  })

  it('ignores time components when text dates are DATEVALUE-coerced', () => {
    expect(evaluateAst(parseFormula('DAYS("15-MAR-2021 6:00 AM","1-FEB-2021 11:59 PM")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })
  })

  it('returns #VALUE! for text that cannot be parsed as a valid date', () => {
    expect(
      datetimeBuiltins.DAYS(
        { tag: ValueTag.String, value: 'not-a-date', stringId: 1 },
        { tag: ValueTag.String, value: '1-FEB-2021', stringId: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it("returns #NUM! for numeric date arguments outside Excel's date range", () => {
    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(evaluateAst(parseFormula('DAYS(2958466,1)'), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })
})
