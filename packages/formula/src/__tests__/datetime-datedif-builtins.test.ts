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

describe('DATEDIF text date semantics', () => {
  it('treats text date arguments as DATEVALUE dates', () => {
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.String, value: '2001/6/1', stringId: 1 },
        { tag: ValueTag.String, value: '2002/8/15', stringId: 2 },
        { tag: ValueTag.String, value: 'D', stringId: 3 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 440,
    })
    expect(evaluateAst(parseFormula('DATEDIF("2001/6/1","2002/8/15","YD")'), context)).toEqual({
      tag: ValueTag.Number,
      value: 75,
    })
  })

  it('returns #NUM! when the start date is after the end date', () => {
    expect(evaluateAst(parseFormula('DATEDIF("2002/8/15","2001/6/1","D")'), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })
})
