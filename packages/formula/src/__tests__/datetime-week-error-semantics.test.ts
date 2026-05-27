import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { datetimeBuiltins, excelDatePartsToSerial } from '../builtins/datetime.js'

describe('datetime week error semantics', () => {
  it('preserves Excel 1900-system WEEKDAY compatibility before March 1 1900', () => {
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 59 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 60 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 61 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('returns #NUM for DATE year domains documented by Excel', () => {
    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('returns #NUM for invalid ISOWEEKNUM date domains', () => {
    expect(datetimeBuiltins.ISOWEEKNUM({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.ISOWEEKNUM({ tag: ValueTag.Number, value: 2_958_466 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('returns #NUM for invalid WEEKDAY return types and date domains', () => {
    const serial = excelDatePartsToSerial(2026, 3, 15)!

    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: serial }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: serial }, { tag: ValueTag.Number, value: 99 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: 2_958_466 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('returns #NUM for invalid WEEKNUM return types and date domains', () => {
    const serial = excelDatePartsToSerial(2026, 3, 15)!

    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: serial }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: serial }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: 2_958_466 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('returns #NUM for invalid EOMONTH date domains', () => {
    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 2_958_465 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('returns #VALUE for invalid EDATE start date domains documented by Excel', () => {
    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })
})
