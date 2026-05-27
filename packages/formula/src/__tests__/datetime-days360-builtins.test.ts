import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { datetimeBuiltins, excelDatePartsToSerial } from '../builtins/datetime.js'

describe('DAYS360 date-time builtin', () => {
  it('treats February month-end start dates as day 30 under the US method', () => {
    const nonLeapFeb28 = excelDatePartsToSerial(2023, 2, 28)!
    const nonLeapMar31 = excelDatePartsToSerial(2023, 3, 31)!
    const leapFeb29 = excelDatePartsToSerial(2024, 2, 29)!
    const leapMar31 = excelDatePartsToSerial(2024, 3, 31)!

    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: nonLeapFeb28 }, { tag: ValueTag.Number, value: nonLeapMar31 })).toEqual({
      tag: ValueTag.Number,
      value: 30,
    })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: leapFeb29 }, { tag: ValueTag.Number, value: leapMar31 })).toEqual({
      tag: ValueTag.Number,
      value: 30,
    })
  })
})
