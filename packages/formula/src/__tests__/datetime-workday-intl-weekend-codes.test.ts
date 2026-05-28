import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { datetimeBuiltins, excelDatePartsToSerial } from '../builtins/datetime.js'

const numberValue = (value: number) => ({ tag: ValueTag.Number, value }) as const

describe('WORKDAY.INTL and NETWORKDAYS.INTL weekend codes', () => {
  it('maps Microsoft single-day weekend codes 11 through 17 to Sunday through Saturday', () => {
    const days = [
      { code: 11, before: [2026, 3, 14], weekend: [2026, 3, 15], after: [2026, 3, 16] },
      { code: 12, before: [2026, 3, 15], weekend: [2026, 3, 16], after: [2026, 3, 17] },
      { code: 13, before: [2026, 3, 16], weekend: [2026, 3, 17], after: [2026, 3, 18] },
      { code: 14, before: [2026, 3, 17], weekend: [2026, 3, 18], after: [2026, 3, 19] },
      { code: 15, before: [2026, 3, 18], weekend: [2026, 3, 19], after: [2026, 3, 20] },
      { code: 16, before: [2026, 3, 19], weekend: [2026, 3, 20], after: [2026, 3, 21] },
      { code: 17, before: [2026, 3, 20], weekend: [2026, 3, 21], after: [2026, 3, 22] },
    ] as const

    for (const { code, before, weekend, after } of days) {
      const beforeSerial = excelDatePartsToSerial(...before)
      const weekendSerial = excelDatePartsToSerial(...weekend)
      const afterSerial = excelDatePartsToSerial(...after)
      expect(beforeSerial).toBeDefined()
      expect(weekendSerial).toBeDefined()
      expect(afterSerial).toBeDefined()

      expect(datetimeBuiltins['WORKDAY.INTL'](numberValue(beforeSerial!), numberValue(1), numberValue(code))).toEqual(
        numberValue(afterSerial!),
      )
      expect(datetimeBuiltins['NETWORKDAYS.INTL'](numberValue(weekendSerial!), numberValue(weekendSerial!), numberValue(code))).toEqual(
        numberValue(0),
      )
    }
  })
})
