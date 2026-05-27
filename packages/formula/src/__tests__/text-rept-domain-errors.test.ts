import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getTextBuiltin } from '../builtins/text.js'

const maxExcelCellTextLength = 32_767

describe('REPT domain errors', () => {
  it('returns #VALUE when the repeated text would exceed Excel cell text length', () => {
    const REPT = getTextBuiltin('REPT')!

    expect(REPT(text('x'), number(maxExcelCellTextLength))).toEqual(text('x'.repeat(maxExcelCellTextLength)))
    expect(REPT(text('x'), number(maxExcelCellTextLength + 1))).toEqual(valueError())
    expect(REPT(text('xx'), number(16_384))).toEqual(valueError())
    expect(REPT(text('xx'), number(16_383.9))).toEqual(text('xx'.repeat(16_383)))
  })
})

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}
