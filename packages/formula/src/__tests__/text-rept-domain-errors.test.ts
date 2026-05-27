import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getTextBuiltin } from '../builtins/text.js'

const maxExcelCellTextLength = 32_767

describe('text result length domain errors', () => {
  it('returns #VALUE when the repeated text would exceed Excel cell text length', () => {
    const REPT = getTextBuiltin('REPT')!

    expect(REPT(text('x'), number(maxExcelCellTextLength))).toEqual(text('x'.repeat(maxExcelCellTextLength)))
    expect(REPT(text('x'), number(maxExcelCellTextLength + 1))).toEqual(valueError())
    expect(REPT(text('xx'), number(16_384))).toEqual(valueError())
    expect(REPT(text('xx'), number(16_383.9))).toEqual(text('xx'.repeat(16_383)))
  })

  it('returns #VALUE when concatenated text would exceed Excel cell text length', () => {
    const CONCAT = getTextBuiltin('CONCAT')!
    const TEXTJOIN = getTextBuiltin('TEXTJOIN')!

    expect(CONCAT(text('x'.repeat(maxExcelCellTextLength)))).toEqual(text('x'.repeat(maxExcelCellTextLength)))
    expect(CONCAT(text('x'.repeat(maxExcelCellTextLength)), text('x'))).toEqual(valueError())
    expect(TEXTJOIN(text(''), boolean(true), text('x'.repeat(maxExcelCellTextLength)))).toEqual(text('x'.repeat(maxExcelCellTextLength)))
    expect(TEXTJOIN(text(''), boolean(true), text('x'.repeat(maxExcelCellTextLength)), text('x'))).toEqual(valueError())
    expect(TEXTJOIN(text('-'), boolean(true), text('x'.repeat(16_383)), text('x'.repeat(16_383)))).toEqual(
      text(`${'x'.repeat(16_383)}-${'x'.repeat(16_383)}`),
    )
    expect(TEXTJOIN(text('-'), boolean(true), text('x'.repeat(16_384)), text('x'.repeat(16_383)))).toEqual(valueError())
  })
})

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function boolean(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}
