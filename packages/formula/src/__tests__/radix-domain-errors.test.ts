import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

describe('radix conversion domain errors', () => {
  it('returns #NUM! for invalid signed-radix source text and width overflows', () => {
    expect(getBuiltin('BIN2DEC')?.(text('102'))).toEqual(numError)
    expect(getBuiltin('BIN2DEC')?.(text('11111111111'))).toEqual(numError)
    expect(getBuiltin('HEX2DEC')?.(text('G'))).toEqual(numError)
    expect(getBuiltin('OCT2DEC')?.(text('8'))).toEqual(numError)

    expect(getBuiltin('BIN2HEX')?.(text('102'))).toEqual(numError)
    expect(getBuiltin('BIN2OCT')?.(text('11111111111'))).toEqual(numError)
    expect(getBuiltin('HEX2BIN')?.(text('G'))).toEqual(numError)
    expect(getBuiltin('OCT2HEX')?.(text('8'))).toEqual(numError)
  })

  it('returns #NUM! for signed-radix output range and places-domain errors', () => {
    expect(getBuiltin('DEC2BIN')?.(num(512))).toEqual(numError)
    expect(getBuiltin('DEC2HEX')?.(num(549755813888))).toEqual(numError)
    expect(getBuiltin('DEC2OCT')?.(num(536870912))).toEqual(numError)
    expect(getBuiltin('HEX2BIN')?.(text('200'))).toEqual(numError)
    expect(getBuiltin('HEX2OCT')?.(text('20000000'))).toEqual(numError)
    expect(getBuiltin('OCT2BIN')?.(text('1000'))).toEqual(numError)

    expect(getBuiltin('DEC2HEX')?.(num(64), num(1))).toEqual(numError)
    expect(getBuiltin('BIN2HEX')?.(text('1110'), num(0))).toEqual(numError)
    expect(getBuiltin('OCT2HEX')?.(text('100'), num(1))).toEqual(numError)
    expect(getBuiltin('DEC2BIN')?.(num(10), num(-1))).toEqual(numError)
  })

  it('keeps nonnumeric signed-radix decimal inputs and places as #VALUE!', () => {
    expect(getBuiltin('DEC2BIN')?.(text('bad'))).toEqual(valueError)
    expect(getBuiltin('DEC2BIN')?.(num(10), text('bad'))).toEqual(valueError)
    expect(getBuiltin('BIN2HEX')?.(text('1010'), text('bad'))).toEqual(valueError)
    expect(getBuiltin('HEX2BIN')?.(text('F'), text('bad'))).toEqual(valueError)
  })

  it('returns #VALUE! when DECIMAL text exceeds the documented 255-character limit', () => {
    expect(getBuiltin('DECIMAL')?.(text('1'.repeat(255)), num(2))).toEqual(num(5.78960446186581e76))
    expect(getBuiltin('DECIMAL')?.(text('1'.repeat(256)), num(2))).toEqual(valueError)
  })

  it('preserves incoming radix errors before coercion and domain checks', () => {
    expect(getBuiltin('BASE')?.(err(ErrorCode.Ref), num(2))).toEqual(err(ErrorCode.Ref))
    expect(getBuiltin('DECIMAL')?.(text('10'), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(getBuiltin('BIN2DEC')?.(err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(getBuiltin('HEX2DEC')?.(err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(getBuiltin('OCT2HEX')?.(text('10'), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(getBuiltin('DEC2BIN')?.(err(ErrorCode.Div0))).toEqual(err(ErrorCode.Div0))
    expect(getBuiltin('ROMAN')?.(err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(getBuiltin('ARABIC')?.(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
  })
})
