import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const valueError = err(ErrorCode.Value)
const numError = err(ErrorCode.Num)

describe('radix builtins', () => {
  it('covers base, decimal, signed radix, and roman validation branches', () => {
    const BASE = getBuiltin('BASE')!
    const DECIMAL = getBuiltin('DECIMAL')!
    const BIN2DEC = getBuiltin('BIN2DEC')!
    const BIN2HEX = getBuiltin('BIN2HEX')!
    const DEC2BIN = getBuiltin('DEC2BIN')!
    const DEC2HEX = getBuiltin('DEC2HEX')!
    const HEX2DEC = getBuiltin('HEX2DEC')!
    const OCT2BIN = getBuiltin('OCT2BIN')!
    const ROMAN = getBuiltin('ROMAN')!
    const ARABIC = getBuiltin('ARABIC')!

    expect(BASE(num(31), num(16), num(4))).toEqual(text('001F'))
    expect(BASE(num(-1), num(16))).toEqual(numError)
    expect(BASE(num(31), num(1))).toEqual(numError)
    expect(BASE(num(31), num(37))).toEqual(numError)
    expect(BASE(num(31), num(16), text('bad'))).toEqual(valueError)

    expect(DECIMAL(err(ErrorCode.Ref), num(2))).toEqual(err(ErrorCode.Ref))
    expect(DECIMAL(text(' 1f '), num(16))).toEqual(num(31))
    expect(DECIMAL(num(101), num(2))).toEqual(num(5))
    expect(DECIMAL(text(''), num(10))).toEqual(numError)
    expect(DECIMAL(text('2'), num(2))).toEqual(numError)
    expect(DECIMAL(text('10'), num(37))).toEqual(numError)
    expect(DECIMAL(text('10'), text('bad'))).toEqual(valueError)

    expect(BIN2DEC(text('1111111111'))).toEqual(num(-1))
    expect(BIN2DEC(err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(BIN2DEC(text(''))).toEqual(numError)
    expect(BIN2DEC(text('102'))).toEqual(numError)
    expect(BIN2DEC(text('11111111111'))).toEqual(numError)
    expect(BIN2HEX(text('1111111111'))).toEqual(text('FFFFFFFFFF'))
    expect(BIN2HEX(text('1010'), text('bad'))).toEqual(valueError)
    expect(BIN2HEX(text('1010'), num(0))).toEqual(numError)
    expect(DEC2BIN(num(511), num(10))).toEqual(text('0111111111'))
    expect(DEC2BIN(num(512))).toEqual(numError)
    expect(DEC2HEX(num(-1))).toEqual(text('FFFFFFFFFF'))
    expect(HEX2DEC(text('FFFFFFFFFF'))).toEqual(num(-1))
    expect(OCT2BIN(text('7777777777'))).toEqual(text('1111111111'))

    expect(ROMAN(num(3999))).toEqual(text('MMMCMXCIX'))
    expect(ROMAN(num(499), num(0))).toEqual(text('CDXCIX'))
    expect(ROMAN(num(499), num(1))).toEqual(text('LDVLIV'))
    expect(ROMAN(num(499), num(2))).toEqual(text('XDIX'))
    expect(ROMAN(num(499), num(3))).toEqual(text('VDIV'))
    expect(ROMAN(num(499), num(4))).toEqual(text('ID'))
    expect(ROMAN(num(499), bool(true))).toEqual(text('CDXCIX'))
    expect(ROMAN(num(499), bool(false))).toEqual(text('ID'))
    expect(ROMAN(num(499), num(2.9))).toEqual(text('XDIX'))
    expect(ROMAN(text(' 499.9 '), text(' 2.9 '))).toEqual(text('XDIX'))
    expect(ROMAN(num(499), text(''))).toEqual(text('CDXCIX'))
    expect(ROMAN(num(0))).toEqual(valueError)
    expect(ROMAN(num(499), num(-1))).toEqual(valueError)
    expect(ROMAN(num(499), num(5))).toEqual(valueError)
    expect(ROMAN(num(499), text('bad'))).toEqual(valueError)
    expect(ARABIC(text('XLIV'))).toEqual(num(44))
    expect(ARABIC(text(''))).toEqual(num(0))
    expect(ARABIC(text('-MMXI'))).toEqual(num(-2011))
    expect(ARABIC(text('  mxmvii  '))).toEqual(num(1997))
    expect(ARABIC(text('IIV'))).toEqual(valueError)
    expect(ARABIC(text('VX'))).toEqual(valueError)
    expect(ARABIC(text('LC'))).toEqual(valueError)
    expect(ARABIC(text('DM'))).toEqual(valueError)
    expect(ARABIC(num(44))).toEqual(valueError)
  })

  it('matches Microsoft Excel BASE numeric domain errors', () => {
    const BASE = getBuiltin('BASE')!

    expect(BASE(num(-1), num(16))).toEqual(numError)
    expect(BASE(num(2 ** 53), num(16))).toEqual(numError)
    expect(BASE(num(31), num(1))).toEqual(numError)
    expect(BASE(num(31), num(37))).toEqual(numError)
    expect(BASE(num(31), num(16), num(-1))).toEqual(numError)
    expect(BASE(num(31), num(16), num(256))).toEqual(numError)
    expect(BASE(num(31), num(16), text('bad'))).toEqual(valueError)
  })

  it('uses the minimum necessary output length when radix places are omitted', () => {
    expect(getBuiltin('DEC2BIN')?.(num(9))).toEqual(text('1001'))
    expect(getBuiltin('DEC2HEX')?.(num(28))).toEqual(text('1C'))
    expect(getBuiltin('DEC2OCT')?.(num(58))).toEqual(text('72'))
    expect(getBuiltin('DEC2HEX')?.(num(0))).toEqual(text('0'))

    expect(getBuiltin('BIN2HEX')?.(text('1010'))).toEqual(text('A'))
    expect(getBuiltin('BIN2OCT')?.(text('1010'))).toEqual(text('12'))
    expect(getBuiltin('HEX2BIN')?.(text('B7'))).toEqual(text('10110111'))
    expect(getBuiltin('HEX2OCT')?.(text('F'))).toEqual(text('17'))
    expect(getBuiltin('OCT2BIN')?.(text('17'))).toEqual(text('1111'))
    expect(getBuiltin('OCT2HEX')?.(text('17'))).toEqual(text('F'))
  })
})
