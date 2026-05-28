import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getLogicalBuiltin } from '../builtins/logical.js'

const empty = (): CellValue => ({ tag: ValueTag.Empty })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const number = (value: number): CellValue => ({ tag: ValueTag.Number, value })

describe('logical/info builtins', () => {
  it('supports IF with explicit condition coercion and error propagation', () => {
    const IF = getLogicalBuiltin('IF')!

    expect(IF(bool(true), text('yes'), text('no'))).toEqual(text('yes'))
    expect(IF(num(0), text('yes'), text('no'))).toEqual(text('no'))
    expect(IF(empty(), num(1))).toEqual(bool(false))
    expect(IF(bool(false), num(1))).toEqual(bool(false))
    expect(IF(bool(false), num(1), empty())).toEqual(empty())
    expect(IF(text('TRUE'), num(1), num(2))).toEqual(num(1))
    expect(IF(text('FALSE'), num(1), num(2))).toEqual(num(2))
    expect(IF(text(''), num(1), num(2))).toEqual(num(2))
    expect(IF(err(ErrorCode.Ref), num(1), num(2))).toEqual(err(ErrorCode.Ref))
    expect(IF(text('hello'), num(1), num(2))).toEqual(err(ErrorCode.Value))
    expect(IF(bool(true))).toEqual(err(ErrorCode.Value))
  })

  it('supports IFERROR and IFNA with selective error handling', () => {
    const IFERROR = getLogicalBuiltin('IFERROR')!
    const IFNA = getLogicalBuiltin('IFNA')!
    const NA = getLogicalBuiltin('NA')!

    expect(IFERROR(num(7), text('fallback'))).toEqual(num(7))
    expect(IFERROR(err(ErrorCode.Div0), text('fallback'))).toEqual(text('fallback'))
    expect(IFERROR(err(ErrorCode.Value), empty())).toEqual(empty())
    expect(IFERROR()).toEqual(err(ErrorCode.Value))

    expect(NA()).toEqual(err(ErrorCode.NA))
    expect(NA(num(1))).toEqual(err(ErrorCode.Value))
    expect(IFNA(err(ErrorCode.NA), text('missing'))).toEqual(text('missing'))
    expect(IFNA(err(ErrorCode.Ref), text('missing'))).toEqual(err(ErrorCode.Ref))
    expect(IFNA(num(3), text('missing'))).toEqual(num(3))
    expect(IFNA()).toEqual(err(ErrorCode.Value))
  })

  it('supports AND, OR, and NOT with deterministic coercion', () => {
    const AND = getLogicalBuiltin('AND')!
    const OR = getLogicalBuiltin('OR')!
    const NOT = getLogicalBuiltin('NOT')!

    expect(AND(bool(true), num(1), num(-2))).toEqual(bool(true))
    expect(AND(bool(true), empty())).toEqual(bool(false))
    expect(AND(text('TRUE'), text('true'))).toEqual(bool(true))
    expect(AND(text('FALSE'), bool(true))).toEqual(bool(false))
    expect(AND(text(''))).toEqual(bool(false))
    expect(AND(err(ErrorCode.Name), bool(true))).toEqual(err(ErrorCode.Name))
    expect(AND(text('hello'), bool(true))).toEqual(err(ErrorCode.Value))
    expect(AND()).toEqual(err(ErrorCode.Value))

    expect(OR(empty(), bool(true))).toEqual(bool(true))
    expect(OR(num(0), empty())).toEqual(bool(false))
    expect(OR(text('FALSE'), text(''))).toEqual(bool(false))
    expect(OR(text('TRUE'), text('FALSE'))).toEqual(bool(true))
    expect(OR(err(ErrorCode.Ref), bool(true))).toEqual(err(ErrorCode.Ref))
    expect(OR(text('hello'), bool(false))).toEqual(err(ErrorCode.Value))
    expect(OR()).toEqual(err(ErrorCode.Value))

    expect(NOT(bool(false))).toEqual(bool(true))
    expect(NOT(empty())).toEqual(bool(true))
    expect(NOT(num(2))).toEqual(bool(false))
    expect(NOT(text('FALSE'))).toEqual(bool(true))
    expect(NOT(text('TRUE'))).toEqual(bool(false))
    expect(NOT(text(''))).toEqual(bool(true))
    expect(NOT(err(ErrorCode.Div0))).toEqual(err(ErrorCode.Div0))
    expect(NOT(text('hello'))).toEqual(err(ErrorCode.Value))
    expect(NOT()).toEqual(err(ErrorCode.Value))
  })

  it('supports ISBLANK, ISNUMBER, and ISTEXT without propagating errors', () => {
    const ISBLANK = getLogicalBuiltin('ISBLANK')!
    const ISNUMBER = getLogicalBuiltin('ISNUMBER')!
    const ISTEXT = getLogicalBuiltin('ISTEXT')!

    expect(ISBLANK(empty())).toEqual(bool(true))
    expect(ISBLANK(text(''))).toEqual(bool(false))
    expect(ISBLANK(err(ErrorCode.NA))).toEqual(bool(false))

    expect(ISNUMBER(num(42))).toEqual(bool(true))
    expect(ISNUMBER(bool(true))).toEqual(bool(false))
    expect(ISNUMBER(text('42'))).toEqual(bool(false))

    expect(ISTEXT(text('hello'))).toEqual(bool(true))
    expect(ISTEXT(empty())).toEqual(bool(false))
    expect(ISTEXT(err(ErrorCode.Value))).toEqual(bool(false))

    expect(ISBLANK()).toEqual(err(ErrorCode.Value))
    expect(ISNUMBER()).toEqual(err(ErrorCode.Value))
    expect(ISTEXT()).toEqual(err(ErrorCode.Value))
    expect(ISBLANK(empty(), empty())).toEqual(err(ErrorCode.Value))
    expect(ISNUMBER(num(1), num(2))).toEqual(err(ErrorCode.Value))
    expect(ISTEXT(text('x'), text('y'))).toEqual(err(ErrorCode.Value))
  })

  it('supports logical type predicates and ERROR.TYPE', () => {
    const ISERROR = getLogicalBuiltin('ISERROR')!
    const ISERR = getLogicalBuiltin('ISERR')!
    const ISLOGICAL = getLogicalBuiltin('ISLOGICAL')!
    const ISNONTEXT = getLogicalBuiltin('ISNONTEXT')!
    const ISEVEN = getLogicalBuiltin('ISEVEN')!
    const ISODD = getLogicalBuiltin('ISODD')!
    const ISNA = getLogicalBuiltin('ISNA')!
    const ISREF = getLogicalBuiltin('ISREF')!
    const ERROR_TYPE = getLogicalBuiltin('ERROR.TYPE')!

    expect(ISERROR(err(ErrorCode.Div0))).toEqual(bool(true))
    expect(ISERROR(num(1))).toEqual(bool(false))
    expect(ISERR(err(ErrorCode.NA))).toEqual(bool(false))
    expect(ISERR(err(ErrorCode.Name))).toEqual(bool(true))
    expect(ISERR(num(1))).toEqual(bool(false))

    expect(ISLOGICAL(bool(true))).toEqual(bool(true))
    expect(ISLOGICAL(num(0))).toEqual(bool(false))
    expect(ISNONTEXT(text('x'))).toEqual(bool(false))
    expect(ISNONTEXT(num(1))).toEqual(bool(true))

    expect(ISEVEN(num(4))).toEqual(bool(true))
    expect(ISEVEN(text('7'))).toEqual(bool(false))
    expect(ISODD(num(3))).toEqual(bool(true))
    expect(ISODD(empty())).toEqual(bool(false))

    expect(ISNA(err(ErrorCode.NA))).toEqual(bool(true))
    expect(ISNA(err(ErrorCode.Value))).toEqual(bool(false))

    expect(ISREF(text('x'))).toEqual(bool(false))
    expect(ISREF(empty())).toEqual(bool(false))

    expect(ERROR_TYPE(err(ErrorCode.Null))).toEqual(number(1))
    expect(ERROR_TYPE(err(ErrorCode.Div0))).toEqual(number(2))
    expect(ERROR_TYPE(err(ErrorCode.Value))).toEqual(number(3))
    expect(ERROR_TYPE(err(ErrorCode.Ref))).toEqual(number(4))
    expect(ERROR_TYPE(err(ErrorCode.Name))).toEqual(number(5))
    expect(ERROR_TYPE(err(ErrorCode.Num))).toEqual(number(6))
    expect(ERROR_TYPE(err(ErrorCode.NA))).toEqual(number(7))
    expect(ERROR_TYPE(err(ErrorCode.Cycle))).toEqual(err(ErrorCode.NA))
    expect(ERROR_TYPE(num(0))).toEqual(err(ErrorCode.NA))

    for (const predicate of [ISERROR, ISERR, ISLOGICAL, ISNONTEXT, ISEVEN, ISODD, ISNA, ISREF, ERROR_TYPE]) {
      expect(predicate()).toEqual(err(ErrorCode.Value))
      expect(predicate(empty(), empty())).toEqual(err(ErrorCode.Value))
    }
  })

  it('supports TRUE FALSE XOR IFS and SWITCH', () => {
    const TRUE = getLogicalBuiltin('TRUE')!
    const FALSE = getLogicalBuiltin('FALSE')!
    const XOR = getLogicalBuiltin('XOR')!
    const IFS = getLogicalBuiltin('IFS')!
    const SWITCH = getLogicalBuiltin('SWITCH')!

    expect(TRUE()).toEqual(bool(true))
    expect(FALSE()).toEqual(bool(false))
    expect(TRUE(num(1))).toEqual(err(ErrorCode.Value))

    expect(XOR(bool(true), bool(false), num(1))).toEqual(bool(false))
    expect(XOR(text('TRUE'), text('FALSE'), text(''))).toEqual(bool(true))
    expect(XOR(text('TRUE'), text('true'))).toEqual(bool(false))
    expect(XOR()).toEqual(err(ErrorCode.Value))
    expect(XOR(err(ErrorCode.Ref), bool(true))).toEqual(err(ErrorCode.Ref))

    expect(IFS(bool(false), text('no'), num(1), text('yes'))).toEqual(text('yes'))
    expect(IFS(bool(false), text('no'))).toEqual(err(ErrorCode.NA))
    expect(IFS(bool(true))).toEqual(err(ErrorCode.Value))
    expect(IFS(err(ErrorCode.Name), text('no'), bool(true), text('yes'))).toEqual(err(ErrorCode.Name))

    expect(SWITCH(text('b'), text('a'), num(1), text('B'), num(2), num(9))).toEqual(num(2))
    expect(SWITCH(num(12), text('12'), text('numeric text'), text('x'))).toEqual(text('x'))
    expect(SWITCH(text('12'), num(12), text('text numeric'), text('x'))).toEqual(text('x'))
    expect(SWITCH(num(12), text('abc'), text('no'), text('default'))).toEqual(text('default'))
    expect(SWITCH(bool(true), num(1), text('bool numeric'), text('x'))).toEqual(text('x'))
    expect(SWITCH(num(1), bool(true), text('numeric bool'), text('x'))).toEqual(text('x'))
    expect(SWITCH(empty(), text(''), text('empty text'))).toEqual(text('empty text'))
    expect(SWITCH(text('z'), text('a'), num(1), text('b'), num(2))).toEqual(err(ErrorCode.NA))
    expect(SWITCH(err(ErrorCode.Ref), text('a'), num(1))).toEqual(err(ErrorCode.Ref))
    expect(SWITCH(text('b'), err(ErrorCode.Name), num(1), text('b'), num(2))).toEqual(err(ErrorCode.Name))
    expect(getLogicalBuiltin('ISFORMULA')?.(num(1))).toEqual(bool(false))
    expect(getLogicalBuiltin('ISFORMULA')?.()).toEqual(err(ErrorCode.Value))
    expect(getLogicalBuiltin('ISREF')?.()).toEqual(err(ErrorCode.Value))
  })
})
