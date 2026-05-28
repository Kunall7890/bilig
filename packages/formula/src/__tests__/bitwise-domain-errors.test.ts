import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })

const numError = { tag: ValueTag.Error, code: ErrorCode.Num } as const
const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

describe('bitwise math domain errors', () => {
  it('uses full 48-bit integer semantics instead of 32-bit JavaScript wrapping', () => {
    const high = 2 ** 40

    expect(getBuiltin('BITAND')?.(num(high + 3), num(high + 1))).toEqual(num(high + 1))
    expect(getBuiltin('BITOR')?.(num(high + 3), num(high + 5))).toEqual(num(high + 7))
    expect(getBuiltin('BITXOR')?.(num(high + 3), num(5))).toEqual(num(high + 6))
    expect(getBuiltin('BITLSHIFT')?.(num(high), num(4))).toEqual(num(2 ** 44))
    expect(getBuiltin('BITRSHIFT')?.(num(high), num(4))).toEqual(num(2 ** 36))
  })

  it('returns #NUM! for documented bitwise numeric-domain violations', () => {
    const tooLarge = 2 ** 48

    expect(getBuiltin('BITAND')?.(num(-1), num(1))).toEqual(numError)
    expect(getBuiltin('BITOR')?.(num(tooLarge), num(1))).toEqual(numError)
    expect(getBuiltin('BITXOR')?.(num(1.5), num(1))).toEqual(numError)
    expect(getBuiltin('BITLSHIFT')?.(num(-1), num(1))).toEqual(numError)
    expect(getBuiltin('BITRSHIFT')?.(num(tooLarge), num(1))).toEqual(numError)
    expect(getBuiltin('BITLSHIFT')?.(num(4), num(54))).toEqual(numError)
    expect(getBuiltin('BITRSHIFT')?.(num(4), num(-54))).toEqual(numError)
    expect(getBuiltin('BITLSHIFT')?.(num(4), num(1.5))).toEqual(numError)
  })

  it('keeps nonnumeric bitwise arguments as #VALUE!', () => {
    expect(getBuiltin('BITAND')?.(text('bad'), num(1))).toEqual(valueError)
    expect(getBuiltin('BITOR')?.(num(1), text('bad'))).toEqual(valueError)
    expect(getBuiltin('BITXOR')?.(text('bad'), num(1))).toEqual(valueError)
    expect(getBuiltin('BITLSHIFT')?.(num(4), text('bad'))).toEqual(valueError)
    expect(getBuiltin('BITRSHIFT')?.(text('bad'), num(1))).toEqual(valueError)
  })

  it('supports documented negative shift direction reversal', () => {
    expect(getBuiltin('BITLSHIFT')?.(num(8), num(-1))).toEqual(num(4))
    expect(getBuiltin('BITRSHIFT')?.(num(8), num(-1))).toEqual(num(16))
  })

  it('requires exactly two bitwise arguments', () => {
    for (const name of ['BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT']) {
      const builtin = getBuiltin(name)!
      expect(builtin(num(1))).toEqual(valueError)
      expect(builtin(num(6), num(3), num(1))).toEqual(valueError)
    }
  })
})
