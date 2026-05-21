import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  coerceNumber,
  coerceText,
  firstError,
  integerValue,
  isErrorValue,
  numberResult,
  truncArg,
  valueError,
} from '../builtins/cell-value-utils.js'

describe('cell value utility helpers', () => {
  it('coerces scalar values using spreadsheet-compatible defaults', () => {
    expect(coerceNumber({ tag: ValueTag.Number, value: 12.75 })).toBe(12.75)
    expect(coerceNumber({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY })).toBeUndefined()
    expect(coerceNumber({ tag: ValueTag.Boolean, value: true })).toBe(1)
    expect(coerceNumber({ tag: ValueTag.Boolean, value: false })).toBe(0)
    expect(coerceNumber({ tag: ValueTag.Empty })).toBe(0)
    expect(coerceNumber({ tag: ValueTag.String, value: '12', stringId: 1 })).toBeUndefined()
    expect(coerceNumber({ tag: ValueTag.Error, code: ErrorCode.Value })).toBeUndefined()
  })

  it('coerces text, integers, and truncation arguments without swallowing errors', () => {
    const error = { tag: ValueTag.Error, code: ErrorCode.Div0 } as const

    expect(coerceText({ tag: ValueTag.String, value: 'text', stringId: 1 })).toBe('text')
    expect(coerceText({ tag: ValueTag.Number, value: 1234.5 })).toBe('1234.5')
    expect(coerceText({ tag: ValueTag.Boolean, value: true })).toBe('TRUE')
    expect(coerceText({ tag: ValueTag.Empty })).toBe('')
    expect(coerceText(error)).toBeUndefined()

    expect(integerValue(undefined, 7)).toBe(7)
    expect(integerValue({ tag: ValueTag.Number, value: 7.9 })).toBe(7)
    expect(integerValue({ tag: ValueTag.String, value: '7', stringId: 2 })).toBeUndefined()
    expect(truncArg(error)).toBe(error)
    expect(truncArg({ tag: ValueTag.String, value: '7', stringId: 3 })).toEqual(valueError())
    expect(truncArg({ tag: ValueTag.Number, value: -7.9 })).toBe(-7)
  })

  it('finds error cells and distinguishes cell values from set carriers', () => {
    const error = { tag: ValueTag.Error, code: ErrorCode.Ref } as const

    expect(numberResult(3)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(firstError([{ tag: ValueTag.Number, value: 1 }, error])).toBe(error)
    expect(firstError([{ tag: ValueTag.Boolean, value: false }])).toBeUndefined()
    expect(isErrorValue(error)).toBe(true)
    expect(isErrorValue(new Set([1, 2]))).toBe(false)
  })
})
