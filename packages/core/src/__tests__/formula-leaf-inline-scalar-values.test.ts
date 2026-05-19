import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import {
  compareInlineScalars,
  inlineArithmeticNumber,
  inlineNumber,
  inlineOptionalNumber,
  inlinePaymentType,
  inlineRequiredNumber,
  inlineStringValue,
  inlineTruthy,
  roundInlineToDigits,
} from '../engine/services/formula-leaf-inline-scalar-values.js'

const emptyValue: CellValue = { tag: ValueTag.Empty }
const numberValue = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const booleanValue = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const stringValue = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const errorValue = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

describe('formula leaf inline scalar values', () => {
  it('coerces inline numbers using formula scalar semantics', () => {
    expect(inlineNumber(numberValue(12.5))).toBe(12.5)
    expect(inlineNumber(booleanValue(true))).toBe(1)
    expect(inlineNumber(booleanValue(false))).toBe(0)
    expect(inlineNumber(emptyValue)).toBe(0)
    expect(inlineNumber(stringValue('1'))).toBeUndefined()
    expect(inlineNumber(errorValue(ErrorCode.Value))).toBeUndefined()

    expect(inlineRequiredNumber(errorValue(ErrorCode.Ref))).toBeUndefined()
    expect(inlineRequiredNumber(booleanValue(true))).toBe(1)
    expect(inlineOptionalNumber(undefined, 7)).toBe(7)
    expect(inlineOptionalNumber(stringValue('bad'), 7)).toBeUndefined()
  })

  it('coerces arithmetic text only for arithmetic-number paths', () => {
    expect(inlineArithmeticNumber(stringValue(''))).toBe(0)
    expect(inlineArithmeticNumber(stringValue(' 1,234.5 '))).toBe(1234.5)
    expect(inlineArithmeticNumber(stringValue('   '))).toBeUndefined()
    expect(inlineArithmeticNumber(numberValue(9))).toBe(9)
  })

  it('normalizes payment type, rounding, truthiness, and display strings', () => {
    expect(inlinePaymentType(undefined, 1)).toBe(1)
    expect(inlinePaymentType(numberValue(0.8), 0)).toBe(0)
    expect(inlinePaymentType(numberValue(1.2), 0)).toBe(1)
    expect(inlinePaymentType(numberValue(2), 0)).toBeUndefined()
    expect(inlinePaymentType(stringValue('1'), 0)).toBeUndefined()

    expect(roundInlineToDigits(12.345, 2)).toBe(12.35)
    expect(roundInlineToDigits(1234.5, -2)).toBe(1200)

    expect(inlineTruthy(numberValue(0))).toBe(false)
    expect(inlineTruthy(booleanValue(true))).toBe(true)
    expect(inlineTruthy(stringValue('text'))).toBe(false)

    expect(inlineStringValue(emptyValue)).toBe('')
    expect(inlineStringValue(numberValue(1234))).toBe('1234')
    expect(inlineStringValue(booleanValue(false))).toBe('FALSE')
    expect(inlineStringValue(stringValue('abc'))).toBe('abc')
    expect(inlineStringValue(errorValue(ErrorCode.Div0))).toBe('#DIV/0!')
  })

  it('compares inline scalar values with spreadsheet ordering', () => {
    expect(compareInlineScalars(stringValue('alpha'), stringValue('ALPHA'))).toBe(0)
    expect(compareInlineScalars(stringValue('alpha'), stringValue('beta'))).toBe(-1)
    expect(compareInlineScalars(stringValue('beta'), stringValue('alpha'))).toBe(1)
    expect(compareInlineScalars(emptyValue, emptyValue)).toBe(0)
    expect(compareInlineScalars(stringValue(''), emptyValue)).toBe(0)
    expect(compareInlineScalars(emptyValue, stringValue('z'))).toBe(-1)
    expect(compareInlineScalars(stringValue('1'), numberValue(1))).toBe(1)
    expect(compareInlineScalars(numberValue(1), stringValue('1'))).toBe(-1)
    expect(compareInlineScalars(booleanValue(true), numberValue(1))).toBe(0)
    expect(compareInlineScalars(numberValue(-0), numberValue(0))).toBe(0)
    expect(compareInlineScalars(numberValue(1), numberValue(2))).toBe(-1)
    expect(compareInlineScalars(numberValue(2), numberValue(1))).toBe(1)
    expect(compareInlineScalars(errorValue(ErrorCode.Value), numberValue(1))).toBeUndefined()
  })
})
