import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

interface RadixBuiltinHelpers {
  toNumber(this: void, value: CellValue): number | undefined
  integerValue(this: void, value: CellValue | undefined, fallback?: number): number | undefined
  nonNegativeIntegerValue(this: void, value: CellValue | undefined, fallback?: number): number | undefined
  valueError(this: void): CellValue
  numberResult(this: void, value: number): CellValue
}

const arabicMaxRomanLength = 255
const decimalMaxTextLength = 255

export function createRadixBuiltins({ toNumber, integerValue, valueError, numberResult }: RadixBuiltinHelpers): Record<string, Builtin> {
  const baseMaxNumber = 2 ** 53
  const baseMaxMinLength = 255

  const numError = (): CellValue => ({ tag: ValueTag.Error, code: ErrorCode.Num })

  return {
    BASE: (numberArg, radixArg, minLengthArg) => baseString(numberArg, radixArg, minLengthArg),
    DECIMAL: (textArg, radixArg) => decimalValue(textArg, radixArg),
    BIN2DEC: (valueArg) => {
      const numeric = signedRadixNumber(valueArg, 2, 10)
      return typeof numeric === 'number' ? numberResult(numeric) : numeric
    },
    BIN2HEX: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 2, 16, 10, 10, -549755813888, 549755813887, placesArg),
    BIN2OCT: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 2, 8, 10, 10, -536870912, 536870911, placesArg),
    DEC2BIN: (valueArg, placesArg) => {
      const numeric = integerValue(valueArg)
      if (numeric === undefined) {
        return valueError()
      }
      const places = coercePlaces(placesArg)
      if (places !== undefined && typeof places !== 'number') {
        return places
      }
      return formatSignedRadixValue(numeric, 2, places, 10, -512, 511)
    },
    DEC2HEX: (valueArg, placesArg) => {
      const numeric = integerValue(valueArg)
      if (numeric === undefined) {
        return valueError()
      }
      const places = coercePlaces(placesArg)
      if (places !== undefined && typeof places !== 'number') {
        return places
      }
      return formatSignedRadixValue(numeric, 16, places, 10, -549755813888, 549755813887)
    },
    DEC2OCT: (valueArg, placesArg) => {
      const numeric = integerValue(valueArg)
      if (numeric === undefined) {
        return valueError()
      }
      const places = coercePlaces(placesArg)
      if (places !== undefined && typeof places !== 'number') {
        return places
      }
      return formatSignedRadixValue(numeric, 8, places, 10, -536870912, 536870911)
    },
    HEX2BIN: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 16, 2, 10, 10, -512, 511, placesArg),
    HEX2DEC: (valueArg) => {
      const numeric = signedRadixNumber(valueArg, 16, 10)
      return typeof numeric === 'number' ? numberResult(numeric) : numeric
    },
    HEX2OCT: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 16, 8, 10, 10, -536870912, 536870911, placesArg),
    OCT2BIN: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 8, 2, 10, 10, -512, 511, placesArg),
    OCT2DEC: (valueArg) => {
      const numeric = signedRadixNumber(valueArg, 8, 10)
      return typeof numeric === 'number' ? numberResult(numeric) : numeric
    },
    OCT2HEX: (valueArg, placesArg) => convertSignedRadixToRadix(valueArg, 8, 16, 10, 10, -549755813888, 549755813887, placesArg),
    ROMAN: (value, formArg) => {
      const numberValue = romanNumberArgument(value)
      const formValue = romanFormArgument(formArg)
      if (numberValue === undefined || formValue === undefined) {
        return valueError()
      }
      const roman = romanValue(numberValue, formValue)
      return roman === undefined ? valueError() : { tag: ValueTag.String, value: roman, stringId: 0 }
    },
    ARABIC: (value) => {
      if (value.tag !== ValueTag.String) {
        return valueError()
      }
      const numeric = arabicValue(value.value)
      return numeric === undefined ? valueError() : numberResult(numeric)
    },
  }

  function baseString(numberArg: CellValue, radixArg: CellValue, minLengthArg?: CellValue): CellValue {
    const numberValue = integerValue(numberArg)
    const radixValue = integerValue(radixArg)
    const minLengthValue = integerValue(minLengthArg, 0)
    if (numberValue === undefined || radixValue === undefined || minLengthValue === undefined) {
      return valueError()
    }
    if (
      numberValue < 0 ||
      numberValue >= baseMaxNumber ||
      radixValue < 2 ||
      radixValue > 36 ||
      minLengthValue < 0 ||
      minLengthValue > baseMaxMinLength
    ) {
      return numError()
    }
    return {
      tag: ValueTag.String,
      value: numberValue.toString(radixValue).toUpperCase().padStart(minLengthValue, '0'),
      stringId: 0,
    }
  }

  function decimalValue(textArg: CellValue, radixArg: CellValue): CellValue {
    if (textArg.tag === ValueTag.Error) {
      return textArg
    }
    const radixValue = integerValue(radixArg)
    if (radixValue === undefined) {
      return valueError()
    }
    const raw = textArg.tag === ValueTag.String ? textArg.value.trim() : String(Math.trunc(toNumber(textArg) ?? Number.NaN))
    if (raw.length > decimalMaxTextLength) {
      return valueError()
    }
    if (radixValue < 2 || radixValue > 36 || raw === '' || raw === 'NaN' || !isValidBaseDigits(raw, radixValue)) {
      return numError()
    }
    return numberResult(Number.parseInt(raw, radixValue))
  }

  function coercePlaces(value: CellValue | undefined): number | undefined | CellValue {
    if (value === undefined) {
      return undefined
    }
    const places = integerValue(value)
    if (places === undefined) {
      return valueError()
    }
    return places < 0 ? numError() : places
  }

  function signedRadixNumber(valueArg: CellValue, radix: 2 | 8 | 16, width: number): number | CellValue {
    if (valueArg.tag === ValueTag.Error) {
      return valueError()
    }
    const numeric = parseSignedRadixText(valueArg, radix, width, toNumber)
    return numeric === undefined ? numError() : numeric
  }

  function convertSignedRadixToRadix(
    textArg: CellValue,
    fromRadix: 2 | 8 | 16,
    toRadix: 2 | 8 | 16,
    inputWidth: number,
    negativeWidth: number,
    min: number,
    max: number,
    placesArg?: CellValue,
  ): CellValue {
    const numeric = signedRadixNumber(textArg, fromRadix, inputWidth)
    if (typeof numeric !== 'number') {
      return numeric
    }
    const places = coercePlaces(placesArg)
    if (places !== undefined && typeof places !== 'number') {
      return places
    }
    return formatSignedRadixValue(numeric, toRadix, places, negativeWidth, min, max)
  }

  function formatSignedRadixValue(
    numeric: number,
    radix: 2 | 8 | 16,
    minLength: number | undefined,
    negativeWidth: number,
    min: number,
    max: number,
  ): CellValue {
    if (!Number.isSafeInteger(numeric) || numeric < min || numeric > max) {
      return numError()
    }
    if (numeric < 0) {
      const encoded = numeric + radix ** negativeWidth
      return {
        tag: ValueTag.String,
        value: encoded.toString(radix).toUpperCase().padStart(negativeWidth, '0'),
        stringId: 0,
      }
    }
    const raw = numeric.toString(radix).toUpperCase()
    const targetLength = minLength ?? raw.length
    if (targetLength < raw.length) {
      return numError()
    }
    return {
      tag: ValueTag.String,
      value: raw.padStart(targetLength, '0'),
      stringId: 0,
    }
  }
  function romanNumberArgument(value: CellValue): number | undefined {
    if (value.tag === ValueTag.String) {
      return numericTextValue(value.value)
    }
    const numeric = toNumber(value)
    return numeric !== undefined && Number.isFinite(numeric) ? numeric : undefined
  }

  function romanFormArgument(value: CellValue | undefined): number | undefined {
    if (value === undefined) {
      return 0
    }
    if (value.tag === ValueTag.Boolean) {
      return value.value ? 0 : 4
    }
    const numeric = value.tag === ValueTag.String ? numericTextValue(value.value) : toNumber(value)
    if (numeric === undefined || !Number.isFinite(numeric)) {
      return undefined
    }
    const form = Math.trunc(numeric)
    return form >= 0 && form <= 4 ? form : undefined
  }
}

function numericTextValue(text: string): number | undefined {
  const trimmed = text.trim()
  if (trimmed === '') {
    return 0
  }
  const numeric = Number(trimmed)
  return Number.isFinite(numeric) ? numeric : undefined
}

function romanValue(numberValue: number, form: number): string | undefined {
  const number = Math.trunc(numberValue)
  if (!Number.isFinite(numberValue) || number < 1 || number > 3999) {
    return undefined
  }
  const work = {
    value: number % 1000,
    result: 'M'.repeat(Math.floor(number / 1000)),
  }
  if (form === 4) {
    absorbRoman(work, 'IM', 999, 1000)
    absorbRoman(work, 'ID', 499, 500)
  }
  if (form >= 3) {
    absorbRoman(work, 'VM', 995, 1000)
    absorbRoman(work, 'VD', 495, 500)
  }
  if (form >= 2) {
    absorbRoman(work, 'XM', 990, 1000)
    absorbRoman(work, 'XD', 490, 500)
  }
  if (form >= 1) {
    absorbRoman(work, 'LM', 950, 1000)
    absorbRoman(work, 'LD', 450, 500)
  }
  absorbRoman(work, 'CM', 900, 1000)
  absorbRoman(work, 'CD', 400, 500)
  absorbRoman(work, 'D', 500, 900)
  work.result += 'C'.repeat(Math.floor(work.value / 100))
  work.value %= 100
  if (form >= 2) {
    absorbRoman(work, 'IC', 99, 100)
    absorbRoman(work, 'IL', 49, 50)
  }
  if (form >= 1) {
    absorbRoman(work, 'VC', 95, 100)
    absorbRoman(work, 'VL', 45, 50)
  }
  absorbRoman(work, 'XC', 90, 100)
  absorbRoman(work, 'XL', 40, 50)
  absorbRoman(work, 'L', 50, 90)
  work.result += 'X'.repeat(Math.floor(work.value / 10))
  work.value %= 10
  absorbRoman(work, 'IX', 9, 10)
  absorbRoman(work, 'IV', 4, 5)
  absorbRoman(work, 'V', 5, 9)
  work.result += 'I'.repeat(work.value)
  return work.result
}

function absorbRoman(work: { value: number; result: string }, token: string, lower: number, upper: number): void {
  if (work.value >= lower && work.value < upper) {
    work.value -= lower
    work.result += token
  }
}

function arabicValue(text: string): number | undefined {
  const numerals = new Map<string, number>([
    ['I', 1],
    ['V', 5],
    ['X', 10],
    ['L', 50],
    ['C', 100],
    ['D', 500],
    ['M', 1000],
  ])
  const upper = text.trim().toUpperCase()
  if (upper === '') {
    return 0
  }
  const sign = upper.startsWith('-') ? -1 : 1
  const roman = sign < 0 ? upper.slice(1) : upper
  if (roman.length > arabicMaxRomanLength || !/^[IVXLCDM]+$/.test(roman)) {
    return undefined
  }
  if (/(I{4}|X{4}|C{4}|D{4}|V{2}|L{2})/.test(roman)) {
    return undefined
  }
  const subtractivePairs = new Set([
    'IV',
    'IX',
    'IL',
    'IC',
    'ID',
    'IM',
    'VL',
    'VC',
    'VD',
    'VM',
    'XL',
    'XC',
    'XD',
    'XM',
    'LD',
    'LM',
    'CD',
    'CM',
  ])
  let total = 0
  let index = 0
  while (index < roman.length) {
    const current = numerals.get(roman[index] ?? '')
    const next = numerals.get(roman[index + 1] ?? '')
    if (current === undefined) {
      return undefined
    }
    if (next !== undefined && current < next) {
      const pair = `${roman[index] ?? ''}${roman[index + 1] ?? ''}`
      if (!subtractivePairs.has(pair)) {
        return undefined
      }
      if (roman[index - 1] === roman[index]) {
        return undefined
      }
      total += next - current
      index += 2
      continue
    }
    total += current
    index += 1
  }
  return sign * total
}

function isValidBaseDigits(raw: string, radix: number): boolean {
  const upper = raw.toUpperCase()
  const digits = new Set('0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'.slice(0, radix))
  for (const char of upper) {
    if (!digits.has(char)) {
      return false
    }
  }
  return true
}

function parseSignedRadixText(
  textArg: CellValue,
  radix: 2 | 8 | 16,
  width: number,
  toNumber: RadixBuiltinHelpers['toNumber'],
): number | undefined {
  if (textArg.tag === ValueTag.Error) {
    return undefined
  }
  const raw =
    textArg.tag === ValueTag.String ? textArg.value.trim().toUpperCase() : String(Math.trunc(toNumber(textArg) ?? Number.NaN)).toUpperCase()
  if (raw === '' || raw === 'NAN' || raw.length > width || !isValidBaseDigits(raw, radix)) {
    return undefined
  }
  const parsed = Number.parseInt(raw, radix)
  if (!Number.isFinite(parsed)) {
    return undefined
  }
  const fullRange = radix ** width
  const negativeThreshold = fullRange / 2
  return raw.length === width && parsed >= negativeThreshold ? parsed - fullRange : parsed
}
