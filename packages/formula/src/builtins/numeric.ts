import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { ArrayValue } from '../runtime-values.js'
import { parseArithmeticNumericText } from '../numeric-text.js'

interface NumericBuiltinDependencies {
  toNumber: (value: CellValue | undefined) => number | undefined
  numberResult: (value: number) => CellValue
  valueError: () => CellValue
  numError: () => CellValue
}

export interface NumericBuiltinHelpers {
  roundWith: (value: CellValue | undefined, digits: CellValue | undefined) => CellValue
  floorWith: (value: CellValue | undefined, significance?: CellValue) => CellValue
  ceilingWith: (value: CellValue | undefined, significance?: CellValue) => CellValue
  unaryMath: (value: CellValue | undefined, evaluate: (numeric: number) => number) => CellValue
  binaryMath: (left: CellValue | undefined, right: CellValue | undefined, evaluate: (left: number, right: number) => number) => CellValue
}

const ROUND_HALF_TOLERANCE = 1e-12

export function coerceScalarMathNumber(
  value: CellValue | undefined,
  toNumber: (value: CellValue | undefined) => number | undefined,
): number | undefined {
  if (value === undefined) {
    return undefined
  }
  if (value.tag === ValueTag.String) {
    return parseArithmeticNumericText(value.value)
  }
  return toNumber(value)
}

export function createNumericBuiltinHelpers({
  toNumber,
  numberResult,
  valueError,
  numError,
}: NumericBuiltinDependencies): NumericBuiltinHelpers {
  const firstError = (args: readonly (CellValue | undefined)[]): CellValue | undefined =>
    args.find((arg): arg is CellValue => arg?.tag === ValueTag.Error)
  const toScalarMathNumber = (value: CellValue | undefined): number | undefined => coerceScalarMathNumber(value, toNumber)
  const scalarMathResult = (value: number): CellValue => (Number.isFinite(value) ? numberResult(value) : numError())

  return {
    roundWith: (value: CellValue | undefined, digits: CellValue | undefined): CellValue => {
      const error = firstError([value, digits])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const digitValue = digits === undefined ? 0 : toScalarMathNumber(digits)
      if (numberValue === undefined || digitValue === undefined) {
        return valueError()
      }
      return numberResult(roundToDigits(numberValue, Math.trunc(digitValue)))
    },
    floorWith: (value: CellValue | undefined, significance?: CellValue): CellValue => {
      const error = firstError([value, significance])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const significanceValue = significance === undefined ? 1 : toScalarMathNumber(significance)
      if (numberValue === undefined || significanceValue === undefined) {
        return valueError()
      }
      if (significanceValue === 0) {
        return numberValue === 0 ? numberResult(0) : { tag: ValueTag.Error, code: ErrorCode.Div0 }
      }
      if (numberValue > 0 && significanceValue < 0) {
        return { tag: ValueTag.Error, code: ErrorCode.Num }
      }
      return numberResult(floorToMultiple(numberValue, significanceValue))
    },
    ceilingWith: (value: CellValue | undefined, significance?: CellValue): CellValue => {
      const error = firstError([value, significance])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const significanceValue = significance === undefined ? 1 : toScalarMathNumber(significance)
      if (numberValue === undefined || significanceValue === undefined) {
        return valueError()
      }
      if (significanceValue === 0) {
        return numberValue === 0 ? numberResult(0) : { tag: ValueTag.Error, code: ErrorCode.Div0 }
      }
      if (numberValue > 0 && significanceValue < 0) {
        return { tag: ValueTag.Error, code: ErrorCode.Num }
      }
      return numberResult(ceilingToMultiple(numberValue, significanceValue))
    },
    unaryMath: (value: CellValue | undefined, evaluate: (numeric: number) => number): CellValue => {
      if (value?.tag === ValueTag.Error) {
        return value
      }
      const numeric = toScalarMathNumber(value)
      return numeric === undefined ? valueError() : scalarMathResult(evaluate(numeric))
    },
    binaryMath: (
      left: CellValue | undefined,
      right: CellValue | undefined,
      evaluate: (left: number, right: number) => number,
    ): CellValue => {
      const error = firstError([left, right])
      if (error) {
        return error
      }
      const leftNumeric = toScalarMathNumber(left)
      const rightNumeric = toScalarMathNumber(right)
      return leftNumeric === undefined || rightNumeric === undefined ? valueError() : scalarMathResult(evaluate(leftNumeric, rightNumeric))
    },
  }
}

export function roundToDigits(value: number, digits: number): number {
  const sign = value < 0 ? -1 : 1
  const absolute = Math.abs(value)
  if (digits >= 0) {
    const factor = 10 ** digits
    return (sign * roundScaledHalfAwayFromZero(absolute * factor)) / factor
  }
  const factor = 10 ** -digits
  return sign * roundScaledHalfAwayFromZero(absolute / factor) * factor
}

function roundScaledHalfAwayFromZero(value: number): number {
  const floor = Math.floor(value)
  const fraction = value - floor
  return Math.abs(fraction - 0.5) <= ROUND_HALF_TOLERANCE ? floor + 1 : Math.round(value)
}

function snapNearIntegerQuotient(value: number): number {
  if (!Number.isFinite(value)) {
    return value
  }
  const nearest = Math.round(value)
  return Math.abs(value - nearest) <= ROUND_HALF_TOLERANCE ? nearest : value
}

function normalizeMultipleResult(value: number): number {
  if (!Number.isFinite(value) || value === 0) {
    return value
  }
  const absolute = Math.abs(value)
  if (absolute <= ROUND_HALF_TOLERANCE) {
    return 0
  }
  if (absolute >= 1e15) {
    return value
  }
  const exponent = Math.floor(Math.log10(absolute))
  const digits = Math.max(0, Math.min(15, 14 - exponent))
  const normalized = roundToDigits(value, digits)
  return Object.is(normalized, -0) ? 0 : normalized
}

export function floorToMultiple(value: number, multiple: number): number {
  return normalizeMultipleResult(Math.floor(snapNearIntegerQuotient(value / multiple)) * multiple)
}

export function ceilingToMultiple(value: number, multiple: number): number {
  return normalizeMultipleResult(Math.ceil(snapNearIntegerQuotient(value / multiple)) * multiple)
}

export function roundToMultiple(value: number, multiple: number): number {
  const quotient = value / multiple
  const sign = quotient < 0 ? -1 : 1
  return normalizeMultipleResult(sign * roundScaledHalfAwayFromZero(Math.abs(quotient)) * multiple)
}

export function moduloValue(dividend: number, divisor: number): number {
  return normalizeMultipleResult(dividend - divisor * Math.floor(snapNearIntegerQuotient(dividend / divisor)))
}

export function truncateQuotient(numerator: number, denominator: number): number {
  return Math.trunc(snapNearIntegerQuotient(numerator / denominator))
}

export function roundUpToDigits(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits
    return (value >= 0 ? Math.ceil(value * factor) : Math.floor(value * factor)) / factor
  }
  const factor = 10 ** -digits
  return (value >= 0 ? Math.ceil(value / factor) : Math.floor(value / factor)) * factor
}

export function roundDownToDigits(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits
    return (value >= 0 ? Math.floor(value * factor) : Math.ceil(value * factor)) / factor
  }
  const factor = 10 ** -digits
  return (value >= 0 ? Math.floor(value / factor) : Math.ceil(value / factor)) * factor
}

export function roundTowardZero(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits
    return Math.trunc(value * factor) / factor
  }
  const factor = 10 ** -digits
  return Math.trunc(value / factor) * factor
}

export function collectNumericArgs(args: readonly CellValue[], toNumber: (value: CellValue) => number | undefined): number[] {
  return args.map(toNumber).filter((value): value is number => value !== undefined)
}

export function collectStatNumericArgs(args: readonly CellValue[]): number[] {
  const values: number[] = []
  for (const arg of args) {
    if (arg.tag === ValueTag.Number) {
      values.push(arg.value)
      continue
    }
    if (arg.tag === ValueTag.Boolean) {
      values.push(arg.value ? 1 : 0)
    }
  }
  return values
}

export function factorialValue(value: number): number | undefined {
  if (!Number.isFinite(value) || value < 0) {
    return undefined
  }
  const truncated = Math.trunc(value)
  let result = 1
  for (let index = 2; index <= truncated; index += 1) {
    result *= index
    if (!Number.isFinite(result)) {
      return undefined
    }
  }
  return result
}

export function doubleFactorialValue(value: number): number | undefined {
  if (!Number.isFinite(value) || value < 0) {
    return undefined
  }
  const truncated = Math.trunc(value)
  let result = 1
  for (let index = truncated; index >= 2; index -= 2) {
    result *= index
    if (!Number.isFinite(result)) {
      return undefined
    }
  }
  return result
}

export function combinationValue(total: number, chosen: number): number | undefined {
  if (!Number.isFinite(total) || !Number.isFinite(chosen)) {
    return undefined
  }
  const totalValue = Math.trunc(total)
  const chosenValue = Math.trunc(chosen)
  if (totalValue < 0 || chosenValue < 0 || chosenValue > totalValue) {
    return undefined
  }
  const selected = Math.min(chosenValue, totalValue - chosenValue)
  let result = 1
  for (let index = 1; index <= selected; index += 1) {
    result = (result * (totalValue - selected + index)) / index
    if (!Number.isFinite(result)) {
      return undefined
    }
  }
  return result
}

export function permutationValue(total: number, chosen: number): number | undefined {
  if (!Number.isFinite(total) || !Number.isFinite(chosen)) {
    return undefined
  }
  const totalValue = Math.trunc(total)
  const chosenValue = Math.trunc(chosen)
  if (totalValue <= 0 || chosenValue < 0 || chosenValue > totalValue) {
    return undefined
  }
  let result = 1
  for (let index = 0; index < chosenValue; index += 1) {
    result *= totalValue - index
    if (!Number.isFinite(result)) {
      return undefined
    }
  }
  return result
}

export function multinomialValue(values: readonly number[]): number | undefined {
  let total = 0
  let result = 1
  for (const value of values) {
    const ways = combinationValue(total + value, value)
    if (ways === undefined) {
      return undefined
    }
    result *= ways
    if (!Number.isFinite(result)) {
      return undefined
    }
    total += value
  }
  return result
}

export function gcdPair(left: number, right: number): number {
  let a = Math.abs(Math.trunc(left))
  let b = Math.abs(Math.trunc(right))
  while (b !== 0) {
    const next = a % b
    a = b
    b = next
  }
  return a
}

export function lcmPair(left: number, right: number): number {
  const leftTruncated = Math.abs(Math.trunc(left))
  const rightTruncated = Math.abs(Math.trunc(right))
  if (leftTruncated === 0 || rightTruncated === 0) {
    return 0
  }
  return (leftTruncated / gcdPair(leftTruncated, rightTruncated)) * rightTruncated
}

export function evenValue(numberValue: number): number {
  const sign = numberValue < 0 ? -1 : 1
  const rounded = Math.ceil(Math.abs(numberValue) / 2) * 2
  return sign * rounded
}

export function oddValue(numberValue: number): number {
  const sign = numberValue < 0 ? -1 : 1
  const rounded = Math.ceil(Math.abs(numberValue))
  const odd = rounded % 2 === 0 ? rounded + 1 : rounded
  return sign * odd
}

export function buildIdentityMatrix(size: number, numberResult: (value: number) => CellValue): ArrayValue {
  const values: CellValue[] = []
  for (let row = 0; row < size; row += 1) {
    for (let col = 0; col < size; col += 1) {
      values.push(numberResult(row === col ? 1 : 0))
    }
  }
  return { kind: 'array', rows: size, cols: size, values }
}
