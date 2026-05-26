import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { ArrayValue } from '../runtime-values.js'
import { parseArithmeticNumericText } from '../numeric-text.js'

interface NumericBuiltinDependencies {
  toNumber: (value: CellValue) => number | undefined
  numberResult: (value: number) => CellValue
  valueError: () => CellValue
  numError: () => CellValue
}

export interface NumericBuiltinHelpers {
  roundWith: (value: CellValue, digits: CellValue | undefined) => CellValue
  floorWith: (value: CellValue, significance?: CellValue) => CellValue
  ceilingWith: (value: CellValue, significance?: CellValue) => CellValue
  unaryMath: (value: CellValue, evaluate: (numeric: number) => number) => CellValue
  binaryMath: (left: CellValue, right: CellValue, evaluate: (left: number, right: number) => number) => CellValue
}

export function coerceScalarMathNumber(value: CellValue, toNumber: (value: CellValue) => number | undefined): number | undefined {
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
  const toScalarMathNumber = (value: CellValue): number | undefined => coerceScalarMathNumber(value, toNumber)
  const scalarMathResult = (value: number): CellValue => (Number.isFinite(value) ? numberResult(value) : numError())

  return {
    roundWith: (value: CellValue, digits: CellValue | undefined): CellValue => {
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
    floorWith: (value: CellValue, significance?: CellValue): CellValue => {
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
        return { tag: ValueTag.Error, code: ErrorCode.Div0 }
      }
      if (numberValue > 0 && significanceValue < 0) {
        return { tag: ValueTag.Error, code: ErrorCode.Num }
      }
      return numberResult(Math.floor(numberValue / significanceValue) * significanceValue)
    },
    ceilingWith: (value: CellValue, significance?: CellValue): CellValue => {
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
        return { tag: ValueTag.Error, code: ErrorCode.Div0 }
      }
      if (numberValue > 0 && significanceValue < 0) {
        return { tag: ValueTag.Error, code: ErrorCode.Num }
      }
      return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue)
    },
    unaryMath: (value: CellValue, evaluate: (numeric: number) => number): CellValue => {
      if (value.tag === ValueTag.Error) {
        return value
      }
      const numeric = toScalarMathNumber(value)
      return numeric === undefined ? valueError() : scalarMathResult(evaluate(numeric))
    },
    binaryMath: (left: CellValue, right: CellValue, evaluate: (left: number, right: number) => number): CellValue => {
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
    return (sign * Math.round(absolute * factor)) / factor
  }
  const factor = 10 ** -digits
  return sign * Math.round(absolute / factor) * factor
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
