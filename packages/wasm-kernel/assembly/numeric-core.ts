import { toNumberExact } from './operands'

const ROUND_HALF_TOLERANCE: f64 = 1e-12

export function truncAbs(value: f64): f64 {
  if (!isFinite(value)) {
    return NaN
  }
  return Math.abs(value < 0.0 ? Math.ceil(value) : Math.floor(value))
}

export function factorialCalc(value: f64): f64 {
  if (!isFinite(value) || value < 0.0) {
    return NaN
  }
  const truncated = <i32>Math.floor(value)
  let result = 1.0
  for (let index = 2; index <= truncated; index += 1) {
    result *= <f64>index
    if (!isFinite(result)) {
      return NaN
    }
  }
  return result
}

export function doubleFactorialCalc(value: f64): f64 {
  if (!isFinite(value) || value < -1.0) {
    return NaN
  }
  const truncated = <i32>Math.floor(value)
  let result = 1.0
  for (let index = truncated; index >= 2; index -= 2) {
    result *= <f64>index
    if (!isFinite(result)) {
      return NaN
    }
  }
  return result
}

export function combinationCalc(total: f64, chosen: f64): f64 {
  if (!isFinite(total) || !isFinite(chosen)) {
    return NaN
  }
  const totalValue = <i32>Math.trunc(total)
  const chosenValue = <i32>Math.trunc(chosen)
  if (totalValue < 0 || chosenValue < 0 || chosenValue > totalValue) {
    return NaN
  }
  const remainder = totalValue - chosenValue
  const selected = chosenValue < remainder ? chosenValue : remainder
  let result = 1.0
  for (let index = 1; index <= selected; index += 1) {
    result = (result * <f64>(totalValue - selected + index)) / <f64>index
    if (!isFinite(result)) {
      return NaN
    }
  }
  return result
}

export function permutationCalc(total: f64, chosen: f64): f64 {
  if (!isFinite(total) || !isFinite(chosen)) {
    return NaN
  }
  const totalValue = <i32>Math.trunc(total)
  const chosenValue = <i32>Math.trunc(chosen)
  if (totalValue <= 0 || chosenValue < 0 || chosenValue > totalValue) {
    return NaN
  }
  let result = 1.0
  for (let index = 0; index < chosenValue; index += 1) {
    result *= <f64>(totalValue - index)
    if (!isFinite(result)) {
      return NaN
    }
  }
  return result
}

export function gcdPairCalc(left: f64, right: f64): f64 {
  let a = truncAbs(left)
  let b = truncAbs(right)
  while (b != 0.0) {
    const next = a % b
    a = b
    b = next
  }
  return a
}

export function lcmPairCalc(left: f64, right: f64): f64 {
  const a = truncAbs(left)
  const b = truncAbs(right)
  if (a == 0.0 || b == 0.0) {
    return 0.0
  }
  return Math.abs((a * b) / gcdPairCalc(a, b))
}

export function evenCalc(value: f64): f64 {
  const sign = value < 0.0 ? -1.0 : 1.0
  const rounded = Math.ceil(Math.abs(value) / 2.0) * 2.0
  return sign * rounded
}

export function oddCalc(value: f64): f64 {
  const sign = value < 0.0 ? -1.0 : 1.0
  const rounded = Math.ceil(Math.abs(value))
  const odd = rounded % 2.0 == 0.0 ? rounded + 1.0 : rounded
  return sign * odd
}

export function truncToInt(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value)
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>numeric
}

export function roundToDigits(value: f64, digits: i32): f64 {
  const sign = value < 0.0 ? -1.0 : 1.0
  const absolute = Math.abs(value)
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits)
    return (sign * roundScaledHalfAwayFromZero(absolute * factor)) / factor
  }
  const factor = Math.pow(10.0, <f64>-digits)
  return sign * roundScaledHalfAwayFromZero(absolute / factor) * factor
}

function roundScaledHalfAwayFromZero(value: f64): f64 {
  const floor = Math.floor(value)
  const fraction = value - floor
  return Math.abs(fraction - 0.5) <= ROUND_HALF_TOLERANCE ? floor + 1.0 : Math.round(value)
}

function snapNearInteger(value: f64): f64 {
  if (!isFinite(value)) {
    return value
  }
  const nearest = Math.round(value)
  return Math.abs(value - nearest) <= ROUND_HALF_TOLERANCE ? nearest : value
}

function normalizeMultipleResult(value: f64): f64 {
  if (!isFinite(value) || value == 0.0) {
    return value
  }
  const absolute = Math.abs(value)
  if (absolute <= ROUND_HALF_TOLERANCE) {
    return 0.0
  }
  if (absolute >= 1e15) {
    return value
  }
  const exponent = <i32>Math.floor(Math.log(absolute) / Math.log(10.0))
  const rawDigits = 14 - exponent
  const digits = rawDigits < 0 ? 0 : rawDigits > 15 ? 15 : rawDigits
  const normalized = roundToDigits(value, digits)
  return normalized == 0.0 ? 0.0 : normalized
}

export function floorToMultiple(value: f64, multiple: f64): f64 {
  return normalizeMultipleResult(Math.floor(snapNearInteger(value / multiple)) * multiple)
}

export function ceilingToMultiple(value: f64, multiple: f64): f64 {
  return normalizeMultipleResult(Math.ceil(snapNearInteger(value / multiple)) * multiple)
}

export function roundToMultiple(value: f64, multiple: f64): f64 {
  const quotient = value / multiple
  const sign = quotient < 0.0 ? -1.0 : 1.0
  return normalizeMultipleResult(sign * roundScaledHalfAwayFromZero(Math.abs(quotient)) * multiple)
}

export function moduloValue(dividend: f64, divisor: f64): f64 {
  return normalizeMultipleResult(dividend - divisor * Math.floor(snapNearInteger(dividend / divisor)))
}

export function truncateQuotient(numerator: f64, denominator: f64): f64 {
  return Math.trunc(snapNearInteger(numerator / denominator))
}

export function roundAwayFromZeroDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits)
    const scaled = snapNearInteger(value * factor)
    return (value >= 0.0 ? Math.ceil(scaled) : Math.floor(scaled)) / factor
  }
  const factor = Math.pow(10.0, <f64>-digits)
  const scaled = snapNearInteger(value / factor)
  return (value >= 0.0 ? Math.ceil(scaled) : Math.floor(scaled)) * factor
}

export function roundTowardZeroDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits)
    return Math.trunc(snapNearInteger(value * factor)) / factor
  }
  const factor = Math.pow(10.0, <f64>-digits)
  return Math.trunc(snapNearInteger(value / factor)) * factor
}

const MAX_BITWISE_INTEGER_F64: f64 = 281474976710655.0
const MAX_BITWISE_SHIFT_F64: f64 = 53.0
export const BITWISE_VALUE_ERROR: i64 = i64.MIN_VALUE
export const BITWISE_NUM_ERROR: i64 = i64.MIN_VALUE + 1

export function coerceBitwiseInteger(tag: u8, value: f64): i64 {
  const numeric = toNumberExact(tag, value)
  if (!isFinite(numeric)) {
    return BITWISE_VALUE_ERROR
  }
  const truncated = Math.trunc(numeric)
  if (numeric != truncated || truncated < 0.0 || truncated > MAX_BITWISE_INTEGER_F64) {
    return BITWISE_NUM_ERROR
  }
  return <i64>truncated
}

export function coerceBitwiseShift(tag: u8, value: f64): i64 {
  const numeric = toNumberExact(tag, value)
  if (!isFinite(numeric)) {
    return BITWISE_VALUE_ERROR
  }
  const truncated = Math.trunc(numeric)
  if (numeric != truncated || Math.abs(truncated) > MAX_BITWISE_SHIFT_F64) {
    return BITWISE_NUM_ERROR
  }
  return <i64>truncated
}

export function coerceInteger(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value)
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE
  }
  return <i32>numeric
}
