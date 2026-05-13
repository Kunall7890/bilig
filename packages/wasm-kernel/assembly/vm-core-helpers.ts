import { Opcode, ValueTag } from './protocol'
import { STACK_KIND_RANGE } from './result-io'

export function isTextLike(tag: u8): bool {
  return tag == ValueTag.String || tag == ValueTag.Empty
}

export function toNumeric(kind: u8, tag: u8, value: f64): f64 {
  if (kind == STACK_KIND_RANGE) return NaN
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value
  if (tag == ValueTag.Empty) return 0
  return NaN
}

export function isComparisonOpcode(opcode: i32): bool {
  return (
    opcode == Opcode.Eq ||
    opcode == Opcode.Neq ||
    opcode == Opcode.Gt ||
    opcode == Opcode.Gte ||
    opcode == Opcode.Lt ||
    opcode == Opcode.Lte
  )
}

export function compareText(left: string, right: string): i32 {
  const normalizedLeft = left.toUpperCase()
  const normalizedRight = right.toUpperCase()
  if (normalizedLeft == normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}

export function ensureU8(buffer: Uint8Array, size: i32): Uint8Array {
  if (buffer.length >= size) return buffer
  let nextLength = buffer.length
  while (nextLength < size) nextLength *= 2
  const next = new Uint8Array(nextLength)
  next.set(buffer)
  return next
}

export function ensureU16(buffer: Uint16Array, size: i32): Uint16Array {
  if (buffer.length >= size) return buffer
  let nextLength = buffer.length
  while (nextLength < size) nextLength *= 2
  const next = new Uint16Array(nextLength)
  next.set(buffer)
  return next
}

export function ensureU32(buffer: Uint32Array, size: i32): Uint32Array {
  if (buffer.length >= size) return buffer
  let nextLength = buffer.length
  while (nextLength < size) nextLength *= 2
  const next = new Uint32Array(nextLength)
  next.set(buffer)
  return next
}

export function ensureF64(buffer: Float64Array, size: i32): Float64Array {
  if (buffer.length >= size) return buffer
  let nextLength = buffer.length
  while (nextLength < size) nextLength *= 2
  const next = new Float64Array(nextLength)
  next.set(buffer)
  return next
}

export function binaryNumeric(op: i32, left: f64, right: f64): f64 {
  if (op == Opcode.Add) return left + right
  if (op == Opcode.Sub) return left - right
  if (op == Opcode.Mul) return left * right
  if (op == Opcode.Div) return right == 0 ? NaN : left / right
  if (op == Opcode.Pow) return excelPower(left, right)
  if (op == Opcode.Eq) return left == right ? 1 : 0
  if (op == Opcode.Neq) return left != right ? 1 : 0
  if (op == Opcode.Gt) return left > right ? 1 : 0
  if (op == Opcode.Gte) return left >= right ? 1 : 0
  if (op == Opcode.Lt) return left < right ? 1 : 0
  if (op == Opcode.Lte) return left <= right ? 1 : 0
  return NaN
}

export function excelPower(base: f64, exponent: f64): f64 {
  const nativeResult = Math.pow(base, exponent)
  if (isFinite(nativeResult) || base >= 0.0 || !isFinite(base) || !isFinite(exponent) || exponent == Math.trunc(exponent)) {
    return nativeResult
  }

  const sign = exponent < 0.0 ? -1 : 1
  const absoluteExponent = Math.abs(exponent)
  let bestNumerator = 0
  let bestDenominator = 1
  let bestError = Infinity

  for (let denominator = 1; denominator <= 999; denominator++) {
    const numerator = <i32>Math.round(absoluteExponent * <f64>denominator)
    const candidate = <f64>numerator / <f64>denominator
    const error = Math.abs(candidate - absoluteExponent)
    if (error < bestError) {
      bestNumerator = numerator
      bestDenominator = denominator
      bestError = error
    }
  }

  const tolerance = Math.max(1e-12, absoluteExponent * 1e-12)
  if (bestError > tolerance) {
    return nativeResult
  }

  const divisor = greatestCommonDivisor(bestNumerator, bestDenominator)
  const numerator = sign * (bestNumerator / divisor)
  const denominator = bestDenominator / divisor
  if (denominator % 2 == 0) {
    return nativeResult
  }

  const magnitude = Math.pow(Math.abs(base), <f64>numerator / <f64>denominator)
  const absoluteNumerator = numerator < 0 ? -numerator : numerator
  return absoluteNumerator % 2 == 0 ? magnitude : -magnitude
}

function greatestCommonDivisor(left: i32, right: i32): i32 {
  let a = left < 0 ? -left : left
  let b = right < 0 ? -right : right
  while (b != 0) {
    const next = a % b
    a = b
    b = next
  }
  return a == 0 ? 1 : a
}
