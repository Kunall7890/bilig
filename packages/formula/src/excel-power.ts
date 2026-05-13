const maxRationalPowerDenominator = 999
const rationalPowerTolerance = 1e-12

interface RationalApproximation {
  readonly numerator: number
  readonly denominator: number
}

export function excelPower(base: number, exponent: number): number {
  const nativeResult = base ** exponent
  if (Number.isFinite(nativeResult) || base >= 0 || !Number.isFinite(base) || !Number.isFinite(exponent) || Number.isInteger(exponent)) {
    return nativeResult
  }

  const rationalExponent = approximateRationalExponent(exponent)
  if (!rationalExponent || rationalExponent.denominator % 2 === 0) {
    return nativeResult
  }

  const magnitude = Math.abs(base) ** (rationalExponent.numerator / rationalExponent.denominator)
  return Math.abs(rationalExponent.numerator) % 2 === 0 ? magnitude : -magnitude
}

function approximateRationalExponent(value: number): RationalApproximation | undefined {
  const sign = value < 0 ? -1 : 1
  const absolute = Math.abs(value)
  let bestNumerator = 0
  let bestDenominator = 1
  let bestError = Number.POSITIVE_INFINITY

  for (let denominator = 1; denominator <= maxRationalPowerDenominator; denominator += 1) {
    const numerator = Math.round(absolute * denominator)
    const candidate = numerator / denominator
    const error = Math.abs(candidate - absolute)
    if (error < bestError) {
      bestNumerator = numerator
      bestDenominator = denominator
      bestError = error
    }
  }

  const tolerance = Math.max(rationalPowerTolerance, absolute * rationalPowerTolerance)
  if (bestError > tolerance) {
    return undefined
  }

  const divisor = greatestCommonDivisor(bestNumerator, bestDenominator)
  return {
    numerator: sign * (bestNumerator / divisor),
    denominator: bestDenominator / divisor,
  }
}

function greatestCommonDivisor(left: number, right: number): number {
  let a = Math.abs(left)
  let b = Math.abs(right)
  while (b !== 0) {
    const next = a % b
    a = b
    b = next
  }
  return a || 1
}
