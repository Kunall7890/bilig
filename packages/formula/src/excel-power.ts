const oddRootReciprocalTolerance = 1e-12
const maxOddRootDenominator = 1_000_000

export function excelPower(base: number, exponent: number): number {
  return base ** exponent
}

export function excelExponentiation(base: number, exponent: number): number {
  if (base < 0 && isOddRootReciprocalExponent(exponent)) {
    return -((-base) ** exponent)
  }
  return base ** exponent
}

function isOddRootReciprocalExponent(exponent: number): boolean {
  if (!Number.isFinite(exponent) || exponent === 0) {
    return false
  }
  const reciprocal = 1 / exponent
  if (!Number.isFinite(reciprocal)) {
    return false
  }
  const denominator = Math.round(Math.abs(reciprocal))
  if (denominator < 3 || denominator > maxOddRootDenominator || denominator % 2 === 0) {
    return false
  }
  const tolerance = Math.max(oddRootReciprocalTolerance, denominator * oddRootReciprocalTolerance)
  return Math.abs(Math.abs(reciprocal) - denominator) <= tolerance
}
