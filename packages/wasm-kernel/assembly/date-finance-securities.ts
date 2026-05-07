import { ValueTag } from './protocol'
import { addMonthsExcelSerial, couponDaysByBasisValue } from './date-finance'

export function oddLastPriceValue(
  settlementWhole: i32,
  maturityWhole: i32,
  lastInterestWhole: i32,
  rate: f64,
  yieldRate: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    lastInterestWhole >= settlementWhole ||
    settlementWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(yieldRate) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    yieldRate < 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN
  }

  const stepMonths = 12 / frequency
  let periodStart = lastInterestWhole
  let accruedFraction = 0.0
  let remainingFraction = 0.0
  let totalFraction = 0.0
  let iterations = 0
  while (periodStart < maturityWhole && iterations < 32) {
    const normalEndRaw = addMonthsExcelSerial(<u8>ValueTag.Number, <f64>periodStart, <u8>ValueTag.Number, <f64>stepMonths, false)
    if (isNaN(normalEndRaw)) {
      return NaN
    }
    const normalEnd = <i32>normalEndRaw
    if (normalEnd <= periodStart) {
      return NaN
    }
    const actualEnd = normalEnd < maturityWhole ? normalEnd : maturityWhole
    const normalDays = couponDaysByBasisValue(periodStart, normalEnd, basis)
    const countedDays = couponDaysByBasisValue(periodStart, actualEnd, basis)
    if (isNaN(normalDays) || isNaN(countedDays) || normalDays <= 0.0 || countedDays < 0.0) {
      return NaN
    }
    totalFraction += countedDays / normalDays

    if (settlementWhole > periodStart) {
      const accruedEnd = settlementWhole < actualEnd ? settlementWhole : actualEnd
      const accruedDays = couponDaysByBasisValue(periodStart, accruedEnd, basis)
      if (isNaN(accruedDays) || accruedDays < 0.0) {
        return NaN
      }
      accruedFraction += accruedDays / normalDays
    }

    if (settlementWhole < actualEnd) {
      const remainingStart = settlementWhole > periodStart ? settlementWhole : periodStart
      const remainingDays = couponDaysByBasisValue(remainingStart, actualEnd, basis)
      if (isNaN(remainingDays) || remainingDays < 0.0) {
        return NaN
      }
      remainingFraction += remainingDays / normalDays
    }

    periodStart = actualEnd
    iterations += 1
  }

  if (periodStart != maturityWhole || iterations >= 32 || remainingFraction <= 0.0 || totalFraction <= 0.0) {
    return NaN
  }

  const coupon = (100.0 * rate) / <f64>frequency
  const denominator = 1.0 + (yieldRate * remainingFraction) / <f64>frequency
  if (!isFinite(denominator) || denominator <= 0.0) {
    return NaN
  }
  return (redemption + coupon * totalFraction) / denominator - coupon * accruedFraction
}

export function oddLastYieldValue(
  settlementWhole: i32,
  maturityWhole: i32,
  lastInterestWhole: i32,
  rate: f64,
  price: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    lastInterestWhole >= settlementWhole ||
    settlementWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(price) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    price <= 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN
  }

  const stepMonths = 12 / frequency
  let periodStart = lastInterestWhole
  let accruedFraction = 0.0
  let remainingFraction = 0.0
  let totalFraction = 0.0
  let iterations = 0
  while (periodStart < maturityWhole && iterations < 32) {
    const normalEndRaw = addMonthsExcelSerial(<u8>ValueTag.Number, <f64>periodStart, <u8>ValueTag.Number, <f64>stepMonths, false)
    if (isNaN(normalEndRaw)) {
      return NaN
    }
    const normalEnd = <i32>normalEndRaw
    if (normalEnd <= periodStart) {
      return NaN
    }
    const actualEnd = normalEnd < maturityWhole ? normalEnd : maturityWhole
    const normalDays = couponDaysByBasisValue(periodStart, normalEnd, basis)
    const countedDays = couponDaysByBasisValue(periodStart, actualEnd, basis)
    if (isNaN(normalDays) || isNaN(countedDays) || normalDays <= 0.0 || countedDays < 0.0) {
      return NaN
    }
    totalFraction += countedDays / normalDays

    if (settlementWhole > periodStart) {
      const accruedEnd = settlementWhole < actualEnd ? settlementWhole : actualEnd
      const accruedDays = couponDaysByBasisValue(periodStart, accruedEnd, basis)
      if (isNaN(accruedDays) || accruedDays < 0.0) {
        return NaN
      }
      accruedFraction += accruedDays / normalDays
    }

    if (settlementWhole < actualEnd) {
      const remainingStart = settlementWhole > periodStart ? settlementWhole : periodStart
      const remainingDays = couponDaysByBasisValue(remainingStart, actualEnd, basis)
      if (isNaN(remainingDays) || remainingDays < 0.0) {
        return NaN
      }
      remainingFraction += remainingDays / normalDays
    }

    periodStart = actualEnd
    iterations += 1
  }

  if (periodStart != maturityWhole || iterations >= 32 || remainingFraction <= 0.0 || totalFraction <= 0.0) {
    return NaN
  }

  const coupon = (100.0 * rate) / <f64>frequency
  const dirtyPrice = price + coupon * accruedFraction
  if (!isFinite(dirtyPrice) || dirtyPrice <= 0.0) {
    return NaN
  }
  const maturityValue = redemption + coupon * totalFraction
  return ((maturityValue - dirtyPrice) / dirtyPrice) * (<f64>frequency / remainingFraction)
}

export function oddFirstPriceValue(
  settlementWhole: i32,
  maturityWhole: i32,
  issueWhole: i32,
  firstCouponWhole: i32,
  rate: f64,
  yieldRate: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    issueWhole >= settlementWhole ||
    settlementWhole >= firstCouponWhole ||
    firstCouponWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(yieldRate) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN
  }

  const stepMonths = 12 / frequency
  const actualStarts = new Array<i32>()
  const periodEnds = new Array<i32>()
  const normalDaysList = new Array<f64>()
  const countedDaysList = new Array<f64>()

  let periodEnd = firstCouponWhole
  let iterations = 0
  while (periodEnd > issueWhole && iterations < 64) {
    const normalStartRaw = addMonthsExcelSerial(<u8>ValueTag.Number, <f64>periodEnd, <u8>ValueTag.Number, <f64>-stepMonths, false)
    if (isNaN(normalStartRaw)) {
      return NaN
    }
    const normalStart = <i32>normalStartRaw
    if (normalStart >= periodEnd) {
      return NaN
    }
    const actualStart = normalStart > issueWhole ? normalStart : issueWhole
    const normalDays = couponDaysByBasisValue(normalStart, periodEnd, basis)
    const countedDays = couponDaysByBasisValue(actualStart, periodEnd, basis)
    if (isNaN(normalDays) || isNaN(countedDays) || normalDays <= 0.0 || countedDays < 0.0) {
      return NaN
    }
    actualStarts.push(actualStart)
    periodEnds.push(periodEnd)
    normalDaysList.push(normalDays)
    countedDaysList.push(countedDays)
    periodEnd = actualStart
    iterations += 1
  }

  if (periodEnd != issueWhole || iterations >= 64 || actualStarts.length == 0) {
    return NaN
  }

  let accruedFraction = 0.0
  let remainingFraction = 0.0
  let totalFraction = 0.0
  for (let index = actualStarts.length - 1; index >= 0; index -= 1) {
    const actualStart = unchecked(actualStarts[index])
    const segmentEnd = unchecked(periodEnds[index])
    const normalDays = unchecked(normalDaysList[index])
    const countedDays = unchecked(countedDaysList[index])
    totalFraction += countedDays / normalDays

    if (settlementWhole > actualStart) {
      const accruedEnd = settlementWhole < segmentEnd ? settlementWhole : segmentEnd
      const accruedDays = couponDaysByBasisValue(actualStart, accruedEnd, basis)
      if (isNaN(accruedDays) || accruedDays < 0.0) {
        return NaN
      }
      accruedFraction += accruedDays / normalDays
    }

    if (settlementWhole < segmentEnd) {
      const remainingStart = settlementWhole > actualStart ? settlementWhole : actualStart
      const remainingDays = couponDaysByBasisValue(remainingStart, segmentEnd, basis)
      if (isNaN(remainingDays) || remainingDays < 0.0) {
        return NaN
      }
      remainingFraction += remainingDays / normalDays
    }
  }

  let regularPeriodsAfterFirst = 0
  let couponDate = firstCouponWhole
  while (couponDate < maturityWhole && regularPeriodsAfterFirst < 256) {
    const nextCouponRaw = addMonthsExcelSerial(<u8>ValueTag.Number, <f64>couponDate, <u8>ValueTag.Number, <f64>stepMonths, false)
    if (isNaN(nextCouponRaw)) {
      return NaN
    }
    const nextCouponDate = <i32>nextCouponRaw
    if (nextCouponDate <= couponDate) {
      return NaN
    }
    couponDate = nextCouponDate
    regularPeriodsAfterFirst += 1
  }

  if (
    couponDate != maturityWhole ||
    regularPeriodsAfterFirst <= 0 ||
    regularPeriodsAfterFirst >= 256 ||
    remainingFraction <= 0.0 ||
    totalFraction <= 0.0
  ) {
    return NaN
  }

  const discountBase = 1.0 + yieldRate / <f64>frequency
  if (!isFinite(discountBase) || discountBase <= 0.0) {
    return NaN
  }
  const coupon = (100.0 * rate) / <f64>frequency
  let price = (coupon * totalFraction) / Math.pow(discountBase, remainingFraction) - coupon * accruedFraction

  for (let period = 1; period <= regularPeriodsAfterFirst; period += 1) {
    const exponent = remainingFraction + <f64>period
    const cashflow = period == regularPeriodsAfterFirst ? redemption + coupon : coupon
    price += cashflow / Math.pow(discountBase, exponent)
  }
  return price
}

export function oddFirstYieldValue(
  settlementWhole: i32,
  maturityWhole: i32,
  issueWhole: i32,
  firstCouponWhole: i32,
  rate: f64,
  price: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    issueWhole >= settlementWhole ||
    settlementWhole >= firstCouponWhole ||
    firstCouponWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(price) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    price <= 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN
  }

  let lower = -(<f64>frequency) + 1e-10
  let upper = Math.max(1.0, rate * 2.0 + 0.1)
  let lowerPrice = oddFirstPriceValue(
    settlementWhole,
    maturityWhole,
    issueWhole,
    firstCouponWhole,
    rate,
    lower,
    redemption,
    frequency,
    basis,
  )
  let upperPrice = oddFirstPriceValue(
    settlementWhole,
    maturityWhole,
    issueWhole,
    firstCouponWhole,
    rate,
    upper,
    redemption,
    frequency,
    basis,
  )
  for (
    let iteration = 0;
    iteration < 200 && (isNaN(lowerPrice) || isNaN(upperPrice) || lowerPrice < price || upperPrice > price);
    iteration += 1
  ) {
    if (isNaN(upperPrice) || upperPrice > price) {
      upper = upper * 2.0 + 1.0
      upperPrice = oddFirstPriceValue(
        settlementWhole,
        maturityWhole,
        issueWhole,
        firstCouponWhole,
        rate,
        upper,
        redemption,
        frequency,
        basis,
      )
      continue
    }
    lower = (lower - <f64>frequency) / 2.0
    lowerPrice = oddFirstPriceValue(settlementWhole, maturityWhole, issueWhole, firstCouponWhole, rate, lower, redemption, frequency, basis)
  }

  if (isNaN(lowerPrice) || isNaN(upperPrice) || lowerPrice < price || upperPrice > price) {
    return NaN
  }

  let guess = Math.min(Math.max(rate, lower + 1e-8), upper - 1e-8)
  for (let iteration = 0; iteration < 200; iteration += 1) {
    const estimatedPrice = oddFirstPriceValue(
      settlementWhole,
      maturityWhole,
      issueWhole,
      firstCouponWhole,
      rate,
      guess,
      redemption,
      frequency,
      basis,
    )
    if (isNaN(estimatedPrice)) {
      return NaN
    }
    const error = estimatedPrice - price
    if (Math.abs(error) < 1e-14) {
      return guess
    }

    const epsilon = Math.max(1e-7, Math.abs(guess) * 1e-6)
    const shiftedPrice = oddFirstPriceValue(
      settlementWhole,
      maturityWhole,
      issueWhole,
      firstCouponWhole,
      rate,
      guess + epsilon,
      redemption,
      frequency,
      basis,
    )
    const derivative = isNaN(shiftedPrice) ? NaN : (shiftedPrice - estimatedPrice) / epsilon
    let nextGuess = isNaN(derivative) || !isFinite(derivative) || derivative == 0.0 ? (lower + upper) / 2.0 : guess - error / derivative
    if (!isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2.0
    }

    const boundedPrice = oddFirstPriceValue(
      settlementWhole,
      maturityWhole,
      issueWhole,
      firstCouponWhole,
      rate,
      nextGuess,
      redemption,
      frequency,
      basis,
    )
    if (isNaN(boundedPrice)) {
      return NaN
    }
    if (boundedPrice > price) {
      lower = nextGuess
    } else {
      upper = nextGuess
    }
    guess = nextGuess
    if (Math.abs(upper - lower) < 1e-14) {
      return (lower + upper) / 2.0
    }
  }

  return (lower + upper) / 2.0
}

export function couponPriceFromMetricsValue(
  periodsRemaining: i32,
  accruedDays: f64,
  daysToNextCoupon: f64,
  daysInPeriod: f64,
  rate: f64,
  yieldRate: f64,
  redemption: f64,
  frequency: i32,
): f64 {
  if (
    periodsRemaining < 1 ||
    !isFinite(accruedDays) ||
    !isFinite(daysToNextCoupon) ||
    !isFinite(daysInPeriod) ||
    !isFinite(rate) ||
    !isFinite(yieldRate) ||
    !isFinite(redemption) ||
    daysInPeriod <= 0.0 ||
    redemption <= 0.0
  ) {
    return NaN
  }
  const coupon = (100.0 * rate) / <f64>frequency
  const periodsToNextCoupon = daysToNextCoupon / daysInPeriod
  if (periodsRemaining == 1) {
    const denominator = 1.0 + (yieldRate / <f64>frequency) * periodsToNextCoupon
    return denominator <= 0.0 ? NaN : (redemption + coupon) / denominator - coupon * (accruedDays / daysInPeriod)
  }
  const discountBase = 1.0 + yieldRate / <f64>frequency
  if (discountBase <= 0.0) {
    return NaN
  }
  let price = 0.0
  for (let period = 1; period <= periodsRemaining; period += 1) {
    const periodsToCashflow = <f64>(period - 1) + periodsToNextCoupon
    price += coupon / Math.pow(discountBase, periodsToCashflow)
  }
  price += redemption / Math.pow(discountBase, <f64>(periodsRemaining - 1) + periodsToNextCoupon)
  return price - coupon * (accruedDays / daysInPeriod)
}

export function solveCouponYieldValue(
  periodsRemaining: i32,
  accruedDays: f64,
  daysToNextCoupon: f64,
  daysInPeriod: f64,
  rate: f64,
  price: f64,
  redemption: f64,
  frequency: i32,
): f64 {
  if (periodsRemaining < 1 || !isFinite(price) || price <= 0.0) {
    return NaN
  }
  const coupon = (100.0 * rate) / <f64>frequency
  if (periodsRemaining == 1) {
    const dirtyPrice = price + coupon * (accruedDays / daysInPeriod)
    if (dirtyPrice <= 0.0 || daysToNextCoupon <= 0.0) {
      return NaN
    }
    return ((redemption + coupon) / dirtyPrice - 1.0) * <f64>frequency * (daysInPeriod / daysToNextCoupon)
  }

  const targetPrice = price
  let lower = -(<f64>frequency) + 1e-10
  let upper = max<f64>(1.0, rate * 2.0 + 0.1)
  let lowerPrice = couponPriceFromMetricsValue(
    periodsRemaining,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
    rate,
    lower,
    redemption,
    frequency,
  )
  let upperPrice = couponPriceFromMetricsValue(
    periodsRemaining,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
    rate,
    upper,
    redemption,
    frequency,
  )
  for (
    let iteration = 0;
    iteration < 100 && (isNaN(lowerPrice) || isNaN(upperPrice) || lowerPrice < targetPrice || upperPrice > targetPrice);
    iteration += 1
  ) {
    if (isNaN(upperPrice) || upperPrice > targetPrice) {
      upper = upper * 2.0 + 1.0
      upperPrice = couponPriceFromMetricsValue(
        periodsRemaining,
        accruedDays,
        daysToNextCoupon,
        daysInPeriod,
        rate,
        upper,
        redemption,
        frequency,
      )
      continue
    }
    lower = (lower - <f64>frequency) / 2.0
    lowerPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      lower,
      redemption,
      frequency,
    )
  }
  if (isNaN(lowerPrice) || isNaN(upperPrice) || lowerPrice < targetPrice || upperPrice > targetPrice) {
    return NaN
  }

  let guess = min<f64>(max<f64>(rate, lower + 1e-8), upper - 1e-8)
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const estimatedPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      guess,
      redemption,
      frequency,
    )
    if (isNaN(estimatedPrice)) {
      return NaN
    }
    const error = estimatedPrice - targetPrice
    if (Math.abs(error) < 1e-12) {
      return guess
    }
    const epsilon = max<f64>(1e-7, Math.abs(guess) * 1e-6)
    const shiftedPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      guess + epsilon,
      redemption,
      frequency,
    )
    const derivative = isNaN(shiftedPrice) ? NaN : (shiftedPrice - estimatedPrice) / epsilon
    let nextGuess = !isFinite(derivative) || derivative == 0.0 ? (lower + upper) / 2.0 : guess - error / derivative
    if (!isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2.0
    }
    const boundedPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      nextGuess,
      redemption,
      frequency,
    )
    if (isNaN(boundedPrice)) {
      return NaN
    }
    if (boundedPrice > targetPrice) {
      lower = nextGuess
    } else {
      upper = nextGuess
    }
    guess = nextGuess
    if (Math.abs(upper - lower) < 1e-12) {
      return guess
    }
  }
  return guess
}

export function macaulayDurationValue(
  periodsRemaining: i32,
  accruedDays: f64,
  daysToNextCoupon: f64,
  daysInPeriod: f64,
  couponRate: f64,
  yieldRate: f64,
  frequency: i32,
): f64 {
  const price = couponPriceFromMetricsValue(
    periodsRemaining,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
    couponRate,
    yieldRate,
    100.0,
    frequency,
  )
  if (isNaN(price) || price <= 0.0) {
    return NaN
  }
  const discountBase = 1.0 + yieldRate / <f64>frequency
  if (discountBase <= 0.0) {
    return NaN
  }
  const coupon = (100.0 * couponRate) / <f64>frequency
  const periodsToNextCoupon = daysToNextCoupon / daysInPeriod
  let weightedPresentValue = 0.0
  for (let period = 1; period <= periodsRemaining; period += 1) {
    const periodsToCashflow = <f64>(period - 1) + periodsToNextCoupon
    const timeInYears = periodsToCashflow / <f64>frequency
    const cashflow = period == periodsRemaining ? 100.0 + coupon : coupon
    weightedPresentValue += (timeInYears * cashflow) / Math.pow(discountBase, periodsToCashflow)
  }
  return weightedPresentValue / price
}
