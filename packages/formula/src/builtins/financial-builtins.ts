import type { CellValue } from '@bilig/protocol'
import {
  cumulativePeriodicPayment,
  dbDepreciation,
  ddbDepreciation,
  futureValue,
  interestPayment,
  periodicPayment,
  presentValue,
  principalPayment,
  solveRate,
  totalPeriods,
  vdbDepreciation,
} from './financial.js'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

interface FinancialBuiltinDeps {
  toNumber: (value: CellValue) => number | undefined
  coerceBoolean: (value: CellValue | undefined, fallback: boolean) => boolean | undefined
  coerceNumber: (value: CellValue | undefined, fallback: number) => number | undefined
  coercePaymentType: (value: CellValue | undefined, fallback: number) => number | undefined
  integerValue: (value: CellValue | undefined, fallback?: number) => number | undefined
  numberResult: (value: number) => EvaluationResult
  valueError: () => EvaluationResult
  div0Error: () => EvaluationResult
  numError: () => EvaluationResult
}

export function createFinancialBuiltins({
  toNumber,
  coerceBoolean,
  coerceNumber,
  coercePaymentType,
  integerValue,
  numberResult,
  valueError,
  div0Error,
  numError,
}: FinancialBuiltinDeps): Record<string, Builtin> {
  const cumulativePeriodicPaymentResult = (
    rateArg: CellValue,
    periodsArg: CellValue,
    presentArg: CellValue,
    startPeriodArg: CellValue,
    endPeriodArg: CellValue,
    typeArg: CellValue,
    principalOnly: boolean,
  ): EvaluationResult => {
    const rate = toNumber(rateArg)
    const periods = toNumber(periodsArg)
    const present = toNumber(presentArg)
    const startPeriod = integerValue(startPeriodArg)
    const endPeriod = integerValue(endPeriodArg)
    const typeRaw = toNumber(typeArg)
    if (
      rate === undefined ||
      periods === undefined ||
      present === undefined ||
      startPeriod === undefined ||
      endPeriod === undefined ||
      typeRaw === undefined
    ) {
      return valueError()
    }

    const type = Math.trunc(typeRaw)
    if (
      !Number.isFinite(rate) ||
      !Number.isFinite(periods) ||
      !Number.isFinite(present) ||
      !Number.isFinite(typeRaw) ||
      rate <= 0 ||
      periods <= 0 ||
      present <= 0 ||
      startPeriod < 1 ||
      endPeriod < 1 ||
      startPeriod > endPeriod ||
      type < 0 ||
      type > 1
    ) {
      return numError()
    }

    const total = cumulativePeriodicPayment(rate, periods, present, startPeriod, endPeriod, type, principalOnly)
    return total === undefined ? valueError() : numberResult(total)
  }

  return {
    EFFECT: (nominalRateArg, periodsArg) => {
      const nominalRate = toNumber(nominalRateArg)
      const periodsRaw = toNumber(periodsArg)
      if (nominalRate === undefined || periodsRaw === undefined) {
        return valueError()
      }
      const periods = Math.trunc(periodsRaw)
      if (!Number.isFinite(nominalRate) || !Number.isFinite(periodsRaw) || nominalRate <= 0 || periods < 1) {
        return numError()
      }
      return numberResult((1 + nominalRate / periods) ** periods - 1)
    },
    NOMINAL: (effectiveRateArg, periodsArg) => {
      const effectiveRate = toNumber(effectiveRateArg)
      const periodsRaw = toNumber(periodsArg)
      if (effectiveRate === undefined || periodsRaw === undefined) {
        return valueError()
      }
      const periods = Math.trunc(periodsRaw)
      if (!Number.isFinite(effectiveRate) || !Number.isFinite(periodsRaw) || effectiveRate <= 0 || periods < 1) {
        return numError()
      }
      const result = periods * ((1 + effectiveRate) ** (1 / periods) - 1)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    PDURATION: (rateArg, presentArg, futureArg) => {
      const rate = toNumber(rateArg)
      const present = toNumber(presentArg)
      const future = toNumber(futureArg)
      if (rate === undefined || present === undefined || future === undefined) {
        return valueError()
      }
      if (!Number.isFinite(rate) || !Number.isFinite(present) || !Number.isFinite(future) || rate <= 0 || present <= 0 || future <= 0) {
        return numError()
      }
      const result = Math.log(future / present) / Math.log(1 + rate)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    RRI: (periodsArg, presentArg, futureArg) => {
      const periods = toNumber(periodsArg)
      const present = toNumber(presentArg)
      const future = toNumber(futureArg)
      if (periods === undefined || present === undefined || future === undefined) {
        return valueError()
      }
      if (!Number.isFinite(periods) || !Number.isFinite(present) || !Number.isFinite(future) || periods <= 0 || present === 0) {
        return numError()
      }
      const result = (future / present) ** (1 / periods) - 1
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    FV: (rateArg, periodsArg, paymentArg, presentArg, typeArg) => {
      const rate = toNumber(rateArg)
      const periods = toNumber(periodsArg)
      const payment = toNumber(paymentArg)
      const present = coerceNumber(presentArg, 0)
      const type = coercePaymentType(typeArg, 0)
      if (rate === undefined || periods === undefined || payment === undefined || present === undefined || type === undefined) {
        return valueError()
      }
      return numberResult(futureValue(rate, periods, payment, present, type))
    },
    FVSCHEDULE: (principalArg, ...scheduleArgs) => {
      const principal = toNumber(principalArg)
      if (principal === undefined) {
        return valueError()
      }
      let result = principal
      for (const scheduleArg of scheduleArgs) {
        const rate = toNumber(scheduleArg)
        if (rate === undefined) {
          return valueError()
        }
        result *= 1 + rate
      }
      return numberResult(result)
    },
    DB: (costArg, salvageArg, lifeArg, periodArg, monthArg) => {
      const cost = toNumber(costArg)
      const salvage = toNumber(salvageArg)
      const life = toNumber(lifeArg)
      const period = toNumber(periodArg)
      const month = coerceNumber(monthArg, 12)
      if (cost === undefined || salvage === undefined || life === undefined || period === undefined || month === undefined) {
        return valueError()
      }
      const depreciation = dbDepreciation(cost, salvage, life, period, month)
      return depreciation === undefined ? valueError() : numberResult(depreciation)
    },
    DDB: (costArg, salvageArg, lifeArg, periodArg, factorArg) => {
      const cost = toNumber(costArg)
      const salvage = toNumber(salvageArg)
      const life = toNumber(lifeArg)
      const period = toNumber(periodArg)
      const factor = factorArg === undefined ? 2 : toNumber(factorArg)
      if (cost === undefined || salvage === undefined || life === undefined || period === undefined || factor === undefined) {
        return valueError()
      }
      if (
        !Number.isFinite(cost) ||
        !Number.isFinite(salvage) ||
        !Number.isFinite(life) ||
        !Number.isFinite(period) ||
        !Number.isFinite(factor) ||
        cost <= 0 ||
        salvage < 0 ||
        life <= 0 ||
        period <= 0 ||
        factor <= 0
      ) {
        return numError()
      }
      const depreciation = ddbDepreciation(cost, salvage, life, period, factor)
      return depreciation === undefined ? numError() : numberResult(depreciation)
    },
    VDB: (costArg, salvageArg, lifeArg, startArg, endArg, factorArg, noSwitchArg) => {
      const cost = toNumber(costArg)
      const salvage = toNumber(salvageArg)
      const life = toNumber(lifeArg)
      const start = toNumber(startArg)
      const end = toNumber(endArg)
      const factor = factorArg === undefined ? 2 : toNumber(factorArg)
      const noSwitch = coerceBoolean(noSwitchArg, false)
      if (
        cost === undefined ||
        salvage === undefined ||
        life === undefined ||
        start === undefined ||
        end === undefined ||
        factor === undefined ||
        noSwitch === undefined
      ) {
        return valueError()
      }
      if (
        !Number.isFinite(cost) ||
        !Number.isFinite(salvage) ||
        !Number.isFinite(life) ||
        !Number.isFinite(start) ||
        !Number.isFinite(end) ||
        !Number.isFinite(factor) ||
        cost <= 0 ||
        salvage < 0 ||
        life <= 0 ||
        start < 0 ||
        end < start ||
        factor <= 0
      ) {
        return numError()
      }
      const depreciation = vdbDepreciation(cost, salvage, life, start, end, factor, noSwitch)
      return depreciation === undefined ? numError() : numberResult(depreciation)
    },
    PV: (rateArg, periodsArg, paymentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg)
      const periods = toNumber(periodsArg)
      const payment = toNumber(paymentArg)
      const future = coerceNumber(futureArg, 0)
      const type = coercePaymentType(typeArg, 0)
      if (rate === undefined || periods === undefined || payment === undefined || future === undefined || type === undefined) {
        return valueError()
      }
      return numberResult(presentValue(rate, periods, payment, future, type))
    },
    PMT: (rateArg, periodsArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg)
      const periods = toNumber(periodsArg)
      const present = toNumber(presentArg)
      const future = coerceNumber(futureArg, 0)
      const type = coercePaymentType(typeArg, 0)
      if (rate === undefined || periods === undefined || present === undefined || future === undefined || type === undefined) {
        return valueError()
      }
      const payment = periodicPayment(rate, periods, present, future, type)
      return payment === undefined ? valueError() : numberResult(payment)
    },
    RATE: (periodsArg, paymentArg, presentArg, futureArg, typeArg, guessArg) => {
      const periods = toNumber(periodsArg)
      const payment = toNumber(paymentArg)
      const present = toNumber(presentArg)
      const future = coerceNumber(futureArg, 0)
      const type = coercePaymentType(typeArg, 0)
      const guess = coerceNumber(guessArg, 0.1)
      if (
        periods === undefined ||
        payment === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined ||
        guess === undefined
      ) {
        return valueError()
      }
      const rate = solveRate(periods, payment, present, future, type, guess)
      return rate === undefined ? valueError() : numberResult(rate)
    },
    SLN: (costArg, salvageArg, lifeArg) => {
      const cost = toNumber(costArg)
      const salvage = toNumber(salvageArg)
      const life = toNumber(lifeArg)
      if (cost === undefined || salvage === undefined || life === undefined || life <= 0) {
        return valueError()
      }
      return numberResult((cost - salvage) / life)
    },
    SYD: (costArg, salvageArg, lifeArg, periodArg) => {
      const cost = toNumber(costArg)
      const salvage = toNumber(salvageArg)
      const life = toNumber(lifeArg)
      const period = toNumber(periodArg)
      if (
        cost === undefined ||
        salvage === undefined ||
        life === undefined ||
        period === undefined ||
        life <= 0 ||
        period <= 0 ||
        period > life
      ) {
        return valueError()
      }
      return numberResult(((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1)))
    },
    NPER: (rateArg, paymentArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg)
      const payment = toNumber(paymentArg)
      const present = toNumber(presentArg)
      const future = coerceNumber(futureArg, 0)
      const type = coercePaymentType(typeArg, 0)
      if (rate === undefined || payment === undefined || present === undefined || future === undefined || type === undefined) {
        return valueError()
      }
      const periods = totalPeriods(rate, payment, present, future, type)
      return periods === undefined ? valueError() : numberResult(periods)
    },
    NPV: (rateArg, ...valueArgs) => {
      const rate = toNumber(rateArg)
      if (rate === undefined || valueArgs.length === 0) {
        return valueError()
      }
      const values: number[] = []
      for (const valueArg of valueArgs) {
        const value = toNumber(valueArg)
        if (value === undefined) {
          return valueError()
        }
        values.push(value)
      }
      if (rate === -1) {
        return div0Error()
      }
      let result = 0
      for (let index = 0; index < values.length; index += 1) {
        result += values[index]! / (1 + rate) ** (index + 1)
      }
      return Number.isFinite(result) ? numberResult(result) : valueError()
    },
    IPMT: (rateArg, periodArg, periodsArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg)
      const period = toNumber(periodArg)
      const periods = toNumber(periodsArg)
      const present = toNumber(presentArg)
      const future = coerceNumber(futureArg, 0)
      const type = coercePaymentType(typeArg, 0)
      if (
        rate === undefined ||
        period === undefined ||
        periods === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined
      ) {
        return valueError()
      }
      const interest = interestPayment(rate, period, periods, present, future, type)
      return interest === undefined ? valueError() : numberResult(interest)
    },
    PPMT: (rateArg, periodArg, periodsArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg)
      const period = toNumber(periodArg)
      const periods = toNumber(periodsArg)
      const present = toNumber(presentArg)
      const future = coerceNumber(futureArg, 0)
      const type = coercePaymentType(typeArg, 0)
      if (
        rate === undefined ||
        period === undefined ||
        periods === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined
      ) {
        return valueError()
      }
      const principal = principalPayment(rate, period, periods, present, future, type)
      return principal === undefined ? valueError() : numberResult(principal)
    },
    ISPMT: (rateArg, periodArg, periodsArg, presentArg) => {
      const rate = toNumber(rateArg)
      const period = toNumber(periodArg)
      const periods = toNumber(periodsArg)
      const present = toNumber(presentArg)
      if (
        rate === undefined ||
        period === undefined ||
        periods === undefined ||
        present === undefined ||
        periods <= 0 ||
        period < 0 ||
        period > periods
      ) {
        return valueError()
      }
      return numberResult(present * rate * (period / periods - 1))
    },
    CUMIPMT: (rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg) => {
      return cumulativePeriodicPaymentResult(rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg, false)
    },
    CUMPRINC: (rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg) => {
      return cumulativePeriodicPaymentResult(rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg, true)
    },
  }
}
