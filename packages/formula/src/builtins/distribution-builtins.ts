import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { CellValue } from '@bilig/protocol'
import {
  betaDistributionCdf,
  betaDistributionDensity,
  betaDistributionInverse,
  binomialProbability,
  chiSquareCdf,
  chiSquareDensity,
  fDistributionCdf,
  fDistributionDensity,
  gammaDistributionCdf,
  gammaDistributionDensity,
  gammaFunction,
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseGammaDistribution,
  inverseStandardNormal,
  inverseStudentT,
  logGamma,
  negativeBinomialProbability,
  poissonProbability,
  regularizedUpperGamma,
  studentTCdf,
  studentTDensity,
} from './distributions.js'
import { erfApprox } from './statistics.js'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

interface DistributionBuiltinDeps {
  toNumber: (value: CellValue) => number | undefined
  coerceBoolean: (value: CellValue | undefined, fallback: boolean) => boolean | undefined
  coerceNumber: (value: CellValue | undefined, fallback: number) => number | undefined
  nonNegativeIntegerValue: (value: CellValue | undefined, fallback?: number) => number | undefined
  positiveIntegerValue: (value: CellValue | undefined, fallback?: number) => number | undefined
  numberResult: (value: number) => EvaluationResult
  numericResultOrError: (value: number) => EvaluationResult
  valueError: () => EvaluationResult
  numError: () => EvaluationResult
}

function naError(): EvaluationResult {
  return { tag: ValueTag.Error, code: ErrorCode.NA }
}

function div0Error(): EvaluationResult {
  return { tag: ValueTag.Error, code: ErrorCode.Div0 }
}

function hasNonFinite(...values: number[]): boolean {
  return values.some((value) => !Number.isFinite(value))
}

function isInvalidHypergeometricDomain(
  sampleSuccesses: number,
  sampleSize: number,
  populationSuccesses: number,
  populationSize: number,
  sampleSuccessesRaw: number,
  sampleSizeRaw: number,
  populationSuccessesRaw: number,
  populationSizeRaw: number,
): boolean {
  const minimumSampleSuccesses = Math.max(0, sampleSize - populationSize + populationSuccesses)
  const maximumSampleSuccesses = Math.min(sampleSize, populationSuccesses)
  return (
    hasNonFinite(sampleSuccessesRaw, sampleSizeRaw, populationSuccessesRaw, populationSizeRaw) ||
    sampleSuccesses < minimumSampleSuccesses ||
    sampleSuccesses > maximumSampleSuccesses ||
    sampleSize <= 0 ||
    sampleSize > populationSize ||
    populationSuccesses <= 0 ||
    populationSuccesses > populationSize ||
    populationSize <= 0
  )
}

export function createDistributionBuiltins({
  toNumber,
  coerceBoolean,
  numberResult,
  numericResultOrError,
  valueError,
  numError,
}: DistributionBuiltinDeps): Record<string, Builtin> {
  const requiredNumber = (value: CellValue | undefined): number | undefined => (value === undefined ? undefined : toNumber(value))
  const optionalNumber = (value: CellValue | undefined, fallback: number): number | undefined =>
    value === undefined ? fallback : toNumber(value)
  const requiredInteger = (value: CellValue | undefined): number | undefined => {
    const numeric = requiredNumber(value)
    return numeric === undefined ? undefined : Math.trunc(numeric)
  }

  const builtins: Record<string, Builtin> = {
    'CONFIDENCE.NORM': (alphaArg, standardDeviationArg, sizeArg) => {
      const alpha = toNumber(alphaArg)
      const standardDeviation = toNumber(standardDeviationArg)
      const sizeRaw = toNumber(sizeArg)
      if (alpha === undefined || standardDeviation === undefined || sizeRaw === undefined) {
        return valueError()
      }
      const size = Math.trunc(sizeRaw)
      if (hasNonFinite(alpha, standardDeviation, sizeRaw) || alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || size < 1) {
        return numError()
      }
      const criticalValue = inverseStandardNormal(1 - alpha / 2)
      return criticalValue === undefined ? valueError() : numberResult((criticalValue * standardDeviation) / Math.sqrt(size))
    },
    ERF: (lowerArg, upperArg) => {
      const lower = toNumber(lowerArg)
      if (lower === undefined) {
        return valueError()
      }
      if (upperArg === undefined) {
        return numberResult(erfApprox(lower))
      }
      const upper = toNumber(upperArg)
      return upper === undefined ? valueError() : numberResult(erfApprox(upper) - erfApprox(lower))
    },
    'ERF.PRECISE': (valueArg) => {
      const value = toNumber(valueArg)
      return value === undefined ? valueError() : numberResult(erfApprox(value))
    },
    ERFC: (valueArg) => {
      const value = toNumber(valueArg)
      return value === undefined ? valueError() : numberResult(1 - erfApprox(value))
    },
    'ERFC.PRECISE': (valueArg) => {
      const value = toNumber(valueArg)
      return value === undefined ? valueError() : numberResult(1 - erfApprox(value))
    },
    FISHER: (valueArg) => {
      const value = toNumber(valueArg)
      if (value === undefined) {
        return valueError()
      }
      if (hasNonFinite(value) || value <= -1 || value >= 1) {
        return numError()
      }
      return numberResult(0.5 * Math.log((1 + value) / (1 - value)))
    },
    FISHERINV: (valueArg) => {
      const value = toNumber(valueArg)
      if (value === undefined) {
        return valueError()
      }
      const exponent = Math.exp(2 * value)
      return numberResult((exponent - 1) / (exponent + 1))
    },
    GAMMALN: (valueArg) => {
      const value = toNumber(valueArg)
      if (value === undefined) {
        return valueError()
      }
      if (hasNonFinite(value) || value <= 0) {
        return numError()
      }
      const result = logGamma(value)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    'GAMMALN.PRECISE': (valueArg) => builtins['GAMMALN']!(valueArg),
    GAMMA: (valueArg) => {
      const value = toNumber(valueArg)
      if (value === undefined) {
        return valueError()
      }
      if (hasNonFinite(value) || (value <= 0 && Number.isInteger(value))) {
        return numError()
      }
      const result = gammaFunction(value)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    CONFIDENCE: (alphaArg, standardDeviationArg, sizeArg) => builtins['CONFIDENCE.NORM']!(alphaArg, standardDeviationArg, sizeArg),
    'CONFIDENCE.T': (alphaArg, standardDeviationArg, sizeArg) => {
      const alpha = toNumber(alphaArg)
      const standardDeviation = toNumber(standardDeviationArg)
      const sizeRaw = toNumber(sizeArg)
      if (alpha === undefined || standardDeviation === undefined || sizeRaw === undefined) {
        return valueError()
      }
      const size = Math.trunc(sizeRaw)
      if (hasNonFinite(alpha, standardDeviation, sizeRaw) || alpha <= 0 || alpha >= 1 || standardDeviation <= 0 || size < 1) {
        return numError()
      }
      if (size === 1) {
        return div0Error()
      }
      const critical = inverseStudentT(1 - alpha / 2, size - 1)
      return critical === undefined ? valueError() : numericResultOrError((critical * standardDeviation) / Math.sqrt(size))
    },
    'BETA.DIST': (xArg, alphaArg, betaArg, cumulativeArg, lowerBoundArg, upperBoundArg) => {
      const x = toNumber(xArg)
      const alpha = toNumber(alphaArg)
      const beta = toNumber(betaArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      const lowerBound = optionalNumber(lowerBoundArg, 0)
      const upperBound = optionalNumber(upperBoundArg, 1)
      if (
        x === undefined ||
        alpha === undefined ||
        beta === undefined ||
        cumulative === undefined ||
        lowerBound === undefined ||
        upperBound === undefined
      ) {
        return valueError()
      }
      if (
        hasNonFinite(x, alpha, beta, lowerBound, upperBound) ||
        alpha <= 0 ||
        beta <= 0 ||
        x < lowerBound ||
        x > upperBound ||
        lowerBound >= upperBound
      ) {
        return numError()
      }
      return numericResultOrError(
        cumulative
          ? betaDistributionCdf(x, alpha, beta, lowerBound, upperBound)
          : betaDistributionDensity(x, alpha, beta, lowerBound, upperBound),
      )
    },
    BETADIST: (xArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) =>
      builtins['BETA.DIST']!(xArg, alphaArg, betaArg, { tag: ValueTag.Boolean, value: true }, lowerBoundArg, upperBoundArg),
    'BETA.INV': (probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) => {
      const probability = toNumber(probabilityArg)
      const alpha = toNumber(alphaArg)
      const beta = toNumber(betaArg)
      const lowerBound = optionalNumber(lowerBoundArg, 0)
      const upperBound = optionalNumber(upperBoundArg, 1)
      if (probability === undefined || alpha === undefined || beta === undefined || lowerBound === undefined || upperBound === undefined) {
        return valueError()
      }
      if (
        hasNonFinite(probability, alpha, beta, lowerBound, upperBound) ||
        probability <= 0 ||
        probability > 1 ||
        alpha <= 0 ||
        beta <= 0 ||
        lowerBound >= upperBound
      ) {
        return numError()
      }
      if (probability === 1) {
        return numberResult(upperBound)
      }
      const result = betaDistributionInverse(probability, alpha, beta, lowerBound, upperBound)
      return result === undefined ? valueError() : numericResultOrError(result)
    },
    BETAINV: (probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) =>
      builtins['BETA.INV']!(probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg),
    EXPONDIST: (xArg, lambdaArg, cumulativeArg) => {
      const x = toNumber(xArg)
      const lambda = toNumber(lambdaArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (x === undefined || lambda === undefined || cumulative === undefined) {
        return valueError()
      }
      if (!Number.isFinite(x) || !Number.isFinite(lambda) || x < 0 || lambda <= 0) {
        return numError()
      }
      return numberResult(cumulative ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x))
    },
    'EXPON.DIST': (xArg, lambdaArg, cumulativeArg) => builtins['EXPONDIST']!(xArg, lambdaArg, cumulativeArg),
    POISSON: (eventsArg, meanArg, cumulativeArg) => {
      const eventsRaw = toNumber(eventsArg)
      const mean = toNumber(meanArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (eventsRaw === undefined || mean === undefined || cumulative === undefined) {
        return valueError()
      }
      const events = Math.trunc(eventsRaw)
      if (!Number.isFinite(eventsRaw) || !Number.isFinite(mean) || events < 0 || mean < 0) {
        return numError()
      }
      if (!cumulative) {
        return numericResultOrError(poissonProbability(events, mean))
      }
      let total = 0
      for (let index = 0; index <= events; index += 1) {
        total += poissonProbability(index, mean)
      }
      return numericResultOrError(total)
    },
    'POISSON.DIST': (eventsArg, meanArg, cumulativeArg) => builtins['POISSON']!(eventsArg, meanArg, cumulativeArg),
    WEIBULL: (xArg, alphaArg, betaArg, cumulativeArg) => {
      const x = toNumber(xArg)
      const alpha = toNumber(alphaArg)
      const beta = toNumber(betaArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (x === undefined || alpha === undefined || beta === undefined || cumulative === undefined) {
        return valueError()
      }
      if (!Number.isFinite(x) || !Number.isFinite(alpha) || !Number.isFinite(beta) || x < 0 || alpha <= 0 || beta <= 0) {
        return numError()
      }
      if (cumulative) {
        return numberResult(1 - Math.exp(-((x / beta) ** alpha)))
      }
      if (x === 0) {
        return numberResult(alpha === 1 ? 1 / beta : alpha < 1 ? Number.POSITIVE_INFINITY : 0)
      }
      return numberResult((alpha / beta ** alpha) * x ** (alpha - 1) * Math.exp(-((x / beta) ** alpha)))
    },
    'WEIBULL.DIST': (xArg, alphaArg, betaArg, cumulativeArg) => builtins['WEIBULL']!(xArg, alphaArg, betaArg, cumulativeArg),
    GAMMADIST: (xArg, alphaArg, betaArg, cumulativeArg) => {
      const x = toNumber(xArg)
      const alpha = toNumber(alphaArg)
      const beta = toNumber(betaArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (x === undefined || alpha === undefined || beta === undefined || cumulative === undefined) {
        return valueError()
      }
      if (!Number.isFinite(x) || !Number.isFinite(alpha) || !Number.isFinite(beta) || x < 0 || alpha <= 0 || beta <= 0) {
        return numError()
      }
      return numberResult(cumulative ? gammaDistributionCdf(x, alpha, beta) : gammaDistributionDensity(x, alpha, beta))
    },
    'GAMMA.DIST': (xArg, alphaArg, betaArg, cumulativeArg) => builtins['GAMMADIST']!(xArg, alphaArg, betaArg, cumulativeArg),
    'GAMMA.INV': (probabilityArg, alphaArg, betaArg) => {
      const probability = toNumber(probabilityArg)
      const alpha = toNumber(alphaArg)
      const beta = toNumber(betaArg)
      if (probability === undefined || alpha === undefined || beta === undefined) {
        return valueError()
      }
      if (
        !Number.isFinite(probability) ||
        !Number.isFinite(alpha) ||
        !Number.isFinite(beta) ||
        probability < 0 ||
        probability >= 1 ||
        alpha <= 0 ||
        beta <= 0
      ) {
        return numError()
      }
      if (probability === 0) {
        return numberResult(0)
      }
      const result = inverseGammaDistribution(probability, alpha, beta)
      return result === undefined ? naError() : numericResultOrError(result)
    },
    GAMMAINV: (probabilityArg, alphaArg, betaArg) => builtins['GAMMA.INV']!(probabilityArg, alphaArg, betaArg),
    CHIDIST: (xArg, degreesArg) => {
      const x = toNumber(xArg)
      const degreesRaw = toNumber(degreesArg)
      if (x === undefined || degreesRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(x, degreesRaw) || x < 0 || degrees < 1 || degrees > 1e10) {
        return numError()
      }
      return numericResultOrError(regularizedUpperGamma(degrees / 2, x / 2))
    },
    'LEGACY.CHIDIST': (xArg, degreesArg) => builtins['CHIDIST']!(xArg, degreesArg),
    CHIINV: (probabilityArg, degreesArg) => builtins['CHISQ.INV.RT']!(probabilityArg, degreesArg),
    'CHISQ.DIST.RT': (xArg, degreesArg) => builtins['CHIDIST']!(xArg, degreesArg),
    CHISQDIST: (xArg, degreesArg) => builtins['CHISQ.DIST.RT']!(xArg, degreesArg),
    'CHISQ.DIST': (xArg, degreesArg, cumulativeArg) => {
      const x = toNumber(xArg)
      const degreesRaw = toNumber(degreesArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (x === undefined || degreesRaw === undefined || cumulative === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(x, degreesRaw) || x < 0 || degrees < 1 || degrees > 1e10) {
        return numError()
      }
      return numberResult(cumulative ? chiSquareCdf(x, degrees) : chiSquareDensity(x, degrees))
    },
    'CHISQ.INV.RT': (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg)
      const degreesRaw = toNumber(degreesArg)
      if (probability === undefined || degreesRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(probability, degreesRaw) || probability < 0 || probability > 1 || degrees < 1 || degrees > 1e10) {
        return numError()
      }
      const result = inverseChiSquare(1 - probability, degrees)
      return result === undefined ? naError() : numberResult(result)
    },
    CHISQINV: (probabilityArg, degreesArg) => builtins['CHISQ.INV.RT']!(probabilityArg, degreesArg),
    'LEGACY.CHIINV': (probabilityArg, degreesArg) => builtins['CHISQ.INV.RT']!(probabilityArg, degreesArg),
    'CHISQ.INV': (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg)
      const degreesRaw = toNumber(degreesArg)
      if (probability === undefined || degreesRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(probability, degreesRaw) || probability < 0 || probability > 1 || degrees < 1 || degrees > 1e10) {
        return numError()
      }
      const result = inverseChiSquare(probability, degrees)
      return result === undefined ? naError() : numberResult(result)
    },
    'F.DIST': (xArg, degrees1Arg, degrees2Arg, cumulativeArg) => {
      const x = toNumber(xArg)
      const degrees1Raw = toNumber(degrees1Arg)
      const degrees2Raw = toNumber(degrees2Arg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (x === undefined || degrees1Raw === undefined || degrees2Raw === undefined || cumulative === undefined) {
        return valueError()
      }
      const degrees1 = Math.trunc(degrees1Raw)
      const degrees2 = Math.trunc(degrees2Raw)
      if (hasNonFinite(x, degrees1Raw, degrees2Raw) || x < 0 || degrees1 < 1 || degrees2 < 1) {
        return numError()
      }
      return numericResultOrError(cumulative ? fDistributionCdf(x, degrees1, degrees2) : fDistributionDensity(x, degrees1, degrees2))
    },
    'F.DIST.RT': (xArg, degrees1Arg, degrees2Arg) => {
      const x = toNumber(xArg)
      const degrees1Raw = toNumber(degrees1Arg)
      const degrees2Raw = toNumber(degrees2Arg)
      if (x === undefined || degrees1Raw === undefined || degrees2Raw === undefined) {
        return valueError()
      }
      const degrees1 = Math.trunc(degrees1Raw)
      const degrees2 = Math.trunc(degrees2Raw)
      if (hasNonFinite(x, degrees1Raw, degrees2Raw) || x < 0 || degrees1 < 1 || degrees2 < 1) {
        return numError()
      }
      return numericResultOrError(1 - fDistributionCdf(x, degrees1, degrees2))
    },
    FDIST: (xArg, degrees1Arg, degrees2Arg) => builtins['F.DIST.RT']!(xArg, degrees1Arg, degrees2Arg),
    'LEGACY.FDIST': (xArg, degrees1Arg, degrees2Arg) => builtins['F.DIST.RT']!(xArg, degrees1Arg, degrees2Arg),
    'F.INV': (probabilityArg, degrees1Arg, degrees2Arg) => {
      const probability = toNumber(probabilityArg)
      const degrees1Raw = toNumber(degrees1Arg)
      const degrees2Raw = toNumber(degrees2Arg)
      if (probability === undefined || degrees1Raw === undefined || degrees2Raw === undefined) {
        return valueError()
      }
      const degrees1 = Math.trunc(degrees1Raw)
      const degrees2 = Math.trunc(degrees2Raw)
      if (hasNonFinite(probability, degrees1Raw, degrees2Raw) || probability < 0 || probability > 1 || degrees1 < 1 || degrees2 < 1) {
        return numError()
      }
      const result = inverseFDistribution(probability, degrees1, degrees2)
      return result === undefined ? valueError() : numericResultOrError(result)
    },
    'F.INV.RT': (probabilityArg, degrees1Arg, degrees2Arg) => {
      const probability = toNumber(probabilityArg)
      const degrees1Raw = toNumber(degrees1Arg)
      const degrees2Raw = toNumber(degrees2Arg)
      if (probability === undefined || degrees1Raw === undefined || degrees2Raw === undefined) {
        return valueError()
      }
      const degrees1 = Math.trunc(degrees1Raw)
      const degrees2 = Math.trunc(degrees2Raw)
      if (hasNonFinite(probability, degrees1Raw, degrees2Raw) || probability < 0 || probability > 1 || degrees1 < 1 || degrees2 < 1) {
        return numError()
      }
      const result = inverseFDistribution(1 - probability, degrees1, degrees2)
      return result === undefined ? valueError() : numericResultOrError(result)
    },
    FINV: (probabilityArg, degrees1Arg, degrees2Arg) => builtins['F.INV.RT']!(probabilityArg, degrees1Arg, degrees2Arg),
    'LEGACY.FINV': (probabilityArg, degrees1Arg, degrees2Arg) => builtins['F.INV.RT']!(probabilityArg, degrees1Arg, degrees2Arg),
    'T.DIST': (xArg, degreesArg, cumulativeArg) => {
      const x = toNumber(xArg)
      const degreesRaw = toNumber(degreesArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (x === undefined || degreesRaw === undefined || cumulative === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(x, degreesRaw) || degrees < 1) {
        return numError()
      }
      return numericResultOrError(cumulative ? studentTCdf(x, degrees) : studentTDensity(x, degrees))
    },
    'T.DIST.RT': (xArg, degreesArg) => {
      const x = toNumber(xArg)
      const degreesRaw = toNumber(degreesArg)
      if (x === undefined || degreesRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(x, degreesRaw) || degrees < 1) {
        return numError()
      }
      return numericResultOrError(1 - studentTCdf(x, degrees))
    },
    'T.DIST.2T': (xArg, degreesArg) => {
      const x = toNumber(xArg)
      const degreesRaw = toNumber(degreesArg)
      if (x === undefined || degreesRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(x, degreesRaw) || x < 0 || degrees < 1) {
        return numError()
      }
      return numericResultOrError(Math.min(1, 2 * (1 - studentTCdf(x, degrees))))
    },
    TDIST: (xArg, degreesArg, tailsArg) => {
      const x = toNumber(xArg)
      const degreesRaw = toNumber(degreesArg)
      const tailsRaw = toNumber(tailsArg)
      if (x === undefined || degreesRaw === undefined || tailsRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      const tails = Math.trunc(tailsRaw)
      if (hasNonFinite(x, degreesRaw, tailsRaw) || x < 0 || degrees < 1 || (tails !== 1 && tails !== 2)) {
        return numError()
      }
      const upperTail = 1 - studentTCdf(x, degrees)
      return numericResultOrError(tails === 1 ? upperTail : Math.min(1, upperTail * 2))
    },
    'T.INV': (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg)
      const degreesRaw = toNumber(degreesArg)
      if (probability === undefined || degreesRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(probability, degreesRaw) || probability <= 0 || probability >= 1 || degrees < 1) {
        return numError()
      }
      const result = inverseStudentT(probability, degrees)
      return result === undefined ? numError() : numericResultOrError(result)
    },
    'T.INV.2T': (probabilityArg, degreesArg) => {
      const probability = toNumber(probabilityArg)
      const degreesRaw = toNumber(degreesArg)
      if (probability === undefined || degreesRaw === undefined) {
        return valueError()
      }
      const degrees = Math.trunc(degreesRaw)
      if (hasNonFinite(probability, degreesRaw) || probability <= 0 || probability > 1 || degrees < 1) {
        return numError()
      }
      const result = inverseStudentT(1 - probability / 2, degrees)
      return result === undefined ? numError() : numericResultOrError(result)
    },
    TINV: (probabilityArg, degreesArg) => builtins['T.INV.2T']!(probabilityArg, degreesArg),
    BINOMDIST: (successesArg, trialsArg, probabilityArg, cumulativeArg) => {
      const successes = requiredInteger(successesArg)
      const trials = requiredInteger(trialsArg)
      const successesRaw = requiredNumber(successesArg)
      const trialsRaw = requiredNumber(trialsArg)
      const probability = toNumber(probabilityArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (
        successes === undefined ||
        trials === undefined ||
        successesRaw === undefined ||
        trialsRaw === undefined ||
        probability === undefined ||
        cumulative === undefined
      ) {
        return valueError()
      }
      if (
        hasNonFinite(successesRaw, trialsRaw, probability) ||
        successes < 0 ||
        trials < 0 ||
        successes > trials ||
        probability < 0 ||
        probability > 1
      ) {
        return numError()
      }
      if (!cumulative) {
        return numericResultOrError(binomialProbability(successes, trials, probability))
      }
      let total = 0
      for (let index = 0; index <= successes; index += 1) {
        total += binomialProbability(index, trials, probability)
      }
      return numericResultOrError(total)
    },
    'BINOM.DIST': (successesArg, trialsArg, probabilityArg, cumulativeArg) =>
      builtins['BINOMDIST']!(successesArg, trialsArg, probabilityArg, cumulativeArg),
    'BINOM.DIST.RANGE': (trialsArg, probabilityArg, successesArg, upperSuccessesArg) => {
      const trials = requiredInteger(trialsArg)
      const trialsRaw = requiredNumber(trialsArg)
      const probability = toNumber(probabilityArg)
      const lower = requiredInteger(successesArg)
      const lowerRaw = requiredNumber(successesArg)
      const upper = upperSuccessesArg === undefined ? lower : requiredInteger(upperSuccessesArg)
      const upperRaw = upperSuccessesArg === undefined ? lowerRaw : requiredNumber(upperSuccessesArg)
      if (
        trials === undefined ||
        trialsRaw === undefined ||
        probability === undefined ||
        lower === undefined ||
        lowerRaw === undefined ||
        upper === undefined ||
        upperRaw === undefined
      ) {
        return valueError()
      }
      if (
        hasNonFinite(trialsRaw, probability, lowerRaw, upperRaw) ||
        trials < 0 ||
        lower < 0 ||
        upper < 0 ||
        lower > upper ||
        upper > trials ||
        probability < 0 ||
        probability > 1
      ) {
        return numError()
      }
      let total = 0
      for (let index = lower; index <= upper; index += 1) {
        total += binomialProbability(index, trials, probability)
      }
      return numericResultOrError(total)
    },
    CRITBINOM: (trialsArg, probabilityArg, alphaArg) => {
      const trials = requiredInteger(trialsArg)
      const trialsRaw = requiredNumber(trialsArg)
      const probability = toNumber(probabilityArg)
      const alpha = toNumber(alphaArg)
      if (trials === undefined || trialsRaw === undefined || probability === undefined || alpha === undefined) {
        return valueError()
      }
      if (hasNonFinite(trialsRaw, probability, alpha) || trials < 0 || probability <= 0 || probability >= 1 || alpha <= 0 || alpha >= 1) {
        return numError()
      }
      let cumulative = 0
      for (let index = 0; index <= trials; index += 1) {
        cumulative += binomialProbability(index, trials, probability)
        if (cumulative >= alpha) {
          return numberResult(index)
        }
      }
      return numberResult(trials)
    },
    'BINOM.INV': (trialsArg, probabilityArg, alphaArg) => builtins['CRITBINOM']!(trialsArg, probabilityArg, alphaArg),
    HYPGEOMDIST: (sampleSuccessesArg, sampleSizeArg, populationSuccessesArg, populationSizeArg) => {
      const sampleSuccesses = requiredInteger(sampleSuccessesArg)
      const sampleSize = requiredInteger(sampleSizeArg)
      const populationSuccesses = requiredInteger(populationSuccessesArg)
      const populationSize = requiredInteger(populationSizeArg)
      const sampleSuccessesRaw = requiredNumber(sampleSuccessesArg)
      const sampleSizeRaw = requiredNumber(sampleSizeArg)
      const populationSuccessesRaw = requiredNumber(populationSuccessesArg)
      const populationSizeRaw = requiredNumber(populationSizeArg)
      if (
        sampleSuccesses === undefined ||
        sampleSize === undefined ||
        populationSuccesses === undefined ||
        populationSize === undefined ||
        sampleSuccessesRaw === undefined ||
        sampleSizeRaw === undefined ||
        populationSuccessesRaw === undefined ||
        populationSizeRaw === undefined
      ) {
        return valueError()
      }
      if (
        isInvalidHypergeometricDomain(
          sampleSuccesses,
          sampleSize,
          populationSuccesses,
          populationSize,
          sampleSuccessesRaw,
          sampleSizeRaw,
          populationSuccessesRaw,
          populationSizeRaw,
        )
      ) {
        return numError()
      }
      return numericResultOrError(hypergeometricProbability(sampleSuccesses, sampleSize, populationSuccesses, populationSize))
    },
    'HYPGEOM.DIST': (sampleSuccessesArg, sampleSizeArg, populationSuccessesArg, populationSizeArg, cumulativeArg) => {
      const sampleSuccesses = requiredInteger(sampleSuccessesArg)
      const sampleSize = requiredInteger(sampleSizeArg)
      const populationSuccesses = requiredInteger(populationSuccessesArg)
      const populationSize = requiredInteger(populationSizeArg)
      const sampleSuccessesRaw = requiredNumber(sampleSuccessesArg)
      const sampleSizeRaw = requiredNumber(sampleSizeArg)
      const populationSuccessesRaw = requiredNumber(populationSuccessesArg)
      const populationSizeRaw = requiredNumber(populationSizeArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (
        sampleSuccesses === undefined ||
        sampleSize === undefined ||
        populationSuccesses === undefined ||
        populationSize === undefined ||
        sampleSuccessesRaw === undefined ||
        sampleSizeRaw === undefined ||
        populationSuccessesRaw === undefined ||
        populationSizeRaw === undefined ||
        cumulative === undefined
      ) {
        return valueError()
      }
      if (
        isInvalidHypergeometricDomain(
          sampleSuccesses,
          sampleSize,
          populationSuccesses,
          populationSize,
          sampleSuccessesRaw,
          sampleSizeRaw,
          populationSuccessesRaw,
          populationSizeRaw,
        )
      ) {
        return numError()
      }
      if (!cumulative) {
        return numericResultOrError(hypergeometricProbability(sampleSuccesses, sampleSize, populationSuccesses, populationSize))
      }
      const minimum = Math.max(0, sampleSize - (populationSize - populationSuccesses))
      let total = 0
      for (let index = minimum; index <= sampleSuccesses; index += 1) {
        total += hypergeometricProbability(index, sampleSize, populationSuccesses, populationSize)
      }
      return numericResultOrError(total)
    },
    NEGBINOMDIST: (failuresArg, successesArg, probabilityArg) => {
      const failures = requiredInteger(failuresArg)
      const successes = requiredInteger(successesArg)
      const failuresRaw = requiredNumber(failuresArg)
      const successesRaw = requiredNumber(successesArg)
      const probability = toNumber(probabilityArg)
      if (
        failures === undefined ||
        successes === undefined ||
        failuresRaw === undefined ||
        successesRaw === undefined ||
        probability === undefined
      ) {
        return valueError()
      }
      if (hasNonFinite(failuresRaw, successesRaw, probability) || failures < 0 || successes < 1 || probability < 0 || probability > 1) {
        return numError()
      }
      return numericResultOrError(negativeBinomialProbability(failures, successes, probability))
    },
    'NEGBINOM.DIST': (failuresArg, successesArg, probabilityArg, cumulativeArg) => {
      const failures = requiredInteger(failuresArg)
      const successes = requiredInteger(successesArg)
      const failuresRaw = requiredNumber(failuresArg)
      const successesRaw = requiredNumber(successesArg)
      const probability = toNumber(probabilityArg)
      const cumulative = coerceBoolean(cumulativeArg, false)
      if (
        failures === undefined ||
        successes === undefined ||
        failuresRaw === undefined ||
        successesRaw === undefined ||
        probability === undefined ||
        cumulative === undefined
      ) {
        return valueError()
      }
      if (hasNonFinite(failuresRaw, successesRaw, probability) || failures < 0 || successes < 1 || probability < 0 || probability > 1) {
        return numError()
      }
      if (!cumulative) {
        return numericResultOrError(negativeBinomialProbability(failures, successes, probability))
      }
      let total = 0
      for (let index = 0; index <= failures; index += 1) {
        total += negativeBinomialProbability(index, successes, probability)
      }
      return numericResultOrError(total)
    },
  }

  return builtins
}
