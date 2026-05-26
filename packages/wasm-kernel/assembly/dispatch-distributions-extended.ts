import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { toNumberExact } from './operands'
import { isNumericResult, rangeSupportedScalarOnly, scalarErrorAt } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'
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
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseStudentT,
  negativeBinomialProbability,
  poissonProbability,
  regularizedUpperGamma,
  studentTCdf,
  studentTDensity,
} from './distributions'

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

function isInvalidHypergeometricDomain(
  sampleSuccesses: i32,
  sampleSize: i32,
  populationSuccesses: i32,
  populationSize: i32,
  sampleSuccessesRaw: f64,
  sampleSizeRaw: f64,
  populationSuccessesRaw: f64,
  populationSizeRaw: f64,
): bool {
  const minimumSampleSuccesses = max<i32>(0, sampleSize - populationSize + populationSuccesses)
  const maximumSampleSuccesses = min<i32>(sampleSize, populationSuccesses)
  return (
    !isFinite(sampleSuccessesRaw) ||
    !isFinite(sampleSizeRaw) ||
    !isFinite(populationSuccessesRaw) ||
    !isFinite(populationSizeRaw) ||
    sampleSuccesses < minimumSampleSuccesses ||
    sampleSuccesses > maximumSampleSuccesses ||
    sampleSize <= 0 ||
    sampleSize > populationSize ||
    populationSuccesses <= 0 ||
    populationSuccesses > populationSize ||
    populationSize <= 0
  )
}

export function tryApplyExtendedDistributionBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  if (
    (builtinId == BuiltinId.Expondist ||
      builtinId == BuiltinId.ExponDist ||
      builtinId == BuiltinId.Poisson ||
      builtinId == BuiltinId.PoissonDist ||
      builtinId == BuiltinId.Negbinomdist) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = NaN
    let errorCode = ErrorCode.None
    if (builtinId == BuiltinId.Expondist || builtinId == BuiltinId.ExponDist) {
      const x = toNumberExact(tagStack[base], valueStack[base])
      const lambda = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2])
      if (isNaN(x) || isNaN(lambda) || cumulative < 0) {
        errorCode = ErrorCode.Value
      } else if (!isFinite(x) || !isFinite(lambda) || x < 0.0 || lambda <= 0.0) {
        errorCode = ErrorCode.Num
      } else {
        result = cumulative == 1 ? 1.0 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x)
      }
    } else if (builtinId == BuiltinId.Poisson || builtinId == BuiltinId.PoissonDist) {
      const eventsRaw = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2])
      const events = <i32>eventsRaw
      if (isNaN(eventsRaw) || isNaN(mean) || cumulative < 0) {
        errorCode = ErrorCode.Value
      } else if (!isFinite(eventsRaw) || !isFinite(mean) || events < 0 || mean < 0.0) {
        errorCode = ErrorCode.Num
      } else {
        if (cumulative == 1) {
          result = 0.0
          for (let index = 0; index <= events; index += 1) {
            result += poissonProbability(index, mean)
          }
        } else {
          result = poissonProbability(events, mean)
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base])
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const failures = <i32>failuresRaw
      const successes = <i32>successesRaw
      if (isNaN(failuresRaw) || isNaN(successesRaw) || isNaN(probability)) {
        errorCode = ErrorCode.Value
      } else if (
        !isFinite(failuresRaw) ||
        !isFinite(successesRaw) ||
        !isFinite(probability) ||
        failures < 0 ||
        successes < 1 ||
        probability < 0.0 ||
        probability > 1.0
      ) {
        errorCode = ErrorCode.Num
      } else {
        result = negativeBinomialProbability(failures, successes, probability)
      }
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (
    (builtinId == BuiltinId.Weibull ||
      builtinId == BuiltinId.WeibullDist ||
      builtinId == BuiltinId.Gammadist ||
      builtinId == BuiltinId.GammaDist ||
      builtinId == BuiltinId.Binomdist ||
      builtinId == BuiltinId.BinomDist ||
      builtinId == BuiltinId.NegbinomDist) &&
    argc == 4
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = NaN
    let errorCode = ErrorCode.None
    if (builtinId == BuiltinId.Weibull || builtinId == BuiltinId.WeibullDist) {
      const x = toNumberExact(tagStack[base], valueStack[base])
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
      if (isNaN(x) || isNaN(alpha) || isNaN(beta) || cumulative < 0) {
        errorCode = ErrorCode.Value
      } else if (!isFinite(x) || !isFinite(alpha) || !isFinite(beta) || x < 0.0 || alpha <= 0.0 || beta <= 0.0) {
        errorCode = ErrorCode.Num
      } else {
        if (cumulative == 1) {
          result = 1.0 - Math.exp(-Math.pow(x / beta, alpha))
        } else if (x == 0.0) {
          result = alpha == 1.0 ? 1.0 / beta : alpha < 1.0 ? Infinity : 0.0
        } else {
          result = (alpha / Math.pow(beta, alpha)) * Math.pow(x, alpha - 1.0) * Math.exp(-Math.pow(x / beta, alpha))
        }
      }
    } else if (builtinId == BuiltinId.Gammadist || builtinId == BuiltinId.GammaDist) {
      const x = toNumberExact(tagStack[base], valueStack[base])
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
      if (isNaN(x) || isNaN(alpha) || isNaN(beta) || cumulative < 0) {
        errorCode = ErrorCode.Value
      } else if (!isFinite(x) || !isFinite(alpha) || !isFinite(beta) || x < 0.0 || alpha <= 0.0 || beta <= 0.0) {
        errorCode = ErrorCode.Num
      } else {
        result = cumulative == 1 ? gammaDistributionCdf(x, alpha, beta) : gammaDistributionDensity(x, alpha, beta)
      }
    } else if (builtinId == BuiltinId.Binomdist || builtinId == BuiltinId.BinomDist) {
      const successesRaw = toNumberExact(tagStack[base], valueStack[base])
      const trialsRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
      const successes = <i32>successesRaw
      const trials = <i32>trialsRaw
      if (isNaN(successesRaw) || isNaN(trialsRaw) || isNaN(probability) || cumulative < 0) {
        errorCode = ErrorCode.Value
      } else if (
        !isFinite(successesRaw) ||
        !isFinite(trialsRaw) ||
        !isFinite(probability) ||
        successes < 0 ||
        trials < 0 ||
        successes > trials ||
        probability < 0.0 ||
        probability > 1.0
      ) {
        errorCode = ErrorCode.Num
      } else {
        if (cumulative == 1) {
          result = 0.0
          for (let index = 0; index <= successes; index += 1) {
            result += binomialProbability(index, trials, probability)
          }
        } else {
          result = binomialProbability(successes, trials, probability)
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base])
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
      const failures = <i32>failuresRaw
      const successes = <i32>successesRaw
      if (isNaN(failuresRaw) || isNaN(successesRaw) || isNaN(probability) || cumulative < 0) {
        errorCode = ErrorCode.Value
      } else if (
        !isFinite(failuresRaw) ||
        !isFinite(successesRaw) ||
        !isFinite(probability) ||
        failures < 0 ||
        successes < 1 ||
        probability < 0.0 ||
        probability > 1.0
      ) {
        errorCode = ErrorCode.Num
      } else {
        if (cumulative == 1) {
          result = 0.0
          for (let index = 0; index <= failures; index += 1) {
            result += negativeBinomialProbability(index, successes, probability)
          }
        } else {
          result = negativeBinomialProbability(failures, successes, probability)
        }
      }
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (
    (builtinId == BuiltinId.Chidist ||
      builtinId == BuiltinId.LegacyChidist ||
      builtinId == BuiltinId.ChisqDistRt ||
      builtinId == BuiltinId.Chisqdist) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees = <i32>degreesRaw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(degreesRaw)) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(x) || !isFinite(degreesRaw) || x < 0.0 || degrees < 1 || degreesRaw > 1.0e10) {
      errorCode = ErrorCode.Num
    } else {
      result = regularizedUpperGamma(<f64>degrees / 2.0, x / 2.0)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.ChisqDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2])
    const degrees = <i32>degreesRaw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(degreesRaw) || cumulative < 0) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(x) || !isFinite(degreesRaw) || x < 0.0 || degrees < 1 || degreesRaw > 1.0e10) {
      errorCode = ErrorCode.Num
    } else {
      result = cumulative == 1 ? chiSquareCdf(x, <f64>degrees) : chiSquareDensity(x, <f64>degrees)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (
    (builtinId == BuiltinId.Chiinv ||
      builtinId == BuiltinId.ChisqInvRt ||
      builtinId == BuiltinId.Chisqinv ||
      builtinId == BuiltinId.LegacyChiinv) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const probability = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees = <i32>degreesRaw
    let errorCode = ErrorCode.None
    if (isNaN(probability) || isNaN(degreesRaw)) {
      errorCode = ErrorCode.Value
    } else if (
      !isFinite(probability) ||
      !isFinite(degreesRaw) ||
      probability < 0.0 ||
      probability > 1.0 ||
      degrees < 1 ||
      degreesRaw > 1.0e10
    ) {
      errorCode = ErrorCode.Num
    }
    let result = NaN
    if (errorCode == ErrorCode.None && probability > 0.0 && probability < 1.0) {
      result = inverseChiSquare(1.0 - probability, <f64>degrees)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error
    const resultValue = isNumericResult(result) ? result : ErrorCode.NA
    return writeResult(base, STACK_KIND_SCALAR, resultTag, resultValue, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.ChisqInv && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const probability = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees = <i32>degreesRaw
    let errorCode = ErrorCode.None
    if (isNaN(probability) || isNaN(degreesRaw)) {
      errorCode = ErrorCode.Value
    } else if (
      !isFinite(probability) ||
      !isFinite(degreesRaw) ||
      probability < 0.0 ||
      probability > 1.0 ||
      degrees < 1 ||
      degreesRaw > 1.0e10
    ) {
      errorCode = ErrorCode.Num
    }
    let result = NaN
    if (errorCode == ErrorCode.None && probability > 0.0 && probability < 1.0) {
      result = inverseChiSquare(probability, <f64>degrees)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error
    const resultValue = isNumericResult(result) ? result : ErrorCode.NA
    return writeResult(base, STACK_KIND_SCALAR, resultTag, resultValue, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.BetaDist && argc >= 4 && argc <= 6) || (builtinId == BuiltinId.Betadist && argc >= 3 && argc <= 5)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const modern = builtinId == BuiltinId.BetaDist
    const cumulative = modern ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1
    const lowerBound = modern
      ? argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 0.0
      : argc >= 4
        ? toNumberExact(tagStack[base + 3], valueStack[base + 3])
        : 0.0
    const upperBound = modern
      ? argc >= 6
        ? toNumberExact(tagStack[base + 5], valueStack[base + 5])
        : 1.0
      : argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 1.0
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(alpha) || isNaN(beta) || cumulative < 0 || isNaN(lowerBound) || isNaN(upperBound)) {
      errorCode = ErrorCode.Value
    } else if (
      !isFinite(x) ||
      !isFinite(alpha) ||
      !isFinite(beta) ||
      !isFinite(lowerBound) ||
      !isFinite(upperBound) ||
      alpha <= 0.0 ||
      beta <= 0.0 ||
      x < lowerBound ||
      x > upperBound ||
      lowerBound >= upperBound
    ) {
      errorCode = ErrorCode.Num
    } else {
      result =
        cumulative == 1 || !modern
          ? betaDistributionCdf(x, alpha, beta, lowerBound, upperBound)
          : betaDistributionDensity(x, alpha, beta, lowerBound, upperBound)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.BetaInv || builtinId == BuiltinId.Betainv) && argc >= 3 && argc <= 5) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const probability = toNumberExact(tagStack[base], valueStack[base])
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const lowerBound = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0
    const upperBound = argc >= 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 1.0
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(probability) || isNaN(alpha) || isNaN(beta) || isNaN(lowerBound) || isNaN(upperBound)) {
      errorCode = ErrorCode.Value
    } else if (
      !isFinite(probability) ||
      !isFinite(alpha) ||
      !isFinite(beta) ||
      !isFinite(lowerBound) ||
      !isFinite(upperBound) ||
      probability <= 0.0 ||
      probability > 1.0 ||
      alpha <= 0.0 ||
      beta <= 0.0 ||
      lowerBound >= upperBound
    ) {
      errorCode = ErrorCode.Num
    } else if (probability == 1.0) {
      result = upperBound
    } else {
      result = betaDistributionInverse(probability, alpha, beta, lowerBound, upperBound)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.FDist && argc == 4) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
    const degrees1 = <i32>degrees1Raw
    const degrees2 = <i32>degrees2Raw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(degrees1Raw) || isNaN(degrees2Raw) || cumulative < 0) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(x) || !isFinite(degrees1Raw) || !isFinite(degrees2Raw) || x < 0.0 || degrees1 < 1 || degrees2 < 1) {
      errorCode = ErrorCode.Num
    } else {
      result = cumulative == 1 ? fDistributionCdf(x, <f64>degrees1, <f64>degrees2) : fDistributionDensity(x, <f64>degrees1, <f64>degrees2)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.FDistRt || builtinId == BuiltinId.Fdist || builtinId == BuiltinId.LegacyFdist) && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const degrees1 = <i32>degrees1Raw
    const degrees2 = <i32>degrees2Raw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(degrees1Raw) || isNaN(degrees2Raw)) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(x) || !isFinite(degrees1Raw) || !isFinite(degrees2Raw) || x < 0.0 || degrees1 < 1 || degrees2 < 1) {
      errorCode = ErrorCode.Num
    } else {
      const cdf = fDistributionCdf(x, <f64>degrees1, <f64>degrees2)
      result = isFinite(cdf) ? 1.0 - cdf : NaN
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (
    (builtinId == BuiltinId.FInv || builtinId == BuiltinId.FInvRt || builtinId == BuiltinId.Finv || builtinId == BuiltinId.LegacyFinv) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base])
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const degrees1 = <i32>degrees1Raw
    const degrees2 = <i32>degrees2Raw
    const probability = builtinId == BuiltinId.FInv ? probabilityRaw : 1.0 - probabilityRaw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(probabilityRaw) || isNaN(degrees1Raw) || isNaN(degrees2Raw)) {
      errorCode = ErrorCode.Value
    } else if (
      !isFinite(probabilityRaw) ||
      !isFinite(degrees1Raw) ||
      !isFinite(degrees2Raw) ||
      probabilityRaw < 0.0 ||
      probabilityRaw > 1.0 ||
      degrees1 < 1 ||
      degrees2 < 1
    ) {
      errorCode = ErrorCode.Num
    } else {
      result = inverseFDistribution(probability, <f64>degrees1, <f64>degrees2)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.TDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2])
    const degrees = <i32>degreesRaw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(degreesRaw) || cumulative < 0) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(x) || !isFinite(degreesRaw) || degrees < 1) {
      errorCode = ErrorCode.Num
    } else {
      result = cumulative == 1 ? studentTCdf(x, <f64>degrees) : studentTDensity(x, <f64>degrees)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.TDistRt || builtinId == BuiltinId.TDist2T) && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees = <i32>degreesRaw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(degreesRaw)) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(x) || !isFinite(degreesRaw) || degrees < 1 || (builtinId == BuiltinId.TDist2T && x < 0.0)) {
      errorCode = ErrorCode.Num
    } else {
      const upperTail = 1.0 - studentTCdf(x, <f64>degrees)
      result = isFinite(upperTail) ? (builtinId == BuiltinId.TDistRt ? upperTail : min(1.0, upperTail * 2.0)) : NaN
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Tdist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const tailsRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const degrees = <i32>degreesRaw
    const tails = <i32>tailsRaw
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(x) || isNaN(degreesRaw) || isNaN(tailsRaw)) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(x) || !isFinite(degreesRaw) || !isFinite(tailsRaw) || x < 0.0 || degrees < 1 || (tails != 1 && tails != 2)) {
      errorCode = ErrorCode.Num
    } else {
      const upperTail = 1.0 - studentTCdf(x, <f64>degrees)
      result = isFinite(upperTail) ? (tails == 1 ? upperTail : min(1.0, upperTail * 2.0)) : NaN
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.TInv || builtinId == BuiltinId.TInv2T || builtinId == BuiltinId.Tinv) && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base])
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const degrees = <i32>degreesRaw
    const probability = builtinId == BuiltinId.TInv ? probabilityRaw : 1.0 - probabilityRaw / 2.0
    let result = NaN
    let errorCode = ErrorCode.None
    if (isNaN(probabilityRaw) || isNaN(degreesRaw)) {
      errorCode = ErrorCode.Value
    } else if (!isFinite(probabilityRaw) || !isFinite(degreesRaw) || probabilityRaw <= 0.0 || probabilityRaw > 1.0 || degrees < 1) {
      errorCode = ErrorCode.Num
    } else {
      result = inverseStudentT(probability, <f64>degrees)
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (
    (builtinId == BuiltinId.BinomDistRange && (argc == 3 || argc == 4)) ||
    (builtinId == BuiltinId.Critbinom && argc == 3) ||
    (builtinId == BuiltinId.BinomInv && argc == 3) ||
    (builtinId == BuiltinId.Hypgeomdist && argc == 4) ||
    (builtinId == BuiltinId.HypgeomDist && argc == 5)
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = NaN
    let errorCode = ErrorCode.None
    if (builtinId == BuiltinId.BinomDistRange) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base])
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const lowerRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const upperRaw = argc == 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : lowerRaw
      const trials = <i32>trialsRaw
      const lower = <i32>lowerRaw
      const upper = <i32>upperRaw
      if (isNaN(trialsRaw) || isNaN(probability) || isNaN(lowerRaw) || isNaN(upperRaw)) {
        errorCode = ErrorCode.Value
      } else if (
        !isFinite(trialsRaw) ||
        !isFinite(probability) ||
        !isFinite(lowerRaw) ||
        !isFinite(upperRaw) ||
        trials < 0 ||
        lower < 0 ||
        upper < 0 ||
        lower > upper ||
        upper > trials ||
        probability < 0.0 ||
        probability > 1.0
      ) {
        errorCode = ErrorCode.Num
      } else {
        result = 0.0
        for (let index = lower; index <= upper; index += 1) {
          result += binomialProbability(index, trials, probability)
        }
      }
    } else if (builtinId == BuiltinId.Critbinom || builtinId == BuiltinId.BinomInv) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base])
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const alpha = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const trials = <i32>trialsRaw
      if (isNaN(trialsRaw) || isNaN(probability) || isNaN(alpha)) {
        errorCode = ErrorCode.Value
      } else if (
        !isFinite(trialsRaw) ||
        !isFinite(probability) ||
        !isFinite(alpha) ||
        trials < 0 ||
        probability <= 0.0 ||
        probability >= 1.0 ||
        alpha <= 0.0 ||
        alpha >= 1.0
      ) {
        errorCode = ErrorCode.Num
      } else {
        let cumulative = 0.0
        for (let index = 0; index <= trials; index += 1) {
          cumulative += binomialProbability(index, trials, probability)
          if (cumulative >= alpha) {
            result = <f64>index
            break
          }
        }
        if (isNaN(result)) {
          result = <f64>trials
        }
      }
    } else if (builtinId == BuiltinId.Hypgeomdist) {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base])
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3])
      const sampleSuccesses = <i32>sampleSuccessesRaw
      const sampleSize = <i32>sampleSizeRaw
      const populationSuccesses = <i32>populationSuccessesRaw
      const populationSize = <i32>populationSizeRaw
      if (isNaN(sampleSuccessesRaw) || isNaN(sampleSizeRaw) || isNaN(populationSuccessesRaw) || isNaN(populationSizeRaw)) {
        errorCode = ErrorCode.Value
      } else if (
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
        errorCode = ErrorCode.Num
      } else {
        result = hypergeometricProbability(sampleSuccesses, sampleSize, populationSuccesses, populationSize)
      }
    } else {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base])
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3])
      const cumulative = coerceBoolean(tagStack[base + 4], valueStack[base + 4])
      const sampleSuccesses = <i32>sampleSuccessesRaw
      const sampleSize = <i32>sampleSizeRaw
      const populationSuccesses = <i32>populationSuccessesRaw
      const populationSize = <i32>populationSizeRaw
      if (
        isNaN(sampleSuccessesRaw) ||
        isNaN(sampleSizeRaw) ||
        isNaN(populationSuccessesRaw) ||
        isNaN(populationSizeRaw) ||
        cumulative < 0
      ) {
        errorCode = ErrorCode.Value
      } else if (
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
        errorCode = ErrorCode.Num
      } else {
        if (cumulative == 1) {
          result = 0.0
          const minimum = max<i32>(0, sampleSize - (populationSize - populationSuccesses))
          for (let index = minimum; index <= sampleSuccesses; index += 1) {
            result += hypergeometricProbability(index, sampleSize, populationSuccesses, populationSize)
          }
        } else {
          result = hypergeometricProbability(sampleSuccesses, sampleSize, populationSuccesses, populationSize)
        }
      }
    }
    if (errorCode != ErrorCode.None) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  return -1
}
