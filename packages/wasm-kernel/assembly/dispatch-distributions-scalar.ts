import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { valueNumber } from './comparison'
import { toNumberExact } from './operands'
import { isNumericResult, rangeSupportedScalarOnly, scalarErrorAt } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'
import {
  erfApprox,
  gammaFunction,
  gammaDistributionCdf,
  gammaDistributionDensity,
  inverseGammaDistribution,
  inverseStandardNormal,
  inverseStudentT,
  logGamma,
  standardNormalCdf,
  standardNormalPdf,
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

function isNumericScalarTag(tag: u8): bool {
  return tag == ValueTag.Number || tag == ValueTag.Boolean || tag == ValueTag.Empty
}

function numericScalarTagsAt(base: i32, argc: i32, tagStack: Uint8Array): bool {
  for (let index = 0; index < argc; index += 1) {
    if (!isNumericScalarTag(tagStack[base + index])) {
      return false
    }
  }
  return true
}

function valueNumberAt(
  slot: i32,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): f64 {
  return valueNumber(
    tagStack[slot],
    valueStack[slot],
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
}

export function tryApplyScalarDistributionBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if ((builtinId == BuiltinId.Gauss || builtinId == BuiltinId.Phi) && argc == 1) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numeric = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (isNaN(numeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      builtinId == BuiltinId.Gauss ? standardNormalCdf(numeric) - 0.5 : standardNormalPdf(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Erf && (argc == 1 || argc == 2)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const lower = valueNumberAt(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const upper =
      argc == 2
        ? valueNumberAt(
            base + 1,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0
    if (isNaN(lower) || (argc == 2 && isNaN(upper))) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      argc == 2 ? erfApprox(upper) - erfApprox(lower) : erfApprox(lower),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (
    (builtinId == BuiltinId.ErfPrecise ||
      builtinId == BuiltinId.Erfc ||
      builtinId == BuiltinId.ErfcPrecise ||
      builtinId == BuiltinId.Fisher ||
      builtinId == BuiltinId.Fisherinv ||
      builtinId == BuiltinId.Gammaln ||
      builtinId == BuiltinId.GammalnPrecise ||
      builtinId == BuiltinId.Gamma) &&
    argc == 1
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const value = valueNumberAt(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (isNaN(value)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = NaN
    if (builtinId == BuiltinId.ErfPrecise) {
      result = erfApprox(value)
    } else if (builtinId == BuiltinId.Erfc || builtinId == BuiltinId.ErfcPrecise) {
      result = 1.0 - erfApprox(value)
    } else if (builtinId == BuiltinId.Fisher) {
      if (!isFinite(value) || value <= -1.0 || value >= 1.0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result = 0.5 * Math.log((1.0 + value) / (1.0 - value))
    } else if (builtinId == BuiltinId.Fisherinv) {
      const exponent = Math.exp(2.0 * value)
      result = (exponent - 1.0) / (exponent + 1.0)
    } else if (builtinId == BuiltinId.Gammaln || builtinId == BuiltinId.GammalnPrecise) {
      if (!isFinite(value) || value <= 0.0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result = logGamma(value)
      if (!isFinite(result)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    } else if (builtinId == BuiltinId.Gamma) {
      if (!isFinite(value) || (value <= 0.0 && value == Math.floor(value))) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result = gammaFunction(value)
      if (!isFinite(result)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
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

  if ((builtinId == BuiltinId.ConfidenceNorm || builtinId == BuiltinId.Confidence || builtinId == BuiltinId.ConfidenceT) && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const alpha = toNumberExact(tagStack[base], valueStack[base])
    const standardDeviation = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const sizeRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const useNormal = builtinId == BuiltinId.ConfidenceNorm || builtinId == BuiltinId.Confidence
    if (isNaN(alpha) || isNaN(standardDeviation) || isNaN(sizeRaw)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const size = <i32>sizeRaw
    if (
      !isFinite(alpha) ||
      !isFinite(standardDeviation) ||
      !isFinite(sizeRaw) ||
      alpha <= 0.0 ||
      alpha >= 1.0 ||
      standardDeviation <= 0.0 ||
      size < 1
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!useNormal && size == 1) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const critical = useNormal ? inverseStandardNormal(1.0 - alpha / 2.0) : inverseStudentT(1.0 - alpha / 2.0, <f64>(size - 1))
    const result = isNaN(critical) ? NaN : (critical * standardDeviation) / Math.sqrt(<f64>size)
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
    (builtinId == BuiltinId.Standardize && argc == 3) ||
    ((builtinId == BuiltinId.Normdist || builtinId == BuiltinId.NormDist) && argc == 4) ||
    ((builtinId == BuiltinId.Norminv || builtinId == BuiltinId.NormInv) && argc == 3) ||
    (builtinId == BuiltinId.Normsdist && argc == 1) ||
    (builtinId == BuiltinId.NormSDist && (argc == 1 || argc == 2)) ||
    (builtinId == BuiltinId.Normsinv && argc == 1) ||
    (builtinId == BuiltinId.NormSInv && argc == 1) ||
    ((builtinId == BuiltinId.Loginv || builtinId == BuiltinId.LognormInv) && argc == 3) ||
    ((builtinId == BuiltinId.Lognormdist || builtinId == BuiltinId.LognormDist) && (argc == 3 || argc == 4))
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let result = NaN
    if (builtinId == BuiltinId.Standardize) {
      if (!numericScalarTagsAt(base, 3, tagStack)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const x = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      if (isNaN(x) || isNaN(mean) || isNaN(standardDeviation)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (!(standardDeviation > 0.0)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result = (x - mean) / standardDeviation
    } else if (builtinId == BuiltinId.Normdist || builtinId == BuiltinId.NormDist) {
      if (!numericScalarTagsAt(base, 4, tagStack)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const x = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
      if (isNaN(x) || isNaN(mean) || isNaN(standardDeviation) || cumulative < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (!(standardDeviation > 0.0)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result =
        cumulative == 1
          ? standardNormalCdf((x - mean) / standardDeviation)
          : standardNormalPdf((x - mean) / standardDeviation) / standardDeviation
    } else if (builtinId == BuiltinId.Norminv || builtinId == BuiltinId.NormInv) {
      if (!numericScalarTagsAt(base, 3, tagStack)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const probability = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      if (isNaN(probability) || isNaN(mean) || isNaN(standardDeviation)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (probability <= 0.0 || probability >= 1.0 || !(standardDeviation > 0.0)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const inverse = inverseStandardNormal(probability)
      result = isNaN(inverse) ? NaN : mean + standardDeviation * inverse
    } else if (builtinId == BuiltinId.Normsdist || builtinId == BuiltinId.NormSDist) {
      const value = toNumberExact(tagStack[base], valueStack[base])
      const cumulative = builtinId == BuiltinId.NormSDist && argc == 2 ? coerceBoolean(tagStack[base + 1], valueStack[base + 1]) : 1
      result = isNaN(value) || cumulative < 0 ? NaN : cumulative == 1 ? standardNormalCdf(value) : standardNormalPdf(value)
    } else if (builtinId == BuiltinId.Normsinv || builtinId == BuiltinId.NormSInv) {
      if (!numericScalarTagsAt(base, 1, tagStack)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const probability = toNumberExact(tagStack[base], valueStack[base])
      if (isNaN(probability)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (probability <= 0.0 || probability >= 1.0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result = inverseStandardNormal(probability)
    } else if (builtinId == BuiltinId.Loginv || builtinId == BuiltinId.LognormInv) {
      if (!numericScalarTagsAt(base, 3, tagStack)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const probability = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      if (isNaN(probability) || isNaN(mean) || isNaN(standardDeviation)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (probability <= 0.0 || probability >= 1.0 || !(standardDeviation > 0.0)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const inverse = inverseStandardNormal(probability)
      result = isNaN(inverse) ? NaN : Math.exp(mean + standardDeviation * inverse)
    } else {
      if (!numericScalarTagsAt(base, argc, tagStack)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const x = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = argc == 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1
      if (isNaN(x) || isNaN(mean) || isNaN(standardDeviation) || cumulative < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (x <= 0.0 || !(standardDeviation > 0.0)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const z = (Math.log(x) - mean) / standardDeviation
      result = cumulative == 1 ? standardNormalCdf(z) : standardNormalPdf(z) / (x * standardDeviation)
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

  if ((builtinId == BuiltinId.GammaInv || builtinId == BuiltinId.Gammainv) && argc == 3) {
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
    if (isNaN(probability) || isNaN(alpha) || isNaN(beta)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (probability < 0.0 || probability >= 1.0 || alpha <= 0.0 || beta <= 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (probability == 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = inverseGammaDistribution(probability, alpha, beta)
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.NA,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  return -1
}
