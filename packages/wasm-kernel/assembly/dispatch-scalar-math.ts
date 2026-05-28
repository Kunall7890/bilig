import { BuiltinId, ErrorCode } from './protocol'
import {
  combinationCalc,
  doubleFactorialCalc,
  evenCalc,
  factorialCalc,
  oddCalc,
  permutationCalc,
  roundToMultiple,
  truncateQuotient,
  truncToInt,
} from './numeric-core'
import { toNumberExact, toNumberOrZero } from './operands'
import { besselIValue, besselJValue, besselKValue, besselYValue } from './distributions'
import { tryApplyScalarRoundingMathBuiltin } from './dispatch-scalar-rounding-math'
import { excelPower } from './vm-core-helpers'
import {
  firstScalarMathError,
  scalarMathNumberLikeText,
  writeScalarMathError,
  writeScalarMathFiniteNumberOrNum,
  writeScalarMathNumber,
} from './dispatch-scalar-math-helpers'

export function tryApplyScalarMathBuiltin(
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
  if (
    (builtinId == BuiltinId.Besseli ||
      builtinId == BuiltinId.Besselj ||
      builtinId == BuiltinId.Besselk ||
      builtinId == BuiltinId.Bessely) &&
    argc == 2
  ) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = scalarMathNumberLikeText(
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
    const orderNumeric = scalarMathNumberLikeText(
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
    if (!isFinite(x) || !isFinite(orderNumeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const order = <i32>orderNumeric
    if (order < 0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if ((builtinId == BuiltinId.Besselk || builtinId == BuiltinId.Bessely) && x <= 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = NaN
    if (builtinId == BuiltinId.Besseli) {
      result = besselIValue(x, order)
    } else if (builtinId == BuiltinId.Besselj) {
      result = besselJValue(x, order)
    } else if (builtinId == BuiltinId.Besselk) {
      result = besselKValue(x, order)
    } else {
      result = besselYValue(x, order)
    }
    return !isFinite(result)
      ? writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const scalarRoundingMathResult = tryApplyScalarRoundingMathBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
  if (scalarRoundingMathResult >= 0) {
    return scalarRoundingMathResult
  }

  if (builtinId == BuiltinId.Sin && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathFiniteNumberOrNum(base, Math.sin(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Cos && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathFiniteNumberOrNum(base, Math.cos(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Tan && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathFiniteNumberOrNum(base, Math.tan(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Asin && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    if (!isFinite(numeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric < -1.0 || numeric > 1.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, Math.asin(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Acos && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    if (!isFinite(numeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric < -1.0 || numeric > 1.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, Math.acos(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Atan && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathFiniteNumberOrNum(base, Math.atan(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Atan2 && argc == 2) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = scalarMathNumberLikeText(
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
    const y = scalarMathNumberLikeText(
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
    if (isNaN(x) || isNaN(y)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (x == 0.0 && y == 0.0) {
      return writeScalarMathError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, Math.atan2(y, x), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Degrees && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathFiniteNumberOrNum(base, (numeric * 180.0) / Math.PI, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Radians && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathFiniteNumberOrNum(base, (numeric * Math.PI) / 180.0, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Exp && argc == 1) {
    const numeric = scalarMathNumberLikeText(
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
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathFiniteNumberOrNum(base, Math.exp(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Ln && argc == 1) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numeric = scalarMathNumberLikeText(
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
    if (isNaN(numeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric <= 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = Math.log(numeric)
    return isFinite(result)
      ? writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Log10 && argc == 1) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numeric = scalarMathNumberLikeText(
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
    if (isNaN(numeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric <= 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = Math.log10(numeric)
    return isFinite(result)
      ? writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Log && (argc == 1 || argc == 2)) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const num = scalarMathNumberLikeText(
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
    const baseVal =
      argc == 2
        ? scalarMathNumberLikeText(
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
        : 10.0
    if (isNaN(num) || isNaN(baseVal)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (num <= 0.0 || baseVal <= 0.0 || baseVal == 1.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = Math.log(num) / Math.log(baseVal)
    return isFinite(result)
      ? writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Power && argc == 2) {
    const baseValue = scalarMathNumberLikeText(
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
    const exponentValue = scalarMathNumberLikeText(
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
    if (isNaN(baseValue) || isNaN(exponentValue)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathFiniteNumberOrNum(base, excelPower(baseValue, exponentValue), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Sqrt && argc == 1) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numeric = scalarMathNumberLikeText(
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
    if (isNaN(numeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric < 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, Math.sqrt(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Seriessum && argc >= 3) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const x = toNumberExact(tagStack[base], valueStack[base])
    const n = truncToInt(tagStack[base + 1], valueStack[base + 1])
    const m = truncToInt(tagStack[base + 2], valueStack[base + 2])
    if (isNaN(x) || n == i32.MIN_VALUE || m == i32.MIN_VALUE) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let sum = 0.0
    for (let index = 0; index < argc - 3; index += 1) {
      const coefficient = toNumberOrZero(tagStack[base + 3 + index], valueStack[base + 3 + index])
      sum += coefficient * Math.pow(x, <f64>(n + index * m))
    }
    return writeScalarMathNumber(base, sum, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Sqrtpi && argc == 1) {
    const error = firstScalarMathError(base, argc, valueStack, tagStack)
    if (error != ErrorCode.None) {
      return writeScalarMathError(base, error, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numeric = scalarMathNumberLikeText(
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
    if (isNaN(numeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = Math.sqrt(numeric * Math.PI)
    return !isFinite(result)
      ? writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Pi && argc == 0) {
    return writeScalarMathNumber(base, Math.PI, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (builtinId == BuiltinId.Pi) {
    return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    (builtinId == BuiltinId.Sinh ||
      builtinId == BuiltinId.Cosh ||
      builtinId == BuiltinId.Tanh ||
      builtinId == BuiltinId.Asinh ||
      builtinId == BuiltinId.Acosh ||
      builtinId == BuiltinId.Atanh ||
      builtinId == BuiltinId.Acot ||
      builtinId == BuiltinId.Acoth ||
      builtinId == BuiltinId.Cot ||
      builtinId == BuiltinId.Coth ||
      builtinId == BuiltinId.Csc ||
      builtinId == BuiltinId.Csch ||
      builtinId == BuiltinId.Sec ||
      builtinId == BuiltinId.Sech ||
      builtinId == BuiltinId.Sign ||
      builtinId == BuiltinId.Even ||
      builtinId == BuiltinId.Odd ||
      builtinId == BuiltinId.Fact ||
      builtinId == BuiltinId.Factdouble) &&
    argc == 1
  ) {
    const isGenericTranscendental =
      builtinId == BuiltinId.Sinh ||
      builtinId == BuiltinId.Cosh ||
      builtinId == BuiltinId.Tanh ||
      builtinId == BuiltinId.Asinh ||
      builtinId == BuiltinId.Sech
    const numeric = scalarMathNumberLikeText(
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
    if (!isFinite(numeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if ((builtinId == BuiltinId.Fact || builtinId == BuiltinId.Factdouble) && numeric < 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (
      (builtinId == BuiltinId.Acosh && numeric < 1.0) ||
      (builtinId == BuiltinId.Atanh && (numeric <= -1.0 || numeric >= 1.0)) ||
      (builtinId == BuiltinId.Acoth && Math.abs(numeric) <= 1.0)
    ) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let result = 0.0
    let errorCode = ErrorCode.None
    if (builtinId == BuiltinId.Sinh) {
      result = Math.sinh(numeric)
    } else if (builtinId == BuiltinId.Cosh) {
      result = Math.cosh(numeric)
    } else if (builtinId == BuiltinId.Tanh) {
      result = Math.tanh(numeric)
    } else if (builtinId == BuiltinId.Asinh) {
      result = Math.asinh(numeric)
    } else if (builtinId == BuiltinId.Acosh) {
      result = Math.acosh(numeric)
    } else if (builtinId == BuiltinId.Atanh) {
      result = Math.atanh(numeric)
    } else if (builtinId == BuiltinId.Acot) {
      result = numeric == 0.0 ? Math.PI / 2.0 : Math.atan(1.0 / numeric)
    } else if (builtinId == BuiltinId.Acoth) {
      result = 0.5 * Math.log((numeric + 1.0) / (numeric - 1.0))
    } else if (builtinId == BuiltinId.Cot) {
      const tangent = Math.tan(numeric)
      if (tangent == 0.0) {
        errorCode = ErrorCode.Div0
      } else {
        result = 1.0 / tangent
      }
    } else if (builtinId == BuiltinId.Coth) {
      const hyperbolic = Math.tanh(numeric)
      if (hyperbolic == 0.0) {
        errorCode = ErrorCode.Div0
      } else {
        result = 1.0 / hyperbolic
      }
    } else if (builtinId == BuiltinId.Csc) {
      const sine = Math.sin(numeric)
      if (sine == 0.0) {
        errorCode = ErrorCode.Div0
      } else {
        result = 1.0 / sine
      }
    } else if (builtinId == BuiltinId.Csch) {
      const hyperbolic = Math.sinh(numeric)
      if (hyperbolic == 0.0) {
        errorCode = ErrorCode.Div0
      } else {
        result = 1.0 / hyperbolic
      }
    } else if (builtinId == BuiltinId.Sec) {
      const cosine = Math.cos(numeric)
      if (cosine == 0.0) {
        errorCode = ErrorCode.Div0
      } else {
        result = 1.0 / cosine
      }
    } else if (builtinId == BuiltinId.Sech) {
      result = 1.0 / Math.cosh(numeric)
    } else if (builtinId == BuiltinId.Sign) {
      result = numeric == 0.0 ? 0.0 : numeric > 0.0 ? 1.0 : -1.0
    } else if (builtinId == BuiltinId.Even) {
      result = evenCalc(numeric)
    } else if (builtinId == BuiltinId.Odd) {
      result = oddCalc(numeric)
    } else if (builtinId == BuiltinId.Fact) {
      result = factorialCalc(numeric)
    } else {
      result = doubleFactorialCalc(numeric)
    }

    if (errorCode != ErrorCode.None) {
      return writeScalarMathError(base, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return !isFinite(result)
      ? writeScalarMathError(
          base,
          builtinId == BuiltinId.Fact || builtinId == BuiltinId.Factdouble || isGenericTranscendental ? ErrorCode.Num : ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Combin || builtinId == BuiltinId.Combina || builtinId == BuiltinId.Quotient) && argc == 2) {
    const left = scalarMathNumberLikeText(
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
    const right = scalarMathNumberLikeText(
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
    if (!isFinite(left) || !isFinite(right)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (builtinId == BuiltinId.Quotient) {
      if (right == 0.0) {
        return writeScalarMathError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      return writeScalarMathNumber(base, truncateQuotient(left, right), rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const numberValue = Math.trunc(left)
    const chosenValue = Math.trunc(right)
    if (numberValue < 0.0 || chosenValue < 0.0 || (builtinId == BuiltinId.Combina && numberValue < chosenValue)) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (builtinId == BuiltinId.Combin && chosenValue > numberValue) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let result = 0.0
    if (builtinId == BuiltinId.Combin) {
      result = combinationCalc(numberValue, chosenValue)
    } else if (chosenValue == 0.0) {
      result = 1.0
    } else {
      result = combinationCalc(numberValue + chosenValue - 1.0, chosenValue)
    }
    return !isFinite(result)
      ? writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Permut || builtinId == BuiltinId.Permutationa) && argc == 2) {
    const left = scalarMathNumberLikeText(
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
    const right = scalarMathNumberLikeText(
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
    if (!isFinite(left) || !isFinite(right)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numberValue = Math.trunc(left)
    const chosenValue = Math.trunc(right)
    if (numberValue < 0.0 || chosenValue < 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (builtinId == BuiltinId.Permut && (numberValue <= 0.0 || chosenValue > numberValue)) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (builtinId == BuiltinId.Permutationa && numberValue == 0.0 && chosenValue > 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = 0.0
    if (builtinId == BuiltinId.Permut) {
      result = permutationCalc(numberValue, chosenValue)
    } else {
      result = Math.pow(numberValue, chosenValue)
    }
    return !isFinite(result)
      ? writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Mround && argc == 2) {
    const numeric = scalarMathNumberLikeText(
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
    const multiple = scalarMathNumberLikeText(
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
    if (isNaN(numeric) || isNaN(multiple)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (multiple == 0.0) {
      return writeScalarMathNumber(base, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric != 0.0 && Math.sign(numeric) != Math.sign(multiple)) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, roundToMultiple(numeric, multiple), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
