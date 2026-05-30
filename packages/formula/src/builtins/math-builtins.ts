import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { CellValue } from '@bilig/protocol'
import { besselIValue, besselJValue, besselKValue, besselYValue } from './distributions.js'
import { ceilingToMultiple, coerceScalarMathNumber, floorToMultiple, moduloValue, roundToMultiple, truncateQuotient } from './numeric.js'
import { excelPower } from '../excel-power.js'
import { parseNumericText } from '../numeric-text.js'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

const EXCEL_INTEGER_LIMIT = 2 ** 53
const EXCEL_BITWISE_LIMIT = 2 ** 48 - 1
const EXCEL_BITWISE_SHIFT_LIMIT = 53

interface MathBuiltinDeps {
  toNumber: (value: CellValue | undefined) => number | undefined
  firstError: (args: readonly (CellValue | undefined)[]) => CellValue | undefined
  numberResult: (value: number) => EvaluationResult
  valueError: () => EvaluationResult
  numError: () => EvaluationResult
  numericResultOrError: (value: number) => EvaluationResult
  unaryMath: (value: CellValue | undefined, operation: (numeric: number) => number) => EvaluationResult
  binaryMath: (
    left: CellValue | undefined,
    right: CellValue | undefined,
    operation: (leftNumeric: number, rightNumeric: number) => number,
  ) => EvaluationResult
  ceilingWith: (value: CellValue | undefined, significance: CellValue | undefined) => EvaluationResult
  floorWith: (value: CellValue | undefined, significance: CellValue | undefined) => EvaluationResult
  roundWith: (value: CellValue | undefined, digits: CellValue | undefined) => EvaluationResult
  roundUpToDigits: (value: number, digits: number) => number
  roundDownToDigits: (value: number, digits: number) => number
  roundTowardZero: (value: number, digits: number) => number
  evenValue: (value: number) => number
  oddValue: (value: number) => number
  factorialValue: (value: number) => number | undefined
  doubleFactorialValue: (value: number) => number | undefined
  combinationValue: (total: number, chosen: number) => number | undefined
  multinomialValue: (values: readonly number[]) => number | undefined
  gcdPair: (left: number, right: number) => number
  lcmPair: (left: number, right: number) => number
}

function div0Error(): EvaluationResult {
  return { tag: ValueTag.Error, code: ErrorCode.Div0 }
}

function finiteNumberOrNumError(
  value: number,
  numberResult: (value: number) => EvaluationResult,
  numError: () => EvaluationResult,
): EvaluationResult {
  return Number.isFinite(value) ? numberResult(value) : numError()
}

export function createMathBuiltins({
  toNumber,
  firstError,
  numberResult,
  valueError,
  numError,
  unaryMath,
  ceilingWith,
  floorWith,
  roundWith,
  roundUpToDigits,
  roundDownToDigits,
  roundTowardZero,
  evenValue,
  oddValue,
  factorialValue,
  doubleFactorialValue,
  combinationValue,
  multinomialValue,
  gcdPair,
  lcmPair,
}: MathBuiltinDeps): Record<string, Builtin> {
  const toScalarMathNumber = (value: CellValue | undefined): number | undefined => coerceScalarMathNumber(value, toNumber)
  const toDirectAggregateNumber = (value: CellValue): number | undefined =>
    value.tag === ValueTag.String ? parseNumericText(value.value) : toNumber(value)
  const toScalarMathInteger = (value: CellValue): number | undefined => {
    const numeric = toScalarMathNumber(value)
    return numeric === undefined ? undefined : Math.trunc(numeric)
  }

  const coerceBitwiseOperand = (value: CellValue | undefined): bigint | EvaluationResult => {
    if (value === undefined) {
      return valueError()
    }
    if (value.tag === ValueTag.Error) {
      return value
    }
    const numeric = toScalarMathNumber(value)
    if (numeric === undefined) {
      return valueError()
    }
    if (!Number.isFinite(numeric) || !Number.isInteger(numeric) || numeric < 0 || numeric > EXCEL_BITWISE_LIMIT) {
      return numError()
    }
    return BigInt(numeric)
  }

  const coerceBitwiseShift = (value: CellValue | undefined): number | EvaluationResult => {
    if (value === undefined) {
      return valueError()
    }
    if (value.tag === ValueTag.Error) {
      return value
    }
    const numeric = toScalarMathNumber(value)
    if (numeric === undefined) {
      return valueError()
    }
    if (!Number.isFinite(numeric) || !Number.isInteger(numeric) || Math.abs(numeric) > EXCEL_BITWISE_SHIFT_LIMIT) {
      return numError()
    }
    return numeric
  }

  const domainCheckedUnaryMath = (
    value: CellValue,
    isValid: (numeric: number) => boolean,
    operation: (numeric: number) => number,
  ): EvaluationResult => {
    const error = firstError([value])
    if (error) {
      return error
    }
    const numeric = toScalarMathNumber(value)
    if (numeric === undefined) {
      return valueError()
    }
    return isValid(numeric) ? finiteNumberOrNumError(operation(numeric), numberResult, numError) : numError()
  }

  return {
    SIN: (value) => unaryMath(value, Math.sin),
    COS: (value) => unaryMath(value, Math.cos),
    TAN: (value) => unaryMath(value, Math.tan),
    ASIN: (value) => domainCheckedUnaryMath(value, (numeric) => numeric >= -1 && numeric <= 1, Math.asin),
    ACOS: (value) => domainCheckedUnaryMath(value, (numeric) => numeric >= -1 && numeric <= 1, Math.acos),
    ATAN: (value) => unaryMath(value, Math.atan),
    ATAN2: (left, right) => {
      const error = firstError([left, right])
      if (error) {
        return error
      }
      const x = toScalarMathNumber(left)
      const y = toScalarMathNumber(right)
      if (x === undefined || y === undefined) {
        return valueError()
      }
      if (x === 0 && y === 0) {
        return div0Error()
      }
      return numberResult(Math.atan2(y, x))
    },
    DEGREES: (value) => unaryMath(value, (numeric) => (numeric * 180) / Math.PI),
    RADIANS: (value) => unaryMath(value, (numeric) => (numeric * Math.PI) / 180),
    EXP: (value) => unaryMath(value, Math.exp),
    LN: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      return numeric > 0 ? finiteNumberOrNumError(Math.log(numeric), numberResult, numError) : numError()
    },
    LOG10: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      return numeric > 0 ? finiteNumberOrNumError(Math.log10(numeric), numberResult, numError) : numError()
    },
    LOG: (value, base) => {
      const error = firstError([value, base])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      const baseValue = base === undefined ? 10 : toScalarMathNumber(base)
      if (baseValue === undefined) {
        return valueError()
      }
      if (numeric <= 0 || baseValue <= 0) {
        return numError()
      }
      if (baseValue === 1) {
        return div0Error()
      }
      const result = base === undefined ? Math.log10(numeric) : Math.log(numeric) / Math.log(baseValue)
      return finiteNumberOrNumError(result, numberResult, numError)
    },
    POWER: (base, exponent) => {
      const error = firstError([base, exponent])
      if (error) {
        return error
      }
      const baseValue = toScalarMathNumber(base)
      const exponentValue = toScalarMathNumber(exponent)
      if (baseValue === undefined || exponentValue === undefined) {
        return valueError()
      }
      if (baseValue === 0 && exponentValue === 0) {
        return numError()
      }
      if (baseValue === 0 && exponentValue < 0) {
        return div0Error()
      }
      return finiteNumberOrNumError(excelPower(baseValue, exponentValue), numberResult, numError)
    },
    SQRT: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      return numeric < 0 ? numError() : finiteNumberOrNumError(Math.sqrt(numeric), numberResult, numError)
    },
    PI: (...args) => (args.length === 0 ? numberResult(Math.PI) : valueError()),
    SINH: (value) => unaryMath(value, Math.sinh),
    COSH: (value) => unaryMath(value, Math.cosh),
    TANH: (value) => unaryMath(value, Math.tanh),
    ASINH: (value) => unaryMath(value, Math.asinh),
    ACOSH: (value) => domainCheckedUnaryMath(value, (numeric) => numeric >= 1, Math.acosh),
    ATANH: (value) => domainCheckedUnaryMath(value, (numeric) => numeric > -1 && numeric < 1, Math.atanh),
    ACOT: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      return numberResult(numeric === 0 ? Math.PI / 2 : Math.atan(1 / numeric))
    },
    ACOTH: (value) =>
      domainCheckedUnaryMath(
        value,
        (numeric) => Math.abs(numeric) > 1,
        (numeric) => 0.5 * Math.log((numeric + 1) / (numeric - 1)),
      ),
    COT: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      const tangent = Math.tan(numeric)
      return tangent === 0 ? div0Error() : numberResult(1 / tangent)
    },
    COTH: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      const hyperbolic = Math.tanh(numeric)
      return hyperbolic === 0 ? div0Error() : numberResult(1 / hyperbolic)
    },
    CSC: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      const sine = Math.sin(numeric)
      return sine === 0 ? div0Error() : numberResult(1 / sine)
    },
    CSCH: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      const hyperbolic = Math.sinh(numeric)
      return hyperbolic === 0 ? div0Error() : numberResult(1 / hyperbolic)
    },
    SEC: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      const cosine = Math.cos(numeric)
      return cosine === 0 ? div0Error() : numberResult(1 / cosine)
    },
    SECH: (value) => unaryMath(value, (numeric) => 1 / Math.cosh(numeric)),
    SIGN: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numeric = toScalarMathNumber(value)
      if (numeric === undefined) {
        return valueError()
      }
      return numberResult(numeric === 0 ? 0 : numeric > 0 ? 1 : -1)
    },
    ROUND: (value, digits) => roundWith(value, digits),
    FLOOR: (...args) => (args.length === 2 ? floorWith(args[0], args[1]) : valueError()),
    CEILING: (...args) => (args.length === 2 ? ceilingWith(args[0], args[1]) : valueError()),
    'FLOOR.MATH': (value, significance, mode) => {
      const error = firstError([value, significance, mode])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const significanceRaw = significance === undefined ? 1 : toScalarMathNumber(significance)
      const modeValue = mode === undefined ? 0 : toScalarMathNumber(mode)
      if (numberValue === undefined || significanceRaw === undefined || modeValue === undefined) {
        return valueError()
      }
      if (significanceRaw === 0) {
        return numberResult(0)
      }
      const significanceValue = Math.abs(significanceRaw)
      if (numberValue >= 0) {
        return numberResult(floorToMultiple(numberValue, significanceValue))
      }
      const magnitude =
        modeValue === 0
          ? ceilingToMultiple(Math.abs(numberValue), significanceValue)
          : floorToMultiple(Math.abs(numberValue), significanceValue)
      return numberResult(-magnitude)
    },
    'FLOOR.PRECISE': (value, significance) => {
      const error = firstError([value, significance])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const significanceRaw = significance === undefined ? 1 : toScalarMathNumber(significance)
      if (numberValue === undefined || significanceRaw === undefined) {
        return valueError()
      }
      if (significanceRaw === 0) {
        return numberResult(0)
      }
      const significanceValue = Math.abs(significanceRaw)
      return numberResult(floorToMultiple(numberValue, significanceValue))
    },
    'CEILING.MATH': (value, significance, mode) => {
      const error = firstError([value, significance, mode])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const significanceRaw = significance === undefined ? 1 : toScalarMathNumber(significance)
      const modeValue = mode === undefined ? 0 : toScalarMathNumber(mode)
      if (numberValue === undefined || significanceRaw === undefined || modeValue === undefined) {
        return valueError()
      }
      if (significanceRaw === 0) {
        return numberResult(0)
      }
      const significanceValue = Math.abs(significanceRaw)
      if (numberValue >= 0) {
        return numberResult(ceilingToMultiple(numberValue, significanceValue))
      }
      const magnitude =
        modeValue === 0
          ? floorToMultiple(Math.abs(numberValue), significanceValue)
          : ceilingToMultiple(Math.abs(numberValue), significanceValue)
      return numberResult(-magnitude)
    },
    'CEILING.PRECISE': (value, significance) => {
      const error = firstError([value, significance])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const significanceRaw = significance === undefined ? 1 : toScalarMathNumber(significance)
      if (numberValue === undefined || significanceRaw === undefined) {
        return valueError()
      }
      if (significanceRaw === 0) {
        return numberResult(0)
      }
      const significanceValue = Math.abs(significanceRaw)
      return numberResult(ceilingToMultiple(numberValue, significanceValue))
    },
    'ISO.CEILING': (value, significance) => {
      const error = firstError([value, significance])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const significanceRaw = significance === undefined ? 1 : toScalarMathNumber(significance)
      if (numberValue === undefined || significanceRaw === undefined) {
        return valueError()
      }
      if (significanceRaw === 0) {
        return numberResult(0)
      }
      const significanceValue = Math.abs(significanceRaw)
      return numberResult(ceilingToMultiple(numberValue, significanceValue))
    },
    MOD: (left, right) => {
      const error = firstError([left, right])
      if (error) {
        return error
      }
      const divisor = toScalarMathNumber(right)
      const dividend = toScalarMathNumber(left)
      if (divisor === undefined || dividend === undefined) {
        return valueError()
      }
      if (divisor === 0) {
        return div0Error()
      }
      return numberResult(moduloValue(dividend, divisor))
    },
    BITAND: (...args) => {
      if (args.length !== 2) {
        return valueError()
      }
      const first = coerceBitwiseOperand(args[0])
      if (typeof first !== 'bigint') {
        return first
      }
      let value = first
      for (let index = 1; index < args.length; index += 1) {
        const current = coerceBitwiseOperand(args[index])
        if (typeof current !== 'bigint') {
          return current
        }
        value &= current
      }
      return numberResult(Number(value))
    },
    BITOR: (...args) => {
      if (args.length !== 2) {
        return valueError()
      }
      const first = coerceBitwiseOperand(args[0])
      if (typeof first !== 'bigint') {
        return first
      }
      let value = first
      for (let index = 1; index < args.length; index += 1) {
        const current = coerceBitwiseOperand(args[index])
        if (typeof current !== 'bigint') {
          return current
        }
        value |= current
      }
      return numberResult(Number(value))
    },
    BITXOR: (...args) => {
      if (args.length !== 2) {
        return valueError()
      }
      const first = coerceBitwiseOperand(args[0])
      if (typeof first !== 'bigint') {
        return first
      }
      let value = first
      for (let index = 1; index < args.length; index += 1) {
        const current = coerceBitwiseOperand(args[index])
        if (typeof current !== 'bigint') {
          return current
        }
        value ^= current
      }
      return numberResult(Number(value))
    },
    BITLSHIFT: (...args) => {
      if (args.length !== 2) {
        return valueError()
      }
      const [valueArg, shiftArg] = args
      const value = coerceBitwiseOperand(valueArg)
      if (typeof value !== 'bigint') {
        return value
      }
      const shift = coerceBitwiseShift(shiftArg)
      if (typeof shift !== 'number') {
        return shift
      }
      const numeric = Number(value)
      return numberResult(shift >= 0 ? numeric * 2 ** shift : Math.floor(numeric / 2 ** -shift))
    },
    BITRSHIFT: (...args) => {
      if (args.length !== 2) {
        return valueError()
      }
      const [valueArg, shiftArg] = args
      const value = coerceBitwiseOperand(valueArg)
      if (typeof value !== 'bigint') {
        return value
      }
      const shift = coerceBitwiseShift(shiftArg)
      if (typeof shift !== 'number') {
        return shift
      }
      const numeric = Number(value)
      return numberResult(shift >= 0 ? Math.floor(numeric / 2 ** shift) : numeric * 2 ** -shift)
    },
    BESSELI: (xArg, orderArg) => {
      const error = firstError([xArg, orderArg])
      if (error) {
        return error
      }
      const x = toScalarMathNumber(xArg)
      const order = toScalarMathInteger(orderArg)
      if (x === undefined || order === undefined) {
        return valueError()
      }
      if (order < 0) {
        return numError()
      }
      const result = besselIValue(x, order)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    BESSELJ: (xArg, orderArg) => {
      const error = firstError([xArg, orderArg])
      if (error) {
        return error
      }
      const x = toScalarMathNumber(xArg)
      const order = toScalarMathInteger(orderArg)
      if (x === undefined || order === undefined) {
        return valueError()
      }
      if (order < 0) {
        return numError()
      }
      const result = besselJValue(x, order)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    BESSELK: (xArg, orderArg) => {
      const error = firstError([xArg, orderArg])
      if (error) {
        return error
      }
      const x = toScalarMathNumber(xArg)
      const order = toScalarMathInteger(orderArg)
      if (x === undefined || order === undefined) {
        return valueError()
      }
      if (x <= 0 || order < 0) {
        return numError()
      }
      const result = besselKValue(x, order)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    BESSELY: (xArg, orderArg) => {
      const error = firstError([xArg, orderArg])
      if (error) {
        return error
      }
      const x = toScalarMathNumber(xArg)
      const order = toScalarMathInteger(orderArg)
      if (x === undefined || order === undefined) {
        return valueError()
      }
      if (x <= 0 || order < 0) {
        return numError()
      }
      const result = besselYValue(x, order)
      return Number.isFinite(result) ? numberResult(result) : numError()
    },
    INT: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      if (numberValue === undefined) {
        return valueError()
      }
      return numberResult(Math.floor(numberValue))
    },
    ROUNDUP: (value, digits) => {
      const error = firstError([value, digits])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const digitValue = digits === undefined ? 0 : toScalarMathNumber(digits)
      if (numberValue === undefined || digitValue === undefined) {
        return valueError()
      }
      return numberResult(roundUpToDigits(numberValue, Math.trunc(digitValue)))
    },
    ROUNDDOWN: (value, digits) => {
      const error = firstError([value, digits])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const digitValue = digits === undefined ? 0 : toScalarMathNumber(digits)
      if (numberValue === undefined || digitValue === undefined) {
        return valueError()
      }
      return numberResult(roundDownToDigits(numberValue, Math.trunc(digitValue)))
    },
    TRUNC: (value, digits) => {
      const error = firstError([value, digits])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const digitValue = digits === undefined ? 0 : toScalarMathNumber(digits)
      if (numberValue === undefined || digitValue === undefined) {
        return valueError()
      }
      return numberResult(roundTowardZero(numberValue, Math.trunc(digitValue)))
    },
    EVEN: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      return numberValue === undefined ? valueError() : numberResult(evenValue(numberValue))
    },
    ODD: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      return numberValue === undefined ? valueError() : numberResult(oddValue(numberValue))
    },
    FACT: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      if (numberValue === undefined) {
        return valueError()
      }
      const factorial = factorialValue(numberValue)
      return factorial === undefined ? numError() : finiteNumberOrNumError(factorial, numberResult, numError)
    },
    FACTDOUBLE: (value) => {
      const error = firstError([value])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      if (numberValue === undefined) {
        return valueError()
      }
      const factorial = doubleFactorialValue(numberValue)
      return factorial === undefined ? numError() : finiteNumberOrNumError(factorial, numberResult, numError)
    },
    COMBIN: (numberArg, chosenArg) => {
      const error = firstError([numberArg, chosenArg])
      if (error) {
        return error
      }
      const numberRaw = toScalarMathNumber(numberArg)
      const chosenRaw = toScalarMathNumber(chosenArg)
      if (numberRaw === undefined || chosenRaw === undefined) {
        return valueError()
      }
      const numberValue = Math.trunc(numberRaw)
      const chosenValue = Math.trunc(chosenRaw)
      if (!Number.isFinite(numberRaw) || !Number.isFinite(chosenRaw) || numberValue < 0 || chosenValue < 0 || chosenValue > numberValue) {
        return numError()
      }
      const result = combinationValue(numberValue, chosenValue)
      return result === undefined ? numError() : finiteNumberOrNumError(result, numberResult, numError)
    },
    COMBINA: (numberArg, chosenArg) => {
      const error = firstError([numberArg, chosenArg])
      if (error) {
        return error
      }
      const numberRaw = toScalarMathNumber(numberArg)
      const chosenRaw = toScalarMathNumber(chosenArg)
      if (numberRaw === undefined || chosenRaw === undefined) {
        return valueError()
      }
      const numberValue = Math.trunc(numberRaw)
      const chosenValue = Math.trunc(chosenRaw)
      if (!Number.isFinite(numberRaw) || !Number.isFinite(chosenRaw) || numberValue < 0 || chosenValue < 0) {
        return numError()
      }
      if (numberValue === 0 && chosenValue > 0) {
        return numError()
      }
      if (chosenValue === 0) {
        return numberResult(1)
      }
      const result = combinationValue(numberValue + chosenValue - 1, chosenValue)
      return result === undefined ? numError() : finiteNumberOrNumError(result, numberResult, numError)
    },
    GCD: (...args) => {
      const error = firstError(args)
      if (error) {
        return error
      }
      const numbers: number[] = []
      for (const arg of args) {
        const numeric = toScalarMathNumber(arg)
        if (numeric === undefined) {
          return valueError()
        }
        if (!Number.isFinite(numeric) || numeric < 0 || numeric >= EXCEL_INTEGER_LIMIT) {
          return numError()
        }
        numbers.push(Math.trunc(numeric))
      }
      if (numbers.length === 0) {
        return valueError()
      }
      return numberResult(numbers.reduce((acc, value) => gcdPair(acc, value)))
    },
    LCM: (...args) => {
      const error = firstError(args)
      if (error) {
        return error
      }
      const numbers: number[] = []
      for (const arg of args) {
        const numeric = toScalarMathNumber(arg)
        if (numeric === undefined) {
          return valueError()
        }
        if (!Number.isFinite(numeric) || numeric < 0 || numeric >= EXCEL_INTEGER_LIMIT) {
          return numError()
        }
        numbers.push(Math.trunc(numeric))
      }
      if (numbers.length === 0) {
        return valueError()
      }
      let result = numbers[0] ?? 0
      for (let index = 1; index < numbers.length; index += 1) {
        result = lcmPair(result, numbers[index] ?? 0)
        if (!Number.isFinite(result) || result >= EXCEL_INTEGER_LIMIT) {
          return numError()
        }
      }
      return numberResult(result)
    },
    MROUND: (value, multiple) => {
      const error = firstError([value, multiple])
      if (error) {
        return error
      }
      const numberValue = toScalarMathNumber(value)
      const multipleValue = toScalarMathNumber(multiple)
      if (numberValue === undefined || multipleValue === undefined) {
        return valueError()
      }
      if (multipleValue === 0) {
        return numberResult(0)
      }
      if (numberValue !== 0 && Math.sign(numberValue) !== Math.sign(multipleValue)) {
        return numError()
      }
      return numberResult(roundToMultiple(numberValue, multipleValue))
    },
    MULTINOMIAL: (...args) => {
      const error = firstError(args)
      if (error) {
        return error
      }
      const numbers: number[] = []
      for (const arg of args) {
        const numberValue = toScalarMathNumber(arg)
        if (numberValue === undefined) {
          return valueError()
        }
        if (!Number.isFinite(numberValue)) {
          return numError()
        }
        numbers.push(Math.trunc(numberValue))
      }
      if (numbers.some((value) => value < 0)) {
        return numError()
      }
      const result = multinomialValue(numbers)
      return result === undefined ? numError() : finiteNumberOrNumError(result, numberResult, numError)
    },
    PRODUCT: (...args) => {
      const error = firstError(args)
      if (error) {
        return error
      }
      const numbers: number[] = []
      for (const arg of args) {
        const numeric = toDirectAggregateNumber(arg)
        if (numeric === undefined) {
          return valueError()
        }
        numbers.push(numeric)
      }
      return numberResult(numbers.length === 0 ? 0 : numbers.reduce((product, value) => product * value, 1))
    },
    QUOTIENT: (numeratorArg, denominatorArg) => {
      const error = firstError([numeratorArg, denominatorArg])
      if (error) {
        return error
      }
      const numerator = toScalarMathNumber(numeratorArg)
      const denominator = toScalarMathNumber(denominatorArg)
      if (numerator === undefined || denominator === undefined) {
        return valueError()
      }
      if (denominator === 0) {
        return div0Error()
      }
      return numberResult(truncateQuotient(numerator, denominator))
    },
  }
}
