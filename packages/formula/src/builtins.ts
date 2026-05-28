import { BUILTINS, BuiltinId, ErrorCode, ValueTag } from '@bilig/protocol'
import type { CellValue } from '@bilig/protocol'
import { builtinJsSpecialNames } from './builtin-capabilities.js'
import { createComplexBuiltins } from './builtins/complex.js'
import { createDistributionBuiltins } from './builtins/distribution-builtins.js'
import { createFinancialBuiltins } from './builtins/financial-builtins.js'
import { createFixedIncomeBuiltins } from './builtins/fixed-income-builtins.js'
import { countLeadingZeros, formatFixed, parseDollarDecimal, toColumnLabel } from './builtins/formatting.js'
import {
  buildIdentityMatrix,
  combinationValue,
  collectNumericArgs,
  createNumericBuiltinHelpers,
  doubleFactorialValue,
  evenValue,
  factorialValue,
  gcdPair,
  lcmPair,
  multinomialValue,
  oddValue,
  permutationValue,
  roundDownToDigits,
  roundTowardZero,
  roundUpToDigits,
} from './builtins/numeric.js'
import { toAverageNumber, toDirectAggregateNumber, toNumber, toScalarMathNumber } from './builtins/scalar-coercion.js'
import { createMathBuiltins } from './builtins/math-builtins.js'
import { createRadixBuiltins } from './builtins/radix.js'
import { populationVariance, sampleVariance } from './builtins/statistics.js'
import { createStatisticalBuiltins } from './builtins/statistical-builtins.js'
import { createDateTimeBuiltins, datetimeBuiltins, type ExcelDateSystem } from './builtins/datetime.js'
import { convertBuiltin, euroconvertBuiltin } from './builtins/convert.js'
import { logicalBuiltins } from './builtins/logical.js'
import { lookupBuiltins } from './builtins/lookup.js'
import { createBlockedBuiltinMap, scalarPlaceholderBuiltinNames } from './builtins/placeholder.js'
import { getExternalScalarFunction, hasExternalFunction } from './external-function-adapter.js'
import { coerceLogicalValue } from './logical-coercion.js'
import type { ArrayValue, EvaluationResult } from './runtime-values.js'
import { createTextBuiltins, textBuiltins } from './builtins/text.js'
import { quoteSheetNameIfNeeded } from './translation-reference-utils.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

export function normalizeBuiltinLookupName(name: string): string {
  const upper = name.toUpperCase()
  return upper.startsWith('_XLFN.') || upper.startsWith('_XLWS.') ? upper.slice(6) : upper
}

function isCountedDirectNumber(value: CellValue): boolean {
  return (
    value.tag === ValueTag.Number ||
    value.tag === ValueTag.Boolean ||
    (value.tag === ValueTag.String && toDirectAggregateNumber(value) !== undefined)
  )
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}

function div0Error(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Div0 }
}

function blockedError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Blocked }
}

function numError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Num }
}

function numericResultOrError(value: number): CellValue {
  return Number.isFinite(value) ? numberResult(value) : valueError()
}

function firstError(args: readonly (CellValue | undefined)[]): CellValue | undefined {
  return args.find((arg): arg is CellValue => arg?.tag === ValueTag.Error)
}

function preserveIncomingErrors(map: Record<string, Builtin>): Record<string, Builtin> {
  return Object.fromEntries(
    Object.entries(map).map(([name, builtin]) => [
      name,
      (...args: CellValue[]): EvaluationResult => {
        const error = firstError(args)
        return error ?? builtin(...args)
      },
    ]),
  )
}

function coercePositiveInteger(value: CellValue | undefined, fallback: number): number | undefined {
  if (value === undefined) {
    return fallback
  }
  const numeric = toNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  const truncated = Math.trunc(numeric)
  return truncated >= 1 ? truncated : undefined
}

function coerceNumber(value: CellValue | undefined, fallback: number): number | undefined {
  if (value === undefined) {
    return fallback
  }
  const numeric = toNumber(value)
  return numeric !== undefined && Number.isFinite(numeric) ? numeric : undefined
}

function coerceScalarNumber(value: CellValue | undefined, fallback: number): number | undefined {
  if (value === undefined) {
    return fallback
  }
  const numeric = toScalarMathNumber(value)
  return numeric !== undefined && Number.isFinite(numeric) ? numeric : undefined
}

function integerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  if (value === undefined) {
    return fallback
  }
  const numeric = toNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  return Math.trunc(numeric)
}

function scalarIntegerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  if (value === undefined) {
    return fallback
  }
  const numeric = toScalarMathNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  return Math.trunc(numeric)
}

function toInteger(value: CellValue, fallback?: number): number | undefined {
  if (value === undefined) {
    return fallback
  }
  const numeric = toNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  const truncated = Math.trunc(numeric)
  return Number.isSafeInteger(truncated) ? truncated : undefined
}

function nonNegativeIntegerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  const integer = integerValue(value, fallback)
  return integer !== undefined && integer >= 0 ? integer : undefined
}

function positiveIntegerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  const integer = integerValue(value, fallback)
  return integer !== undefined && integer >= 1 ? integer : undefined
}

function sequenceResult(
  rowsArg: CellValue | undefined,
  colsArg: CellValue | undefined,
  startArg: CellValue | undefined,
  stepArg: CellValue | undefined,
): ArrayValue | CellValue {
  const error = firstError([rowsArg, colsArg, startArg, stepArg])
  if (error) {
    return error
  }
  const rows = coercePositiveInteger(rowsArg, 1)
  const cols = coercePositiveInteger(colsArg, 1)
  const start = coerceNumber(startArg, 1)
  const step = coerceNumber(stepArg, 1)
  if (rows === undefined || cols === undefined || start === undefined || step === undefined) {
    return valueError()
  }

  const values: CellValue[] = []
  for (let index = 0; index < rows * cols; index += 1) {
    values.push(numberResult(start + index * step))
  }
  return {
    kind: 'array',
    rows,
    cols,
    values,
  }
}

function coerceDateSerial(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }
  const serial = toNumber(value)
  return serial !== undefined && Number.isFinite(serial) ? Math.trunc(serial) : undefined
}

function coerceBoolean(value: CellValue | undefined, fallback: boolean): boolean | undefined {
  if (value === undefined) {
    return fallback
  }
  const coerced = coerceLogicalValue(value)
  return coerced.ok ? coerced.value : undefined
}

const complexBuiltins = preserveIncomingErrors(createComplexBuiltins({ toNumber, numberResult, valueError }))
const numericBuiltinHelpers = createNumericBuiltinHelpers({
  toNumber,
  numberResult,
  valueError,
  numError,
})
const { binaryMath, ceilingWith, floorWith, roundWith, unaryMath } = numericBuiltinHelpers
const radixBuiltins = preserveIncomingErrors(
  createRadixBuiltins({
    toNumber,
    integerValue,
    nonNegativeIntegerValue,
    valueError,
    numberResult,
  }),
)
const fixedIncomeBuiltins = preserveIncomingErrors(
  createFixedIncomeBuiltins({
    toNumber,
    coerceBoolean,
    coerceDateSerial,
    coerceNumber,
    integerValue,
    numberResult,
    valueError,
    numError,
  }),
)
const statisticalBuiltins = preserveIncomingErrors(
  createStatisticalBuiltins({
    toNumber: toScalarMathNumber,
    coerceBoolean,
    firstError,
    numberResult,
    numericResultOrError,
    valueError,
    numError,
  }),
)
const distributionBuiltins = preserveIncomingErrors(
  createDistributionBuiltins({
    toNumber: toScalarMathNumber,
    coerceBoolean,
    coerceNumber,
    nonNegativeIntegerValue,
    positiveIntegerValue,
    numberResult,
    numericResultOrError,
    valueError,
    numError,
  }),
)
const financialBuiltins = preserveIncomingErrors(
  createFinancialBuiltins({
    toNumber: toScalarMathNumber,
    coerceBoolean,
    coerceNumber: coerceScalarNumber,
    coercePaymentType,
    integerValue: scalarIntegerValue,
    numberResult,
    valueError,
    numError,
  }),
)
const mathBuiltins = createMathBuiltins({
  toNumber,
  firstError,
  numberResult,
  valueError,
  numError,
  numericResultOrError,
  unaryMath,
  binaryMath,
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
})

function coercePaymentType(value: CellValue | undefined, fallback: number): number | undefined {
  const type = scalarIntegerValue(value, fallback)
  return type === 0 || type === 1 ? type : undefined
}

function coerceDollarFraction(value: CellValue | undefined): number | CellValue {
  if (value === undefined) {
    return valueError()
  }
  const numeric = toNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return valueError()
  }
  if (numeric < 0) {
    return numError()
  }
  const fraction = Math.trunc(numeric)
  return fraction < 1 ? div0Error() : fraction
}

function toZeroNumericValue(value: CellValue): number | undefined {
  if (value.tag === ValueTag.String) {
    return 0
  }
  return toNumber(value)
}

function aggregateByCode(functionNum: number, values: CellValue[], options: { readonly propagateErrors?: boolean } = {}): CellValue {
  if (options.propagateErrors) {
    const error = firstError(values)
    if (error) {
      return error
    }
  }
  const normalized = functionNum > 100 ? functionNum - 100 : functionNum
  const numericValues = collectNumericArgs(values, toNumber)
  switch (normalized) {
    case 1:
      return numericValues.length === 0
        ? div0Error()
        : numberResult(numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length)
    case 2:
      return numberResult(values.filter((value) => value.tag === ValueTag.Number || value.tag === ValueTag.Boolean).length)
    case 3:
      return numberResult(values.filter((value) => value.tag !== ValueTag.Empty).length)
    case 4:
      return numberResult(numericValues.length === 0 ? 0 : Math.max(...numericValues))
    case 5:
      return numberResult(numericValues.length === 0 ? 0 : Math.min(...numericValues))
    case 6:
      return numberResult(numericValues.length === 0 ? 0 : numericValues.reduce((product, value) => product * value, 1))
    case 7:
      return numberResult(Math.sqrt(sampleVariance(numericValues)))
    case 8:
      return numberResult(Math.sqrt(populationVariance(numericValues)))
    case 9:
      return numberResult(numericValues.reduce((sum, value) => sum + value, 0))
    case 10:
      return numberResult(sampleVariance(numericValues))
    case 11:
      return numberResult(populationVariance(numericValues))
    default:
      return valueError()
  }
}

function aggregateOptionIgnoresErrors(option: number): boolean {
  return option === 2 || option === 3 || option === 6 || option === 7
}

const scalarPlaceholderBuiltins = createBlockedBuiltinMap(scalarPlaceholderBuiltinNames)

const externalScalarBuiltinNames = [
  'CALL',
  'CUBEKPIMEMBER',
  'CUBEMEMBER',
  'CUBEMEMBERPROPERTY',
  'CUBERANKEDMEMBER',
  'CUBESET',
  'CUBESETCOUNT',
  'CUBEVALUE',
  'DDE',
  'DETECTLANGUAGE',
  'HYPERLINK',
  'IMAGE',
  'INFO',
  'REGISTER.ID',
  'RTD',
  'TRANSLATE',
  'WEBSERVICE',
] as const

function createExternalScalarBuiltin(name: string): Builtin {
  return (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const external = getExternalScalarFunction(name)
    if (external) {
      return external(...args)
    }
    return name === 'HYPERLINK' ? hyperlinkDisplayValue(args) : blockedError()
  }
}

function hyperlinkDisplayValue(args: CellValue[]): CellValue {
  if (args.length < 1 || args.length > 2) {
    return valueError()
  }
  const linkLocation = args[0]
  if (!linkLocation || (linkLocation.tag !== ValueTag.String && linkLocation.tag !== ValueTag.Empty)) {
    return valueError()
  }
  return args[1] ?? linkLocation
}

const externalScalarBuiltins = Object.fromEntries(
  externalScalarBuiltinNames.map((name) => [name, createExternalScalarBuiltin(name)]),
) as Record<string, Builtin>

const scalarBuiltins: Record<string, Builtin> = {
  SUM: (...args) => {
    const error = firstError(args)
    if (error) return error
    return numberResult(args.reduce((sum, arg) => sum + (toNumber(arg) ?? 0), 0))
  },
  AVERAGEA: (...args) => {
    const error = firstError(args)
    if (error) return error
    const numbers = args.map((arg) => toZeroNumericValue(arg)).filter((value): value is number => value !== undefined)
    return numbers.length === 0 ? div0Error() : numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length)
  },
  AVERAGE: (...args) => {
    const error = firstError(args)
    if (error) return error
    const numbers = collectNumericArgs(args, toAverageNumber)
    if (numbers.length === 0) return div0Error()
    return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length)
  },
  AVG: (...args) => {
    const error = firstError(args)
    if (error) return error
    const numbers = collectNumericArgs(args, toAverageNumber)
    if (numbers.length === 0) return div0Error()
    return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length)
  },
  CHOOSE: (indexValue, ...values) => {
    if (indexValue?.tag === ValueTag.Error) {
      return indexValue
    }
    const index = integerValue(indexValue)
    if (index === undefined || index < 1 || index > values.length) {
      return valueError()
    }
    const value = values[index - 1]
    return value === undefined ? valueError() : value
  },
  SINGLE: (value, ...rest) => {
    if (value === undefined || rest.length > 0) {
      return valueError()
    }
    return value
  },
  MIN: (...args) => {
    const error = firstError(args)
    if (error) return error
    const values: number[] = []
    for (const arg of args) {
      const numeric = toDirectAggregateNumber(arg)
      if (numeric === undefined) {
        return valueError()
      }
      values.push(numeric)
    }
    return values.length === 0 ? numberResult(0) : numberResult(Math.min(...values))
  },
  MAX: (...args) => {
    const error = firstError(args)
    if (error) return error
    const values: number[] = []
    for (const arg of args) {
      const numeric = toDirectAggregateNumber(arg)
      if (numeric === undefined) {
        return valueError()
      }
      values.push(numeric)
    }
    return values.length === 0 ? numberResult(0) : numberResult(Math.max(...values))
  },
  MAXA: (...args) => {
    const error = firstError(args)
    if (error) return error
    const values = args.map((arg) => toZeroNumericValue(arg)).filter((value): value is number => value !== undefined)
    return values.length === 0 ? numberResult(0) : numberResult(Math.max(...values))
  },
  MINA: (...args) => {
    const error = firstError(args)
    if (error) return error
    const values = args.map((arg) => toZeroNumericValue(arg)).filter((value): value is number => value !== undefined)
    return values.length === 0 ? numberResult(0) : numberResult(Math.min(...values))
  },
  COUNT: (...args) => {
    return numberResult(args.filter(isCountedDirectNumber).length)
  },
  COUNTA: (...args) => numberResult(args.filter((arg) => arg.tag !== ValueTag.Empty).length),
  COUNTBLANK: (...args) => {
    let blanks = 0
    for (const arg of args) {
      if (arg.tag === ValueTag.Empty || (arg.tag === ValueTag.String && arg.value === '')) {
        blanks += 1
      }
    }
    return numberResult(blanks)
  },
  ABS: (value) => {
    const error = firstError([value])
    if (error) {
      return error
    }
    const numeric = toScalarMathNumber(value)
    return numeric === undefined ? valueError() : numberResult(Math.abs(numeric))
  },
  ADDRESS: (
    rowArg,
    columnArg,
    absNumArg = { tag: ValueTag.Number, value: 1 },
    refStyleArg = { tag: ValueTag.Number, value: 1 },
    sheetTextArg,
  ) => {
    const error = firstError([rowArg, columnArg, absNumArg, refStyleArg, sheetTextArg])
    if (error) {
      return error
    }
    const row = positiveIntegerValue(rowArg)
    const column = positiveIntegerValue(columnArg)
    const absNum = integerValue(absNumArg, 1)
    const a1Style = coerceLogicalValue(refStyleArg)
    if (row === undefined || column === undefined || absNum === undefined || !a1Style.ok) {
      return valueError()
    }
    if (![1, 2, 3, 4].includes(absNum)) {
      return valueError()
    }
    if (sheetTextArg !== undefined && sheetTextArg.tag !== ValueTag.String && sheetTextArg.tag !== ValueTag.Empty) {
      return valueError()
    }
    if (sheetTextArg?.tag === ValueTag.Empty) {
      return valueError()
    }
    const columnLabel = toColumnLabel(column)
    if (columnLabel === undefined) {
      return valueError()
    }
    const sheetPrefix = sheetTextArg?.tag === ValueTag.String ? `${quoteSheetNameIfNeeded(sheetTextArg.value)}!` : ''
    if (!a1Style.value) {
      const rowLabel = absNum === 1 || absNum === 2 ? String(row) : `[${row}]`
      const colLabel = absNum === 1 || absNum === 3 ? String(column) : `[${column}]`
      return {
        tag: ValueTag.String,
        value: `${sheetPrefix}R${rowLabel}C${colLabel}`,
        stringId: 0,
      }
    }
    const rowLabel = absNum === 1 || absNum === 2 ? `$${row}` : `${row}`
    const colLabel = absNum === 1 || absNum === 3 ? `$${columnLabel}` : columnLabel
    return {
      tag: ValueTag.String,
      value: `${sheetPrefix}${colLabel}${rowLabel}`,
      stringId: 0,
    }
  },
  DOLLAR: (valueArg, decimalsArg = { tag: ValueTag.Number, value: 2 }, noCommasArg) => {
    const error = firstError([valueArg, decimalsArg, noCommasArg])
    if (error) {
      return error
    }
    const value = toNumber(valueArg)
    const decimals = toInteger(decimalsArg, 2)
    const noCommasValue = noCommasArg === undefined ? 0 : (toNumber(noCommasArg) ?? 0)
    if (value === undefined || decimals === undefined) {
      return valueError()
    }
    const text = formatFixed(value, decimals, noCommasValue === 0)
    if (text === '') {
      return valueError()
    }
    const normalizedText = text.startsWith('-') ? text.slice(1) : text
    return {
      tag: ValueTag.String,
      value: value < 0 ? `-$${normalizedText}` : `$${text}`,
      stringId: 0,
    }
  },
  FIXED: (valueArg, decimalsArg = { tag: ValueTag.Number, value: 2 }, noCommasArg) => {
    const error = firstError([valueArg, decimalsArg, noCommasArg])
    if (error) {
      return error
    }
    const value = toNumber(valueArg)
    const decimals = toInteger(decimalsArg, 2)
    const noCommasValue = noCommasArg === undefined ? 0 : (toNumber(noCommasArg) ?? 0)
    if (value === undefined || decimals === undefined) {
      return valueError()
    }
    const text = formatFixed(value, decimals, noCommasValue === 0)
    return text === '' ? valueError() : { tag: ValueTag.String, value: text, stringId: 0 }
  },
  DOLLARDE: (valueArg, fractionArg) => {
    const error = firstError([valueArg, fractionArg])
    if (error) {
      return error
    }
    const value = toNumber(valueArg)
    const fraction = coerceDollarFraction(fractionArg)
    if (value === undefined || !Number.isFinite(value)) {
      return valueError()
    }
    if (typeof fraction !== 'number') {
      return fraction
    }
    const { integerPart, fractionalNumerator } = parseDollarDecimal(value)
    if (!Number.isInteger(fractionalNumerator)) {
      return valueError()
    }
    const sign = value < 0 ? -1 : 1
    return numberResult(sign * (integerPart + fractionalNumerator / fraction))
  },
  DOLLARFR: (valueArg, fractionArg) => {
    const error = firstError([valueArg, fractionArg])
    if (error) {
      return error
    }
    const value = toNumber(valueArg)
    const fraction = coerceDollarFraction(fractionArg)
    if (value === undefined || !Number.isFinite(value)) {
      return valueError()
    }
    if (typeof fraction !== 'number') {
      return fraction
    }
    const sign = value < 0 ? -1 : 1
    const absolute = Math.abs(value)
    const integerPart = Math.floor(absolute)
    const fractional = absolute - integerPart
    const width = countLeadingZeros(fraction)
    const scaledNumerator = Math.round(fractional * fraction)
    const carry = Math.floor(scaledNumerator / fraction)
    const numerator = scaledNumerator - carry * fraction
    const outputValue = `${integerPart + carry}.${String(numerator).padStart(width, '0')}`
    return numberResult(sign * Number(outputValue))
  },
  GEOMEAN: (...args) => {
    const error = firstError(args)
    if (error) return error
    const numbers: number[] = []
    for (const arg of args) {
      const numeric = toDirectAggregateNumber(arg)
      if (numeric === undefined) {
        return valueError()
      }
      numbers.push(numeric)
    }
    if (numbers.length === 0) {
      return valueError()
    }
    if (numbers.some((value) => value <= 0)) {
      return numError()
    }
    const logSum = numbers.reduce((sum, value) => sum + Math.log(value), 0)
    return numberResult(Math.exp(logSum / numbers.length))
  },
  HARMEAN: (...args) => {
    const error = firstError(args)
    if (error) return error
    const numbers: number[] = []
    for (const arg of args) {
      const numeric = toDirectAggregateNumber(arg)
      if (numeric === undefined) {
        return valueError()
      }
      numbers.push(numeric)
    }
    if (numbers.length === 0 || numbers.some((value) => value <= 0)) {
      return numbers.length === 0 ? valueError() : numError()
    }
    return numberResult(numbers.length / numbers.reduce((sum, value) => sum + 1 / value, 0))
  },
  ...mathBuiltins,
  ...fixedIncomeBuiltins,
  RANDBETWEEN: (bottomArg, topArg) => {
    const error = firstError([bottomArg, topArg])
    if (error) {
      return error
    }
    const bottom = integerValue(bottomArg)
    const top = integerValue(topArg)
    if (bottom === undefined || top === undefined || top < bottom) {
      return valueError()
    }
    return numberResult(Math.floor(Math.random() * (top - bottom + 1)) + bottom)
  },
  RANDARRAY: (rowsArg, colsArg, minArg, maxArg, wholeArg) => {
    const error = firstError([rowsArg, colsArg, minArg, maxArg, wholeArg])
    if (error) {
      return error
    }
    const rows = coercePositiveInteger(rowsArg, 1)
    const cols = coercePositiveInteger(colsArg, 1)
    const min = coerceNumber(minArg, 0)
    const max = coerceNumber(maxArg, 1)
    const whole = wholeArg === undefined ? false : (toNumber(wholeArg) ?? 0) !== 0
    if (rows === undefined || cols === undefined || min === undefined || max === undefined || max < min) {
      return valueError()
    }
    const values: CellValue[] = []
    for (let index = 0; index < rows * cols; index += 1) {
      const value = whole
        ? Math.floor(Math.random() * (Math.trunc(max) - Math.trunc(min) + 1)) + Math.trunc(min)
        : Math.random() * (max - min) + min
      values.push(numberResult(value))
    }
    return { kind: 'array', rows, cols, values }
  },
  MUNIT: (sizeArg) => {
    if (sizeArg?.tag === ValueTag.Error) {
      return sizeArg
    }
    const size = positiveIntegerValue(sizeArg)
    return size === undefined ? valueError() : buildIdentityMatrix(size, numberResult)
  },
  SERIESSUM: (xArg, nArg, mArg, ...coefficientArgs) => {
    const error = firstError([xArg, nArg, mArg, ...coefficientArgs])
    if (error) {
      return error
    }
    const x = toNumber(xArg)
    const n = integerValue(nArg)
    const m = integerValue(mArg)
    if (x === undefined || n === undefined || m === undefined) {
      return valueError()
    }
    let sum = 0
    coefficientArgs.forEach((coefficientArg, index) => {
      const coefficient = toNumber(coefficientArg) ?? 0
      sum += coefficient * x ** (n + index * m)
    })
    return numberResult(sum)
  },
  SQRTPI: (value) => {
    const error = firstError([value])
    if (error) {
      return error
    }
    const numeric = toScalarMathNumber(value)
    if (numeric === undefined) {
      return valueError()
    }
    const result = Math.sqrt(numeric * Math.PI)
    return numeric < 0 || !Number.isFinite(result) ? numError() : numberResult(result)
  },
  SUMSQ: (...args) => {
    const error = firstError(args)
    if (error) return error
    return numberResult(collectNumericArgs(args, toNumber).reduce((sum, value) => sum + value ** 2, 0))
  },
  CONVERT: (numberArg, fromUnitArg, toUnitArg) => convertBuiltin(numberArg, fromUnitArg, toUnitArg),
  EUROCONVERT: (numberArg, sourceArg, targetArg, fullPrecisionArg, triangulationPrecisionArg) =>
    euroconvertBuiltin(numberArg, sourceArg, targetArg, fullPrecisionArg, triangulationPrecisionArg),
  ...radixBuiltins,
  ...complexBuiltins,
  T: (...args) => {
    if (args.length !== 1) {
      return valueError()
    }
    const value = args[0]!
    if (value.tag === ValueTag.Error) {
      return value
    }
    return value.tag === ValueTag.String ? value : { tag: ValueTag.String, value: '', stringId: 0 }
  },
  ISOMITTED: (...args) => {
    if (args.length !== 1) {
      return valueError()
    }
    return { tag: ValueTag.Boolean, value: false }
  },
  N: (...args) => {
    if (args.length !== 1) {
      return valueError()
    }
    const value = args[0]!
    if (value.tag === ValueTag.Error) {
      return value
    }
    return numberResult(toNumber(value) ?? 0)
  },
  TYPE: (...args) => {
    if (args.length !== 1) {
      return valueError()
    }
    const value = args[0]!
    if (value.tag === ValueTag.Error) {
      return numberResult(16)
    }
    if ((value as EvaluationResult & { kind?: string }).kind === 'array') {
      return numberResult(64)
    }
    switch (value.tag) {
      case ValueTag.Number:
      case ValueTag.Empty:
        return numberResult(1)
      case ValueTag.String:
        return numberResult(2)
      case ValueTag.Boolean:
        return numberResult(4)
    }
  },
  DELTA: (leftArg, rightArg = { tag: ValueTag.Number, value: 0 }) => {
    const error = firstError([leftArg, rightArg])
    if (error) {
      return error
    }
    const left = toScalarMathNumber(leftArg)
    const right = toScalarMathNumber(rightArg)
    if (left === undefined || right === undefined) {
      return valueError()
    }
    return numberResult(left === right ? 1 : 0)
  },
  GESTEP: (numberArg, stepArg = { tag: ValueTag.Number, value: 0 }) => {
    const error = firstError([numberArg, stepArg])
    if (error) {
      return error
    }
    const numberValue = toScalarMathNumber(numberArg)
    const stepValue = toScalarMathNumber(stepArg)
    if (numberValue === undefined || stepValue === undefined) {
      return valueError()
    }
    return numberResult(numberValue >= stepValue ? 1 : 0)
  },
  ...statisticalBuiltins,
  ...financialBuiltins,
  PERMUT: (numberArg, chosenArg) => {
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
    if (!Number.isFinite(numberRaw) || !Number.isFinite(chosenRaw) || numberValue <= 0 || chosenValue < 0 || chosenValue > numberValue) {
      return numError()
    }
    const result = permutationValue(numberValue, chosenValue)
    return result === undefined || !Number.isFinite(result) ? numError() : numberResult(result)
  },
  PERMUTATIONA: (numberArg, chosenArg) => {
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
    if (
      !Number.isFinite(numberRaw) ||
      !Number.isFinite(chosenRaw) ||
      numberValue < 0 ||
      chosenValue < 0 ||
      (numberValue === 0 && chosenValue > 0)
    ) {
      return numError()
    }
    const result = numberValue ** chosenValue
    return Number.isFinite(result) ? numberResult(result) : numError()
  },
  ...distributionBuiltins,
  SUBTOTAL: (functionNumArg, ...args) => {
    const functionNum = integerValue(functionNumArg)
    return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args, { propagateErrors: true })
  },
  AGGREGATE: (functionNumArg, optionsArg, ...args) => {
    const functionNum = integerValue(functionNumArg)
    const options = integerValue(optionsArg)
    if (functionNum === undefined || options === undefined || options < 0 || options > 7) {
      return valueError()
    }
    const ignoreErrors = aggregateOptionIgnoresErrors(options)
    const values = ignoreErrors ? args.filter((value) => value.tag !== ValueTag.Error) : args
    return aggregateByCode(functionNum, values, { propagateErrors: !ignoreErrors })
  },
  SEQUENCE: (...args) => sequenceResult(args[0], args[1], args[2], args[3]),
  ...externalScalarBuiltins,
  ...scalarPlaceholderBuiltins,
}

const builtins: Record<string, Builtin> = {
  ...scalarBuiltins,
  ...logicalBuiltins,
  ...textBuiltins,
  ...datetimeBuiltins,
}

const dateSystemBuiltinCache = new Map<ExcelDateSystem, Record<string, Builtin>>()

function dateSystemBuiltins(dateSystem: ExcelDateSystem): Record<string, Builtin> {
  let cached = dateSystemBuiltinCache.get(dateSystem)
  if (!cached) {
    cached = {
      ...createTextBuiltins({ dateSystem }),
      ...createDateTimeBuiltins(dateSystem),
    }
    dateSystemBuiltinCache.set(dateSystem, cached)
  }
  return cached
}

const builtinIdByName = new Map(BUILTINS.map((builtin) => [builtin.name.toUpperCase(), builtin.id]))
builtinIdByName.set('USE.THE.COUNTIF', BuiltinId.Countif)
builtinIdByName.set('FORECAST.LINEAR', BuiltinId.Forecast)

export function getBuiltin(name: string): Builtin | undefined {
  const normalized = normalizeBuiltinLookupName(name)
  return builtins[normalized] ?? getExternalScalarFunction(normalized)
}

export function getDateSystemBuiltin(name: string, dateSystem: ExcelDateSystem): Builtin | undefined {
  const normalized = normalizeBuiltinLookupName(name)
  return dateSystem === '1900' ? getBuiltin(normalized) : (dateSystemBuiltins(dateSystem)[normalized] ?? getBuiltin(normalized))
}

export function hasBuiltin(name: string): boolean {
  const upper = normalizeBuiltinLookupName(name)
  return (
    builtins[upper] !== undefined ||
    lookupBuiltins[upper] !== undefined ||
    builtinIdByName.has(upper) ||
    builtinJsSpecialNames.has(upper) ||
    hasExternalFunction(upper)
  )
}

export function getBuiltinId(name: string): BuiltinId | undefined {
  return builtinIdByName.get(normalizeBuiltinLookupName(name))
}
