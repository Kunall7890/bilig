import { ErrorCode, ValueTag } from './protocol'

const DIRECT_AGGREGATE_OP_SUM: u8 = 1
const DIRECT_AGGREGATE_OP_AVERAGE: u8 = 2
const DIRECT_AGGREGATE_OP_COUNT: u8 = 3
const DIRECT_AGGREGATE_OP_MIN: u8 = 4
const DIRECT_AGGREGATE_OP_MAX: u8 = 5
const VALUE_TAG_NUMBER: u8 = <u8>ValueTag.Number
const VALUE_TAG_BOOLEAN: u8 = <u8>ValueTag.Boolean
const VALUE_TAG_EMPTY: u8 = <u8>ValueTag.Empty
const VALUE_TAG_ERROR: u8 = <u8>ValueTag.Error
const ERROR_CODE_NONE: u16 = <u16>ErrorCode.None
const ERROR_CODE_DIV0: u16 = <u16>ErrorCode.Div0
const ERROR_CODE_VALUE: u16 = <u16>ErrorCode.Value

export function evalDenseNumericRowAggregateBatch(
  aggregateKind: u8,
  values: Float64Array,
  rowCount: i32,
  prefixColCount: i32,
  startColOffset: i32,
  aggregateColCount: i32,
  resultOffset: f64,
  outNumbers: Float64Array,
): void {
  if (rowCount <= 0 || prefixColCount <= 0 || aggregateColCount <= 0) {
    return
  }
  if (startColOffset < 0 || startColOffset + aggregateColCount > prefixColCount) {
    return
  }

  for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
    const baseOffset = rowOffset * prefixColCount + startColOffset
    let sum: f64 = 0
    let minimum: f64 = Infinity
    let maximum: f64 = -Infinity

    for (let colOffset = 0; colOffset < aggregateColCount; colOffset++) {
      const value = values[baseOffset + colOffset]
      sum += value
      if (value < minimum) {
        minimum = value
      }
      if (value > maximum) {
        maximum = value
      }
    }

    if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
      outNumbers[rowOffset] = sum + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
      outNumbers[rowOffset] = sum / aggregateColCount + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_COUNT) {
      outNumbers[rowOffset] = aggregateColCount + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
      outNumbers[rowOffset] = minimum + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
      outNumbers[rowOffset] = maximum + resultOffset
    } else {
      outNumbers[rowOffset] = NaN
    }
  }
}

export function evalAnchoredPrefixAggregateBatch(
  aggregateKind: u8,
  tags: Uint8Array,
  numbers: Float64Array,
  errors: Uint16Array,
  rowCount: i32,
  colCount: i32,
  formulaRowEnds: Uint32Array,
  resultOffsets: Float64Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const formulaCount = formulaRowEnds.length
  if (rowCount <= 0 || colCount <= 0 || formulaCount <= 0) {
    return
  }
  if (tags.length < rowCount * colCount || numbers.length < rowCount * colCount || errors.length < rowCount * colCount) {
    return
  }
  if (
    resultOffsets.length < formulaCount ||
    outTags.length < formulaCount ||
    outNumbers.length < formulaCount ||
    outErrors.length < formulaCount
  ) {
    return
  }

  let sum: f64 = 0
  let count: i32 = 0
  let averageCount: i32 = 0
  let errorCode: u16 = ERROR_CODE_NONE
  let errorCount: i32 = 0
  let minimum: f64 = Infinity
  let maximum: f64 = -Infinity
  let formulaIndex = 0

  for (let row = 0; row < rowCount && formulaIndex < formulaCount; row++) {
    const baseOffset = row * colCount
    for (let col = 0; col < colCount; col++) {
      const valueOffset = baseOffset + col
      const tag = tags[valueOffset]
      if (tag == VALUE_TAG_NUMBER) {
        const numeric = numbers[valueOffset]
        sum += numeric
        count += 1
        averageCount += 1
        if (numeric < minimum) {
          minimum = numeric
        }
        if (numeric > maximum) {
          maximum = numeric
        }
      } else if (tag == VALUE_TAG_BOOLEAN) {
        const numeric: f64 = numbers[valueOffset] != 0 ? 1 : 0
        sum += numeric
        count += 1
        averageCount += 1
        if (numeric < minimum) {
          minimum = numeric
        }
        if (numeric > maximum) {
          maximum = numeric
        }
      } else if (tag == VALUE_TAG_EMPTY) {
        if (0 < minimum) {
          minimum = 0
        }
        if (0 > maximum) {
          maximum = 0
        }
      } else if (tag == VALUE_TAG_ERROR) {
        if (errorCode == ERROR_CODE_NONE) {
          errorCode = errors[valueOffset]
        }
        errorCount += 1
      }
    }

    while (formulaIndex < formulaCount && formulaRowEnds[formulaIndex] <= <u32>row) {
      const resultOffset = resultOffsets[formulaIndex]
      if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
        if (errorCount > 0 && errorCode != ERROR_CODE_NONE) {
          outTags[formulaIndex] = VALUE_TAG_ERROR
          outErrors[formulaIndex] = errorCode
        } else {
          outTags[formulaIndex] = VALUE_TAG_NUMBER
          outNumbers[formulaIndex] = sum + resultOffset
          outErrors[formulaIndex] = ERROR_CODE_NONE
        }
      } else if (aggregateKind == DIRECT_AGGREGATE_OP_COUNT) {
        outTags[formulaIndex] = VALUE_TAG_NUMBER
        outNumbers[formulaIndex] = <f64>count + resultOffset
        outErrors[formulaIndex] = ERROR_CODE_NONE
      } else if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
        if (errorCount > 0 && errorCode != ERROR_CODE_NONE) {
          outTags[formulaIndex] = VALUE_TAG_ERROR
          outErrors[formulaIndex] = errorCode
        } else if (averageCount == 0) {
          outTags[formulaIndex] = VALUE_TAG_ERROR
          outErrors[formulaIndex] = ERROR_CODE_DIV0
        } else {
          outTags[formulaIndex] = VALUE_TAG_NUMBER
          outNumbers[formulaIndex] = sum / <f64>averageCount + resultOffset
          outErrors[formulaIndex] = ERROR_CODE_NONE
        }
      } else if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
        outTags[formulaIndex] = VALUE_TAG_NUMBER
        outNumbers[formulaIndex] = (minimum == Infinity ? 0 : minimum) + resultOffset
        outErrors[formulaIndex] = ERROR_CODE_NONE
      } else if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
        outTags[formulaIndex] = VALUE_TAG_NUMBER
        outNumbers[formulaIndex] = (maximum == -Infinity ? 0 : maximum) + resultOffset
        outErrors[formulaIndex] = ERROR_CODE_NONE
      } else {
        outTags[formulaIndex] = VALUE_TAG_ERROR
        outErrors[formulaIndex] = ERROR_CODE_VALUE
      }
      formulaIndex += 1
    }
  }
}
