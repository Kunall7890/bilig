import { ErrorCode, ValueTag } from './protocol'

const DIRECT_AGGREGATE_OP_SUM: u8 = 1
const DIRECT_AGGREGATE_OP_AVERAGE: u8 = 2
const DIRECT_AGGREGATE_OP_COUNT: u8 = 3
const DIRECT_AGGREGATE_OP_MIN: u8 = 4
const DIRECT_AGGREGATE_OP_MAX: u8 = 5

function writeNumber(index: i32, value: f64, outTags: Uint8Array, outNumbers: Float64Array, outErrors: Uint16Array): void {
  outTags[index] = <u8>ValueTag.Number
  outNumbers[index] = value
  outErrors[index] = ErrorCode.None
}

function writeError(index: i32, code: u16, outTags: Uint8Array, outNumbers: Float64Array, outErrors: Uint16Array): void {
  outTags[index] = <u8>ValueTag.Error
  outNumbers[index] = 0
  outErrors[index] = code
}

export function evalDirectCriteriaMatchedAggregateBatch(
  aggregateKinds: Uint8Array,
  matchStarts: Uint32Array,
  matchLengths: Uint32Array,
  matchedRows: Uint32Array,
  aggregateTags: Uint8Array,
  aggregateNumbers: Float64Array,
  aggregateErrors: Uint16Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const aggregateLength = <u32>aggregateTags.length
  for (let resultIndex = 0; resultIndex < aggregateKinds.length; resultIndex++) {
    const aggregateKind = aggregateKinds[resultIndex]
    const matchStart = matchStarts[resultIndex]
    const matchLength = matchLengths[resultIndex]
    const matchEnd = matchStart + matchLength

    if (aggregateKind == DIRECT_AGGREGATE_OP_COUNT) {
      writeNumber(resultIndex, <f64>matchLength, outTags, outNumbers, outErrors)
      continue
    }

    let sum: f64 = 0
    let count: u32 = 0
    let minimum: f64 = Infinity
    let maximum: f64 = -Infinity
    let hasResult = true

    for (let matchCursor = matchStart; matchCursor < matchEnd; matchCursor++) {
      if (matchCursor >= <u32>matchedRows.length) {
        writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
        hasResult = false
        break
      }

      const rowOffset = matchedRows[matchCursor]
      if (rowOffset >= aggregateLength) {
        writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
        hasResult = false
        break
      }

      const tag = aggregateTags[rowOffset]
      if (tag == ValueTag.Error) {
        writeError(resultIndex, aggregateErrors[rowOffset], outTags, outNumbers, outErrors)
        hasResult = false
        break
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
        if (tag == ValueTag.Number) {
          sum += aggregateNumbers[rowOffset]
        } else if (tag == ValueTag.Boolean) {
          sum += aggregateNumbers[rowOffset] != 0 ? 1 : 0
        }
        continue
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
        if (tag == ValueTag.Number) {
          sum += aggregateNumbers[rowOffset]
          count += 1
        } else if (tag == ValueTag.Boolean) {
          sum += aggregateNumbers[rowOffset] != 0 ? 1 : 0
          count += 1
        } else if (tag == ValueTag.Empty) {
          count += 1
        }
        continue
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
        if (tag == ValueTag.Number && aggregateNumbers[rowOffset] < minimum) {
          minimum = aggregateNumbers[rowOffset]
        }
        continue
      }

      if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
        if (tag == ValueTag.Number && aggregateNumbers[rowOffset] > maximum) {
          maximum = aggregateNumbers[rowOffset]
        }
        continue
      }

      writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
      hasResult = false
      break
    }

    if (!hasResult) {
      continue
    }
    if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
      writeNumber(resultIndex, sum, outTags, outNumbers, outErrors)
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
      if (count == 0) {
        writeError(resultIndex, <u16>ErrorCode.Div0, outTags, outNumbers, outErrors)
      } else {
        writeNumber(resultIndex, sum / <f64>count, outTags, outNumbers, outErrors)
      }
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
      writeNumber(resultIndex, minimum == Infinity ? 0 : minimum, outTags, outNumbers, outErrors)
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
      writeNumber(resultIndex, maximum == -Infinity ? 0 : maximum, outTags, outNumbers, outErrors)
    } else {
      writeError(resultIndex, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
    }
  }
}
