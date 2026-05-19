import { ErrorCode, ValueTag } from './protocol'

const DIRECT_LOOKUP_KIND_EXACT_UNIFORM_NUMERIC: u8 = 1
const DIRECT_LOOKUP_KIND_APPROXIMATE_UNIFORM_NUMERIC: u8 = 2
const DIRECT_LOOKUP_MATCH_MODE_ASCENDING: u8 = 1
const DIRECT_LOOKUP_MATCH_MODE_DESCENDING: u8 = 2

function writeDirectLookupNumber(index: i32, value: f64, outTags: Uint8Array, outNumbers: Float64Array, outErrors: Uint16Array): void {
  outTags[index] = <u8>ValueTag.Number
  outNumbers[index] = value
  outErrors[index] = ErrorCode.None
}

function writeDirectLookupError(index: i32, code: u16, outTags: Uint8Array, outNumbers: Float64Array, outErrors: Uint16Array): void {
  outTags[index] = <u8>ValueTag.Error
  outNumbers[index] = 0
  outErrors[index] = code
}

function exactUniformNumericLookupResult(start: f64, step: f64, length: u32, lookupValue: f64): f64 {
  if (length == 0 || step == 0 || !isFinite(start) || !isFinite(step) || !isFinite(lookupValue)) {
    return 0
  }
  const relative = (lookupValue - start) / step
  const nearestOffset = Math.round(relative)
  if (nearestOffset < 0 || nearestOffset >= <f64>length) {
    return 0
  }
  if (start + step * nearestOffset != lookupValue) {
    return 0
  }
  return nearestOffset + 1
}

function approximateRepeatedUniformLookupResult(
  length: u32,
  start: f64,
  step: f64,
  repeatedRunLength: u32,
  matchMode: u8,
  lookupValue: f64,
): f64 {
  if (repeatedRunLength == 0 || length == 0 || step == 0) {
    return 0
  }
  const groupCount = Math.ceil(<f64>length / <f64>repeatedRunLength)
  const lastValue = start + step * (groupCount - 1)
  if (matchMode == DIRECT_LOOKUP_MATCH_MODE_ASCENDING && step > 0) {
    if (lookupValue < start) {
      return 0
    }
    if (lookupValue >= lastValue) {
      return <f64>length
    }
    const group = Math.floor((lookupValue - start) / step)
    const position = (group + 1) * repeatedRunLength
    return position >= <f64>length ? <f64>length : position
  }
  if (matchMode == DIRECT_LOOKUP_MATCH_MODE_DESCENDING && step < 0) {
    if (lookupValue > start) {
      return 0
    }
    if (lookupValue <= lastValue) {
      return <f64>length
    }
    const group = Math.floor((start - lookupValue) / -step)
    const position = (group + 1) * repeatedRunLength
    return position >= <f64>length ? <f64>length : position
  }
  return 0
}

function approximateUniformNumericLookupResult(
  start: f64,
  step: f64,
  length: u32,
  repeatedRunLength: u32,
  matchMode: u8,
  lookupValue: f64,
): f64 {
  if (length == 0 || step == 0 || !isFinite(start) || !isFinite(step) || !isFinite(lookupValue)) {
    return 0
  }
  if (repeatedRunLength != 0) {
    return approximateRepeatedUniformLookupResult(length, start, step, repeatedRunLength, matchMode, lookupValue)
  }
  const lastValue = start + step * (<f64>length - 1)
  if (matchMode == DIRECT_LOOKUP_MATCH_MODE_ASCENDING && step > 0) {
    if (lookupValue < start) {
      return 0
    }
    if (lookupValue >= lastValue) {
      return <f64>length
    }
    const position = step == 1 ? Math.floor(lookupValue - start) + 1 : Math.floor((lookupValue - start) / step) + 1
    return Math.min(<f64>length, Math.max(1, position))
  }
  if (matchMode == DIRECT_LOOKUP_MATCH_MODE_DESCENDING && step < 0) {
    if (lookupValue > start) {
      return 0
    }
    if (lookupValue <= lastValue) {
      return <f64>length
    }
    const position = step == -1 ? Math.floor(start - lookupValue) + 1 : Math.floor((start - lookupValue) / -step) + 1
    return Math.min(<f64>length, Math.max(1, position))
  }
  return 0
}

function lookupValueForApproximate(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Boolean) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return value
}

export function evalUniformNumericLookupBatch(
  kinds: Uint8Array,
  matchModes: Uint8Array,
  starts: Float64Array,
  steps: Float64Array,
  lengths: Uint32Array,
  repeatedRunLengths: Uint32Array,
  lookupTags: Uint8Array,
  lookupNumbers: Float64Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  for (let index = 0; index < kinds.length; index++) {
    const kind = kinds[index]
    let position: f64 = 0
    if (kind == DIRECT_LOOKUP_KIND_EXACT_UNIFORM_NUMERIC) {
      if (lookupTags[index] != ValueTag.Number) {
        writeDirectLookupError(index, <u16>ErrorCode.NA, outTags, outNumbers, outErrors)
        continue
      }
      position = exactUniformNumericLookupResult(starts[index], steps[index], lengths[index], lookupNumbers[index])
    } else if (kind == DIRECT_LOOKUP_KIND_APPROXIMATE_UNIFORM_NUMERIC) {
      position = approximateUniformNumericLookupResult(
        starts[index],
        steps[index],
        lengths[index],
        repeatedRunLengths[index],
        matchModes[index],
        lookupValueForApproximate(lookupTags[index], lookupNumbers[index]),
      )
    } else {
      writeDirectLookupError(index, <u16>ErrorCode.Value, outTags, outNumbers, outErrors)
      continue
    }
    if (position == 0) {
      writeDirectLookupError(index, <u16>ErrorCode.NA, outTags, outNumbers, outErrors)
      continue
    }
    writeDirectLookupNumber(index, position, outTags, outNumbers, outErrors)
  }
}
