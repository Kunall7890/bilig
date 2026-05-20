import type { RawKernelExports } from './raw-kernel-exports.js'

type TypedArrayValue = Uint8Array | Uint16Array | Uint32Array | Float64Array

const ARRAY_BUFFER_CLASS_ID = 1
const UINT8_ARRAY_CLASS_ID = 4
const FLOAT64_ARRAY_CLASS_ID = 5
const UINT16_ARRAY_CLASS_ID = 6
const UINT32_ARRAY_CLASS_ID = 7

interface LoweredArraySpec<T extends TypedArrayValue> {
  align: number
  classId: number
  ctor: {
    new (buffer: ArrayBufferLike, byteOffset: number, length: number): T
  }
}

interface LoweringContext {
  dataView: DataView
}

const uint8Spec: LoweredArraySpec<Uint8Array> = {
  align: 0,
  classId: UINT8_ARRAY_CLASS_ID,
  ctor: Uint8Array,
}

const uint16Spec: LoweredArraySpec<Uint16Array> = {
  align: 1,
  classId: UINT16_ARRAY_CLASS_ID,
  ctor: Uint16Array,
}

const uint32Spec: LoweredArraySpec<Uint32Array> = {
  align: 2,
  classId: UINT32_ARRAY_CLASS_ID,
  ctor: Uint32Array,
}

const float64Spec: LoweredArraySpec<Float64Array> = {
  align: 3,
  classId: FLOAT64_ARRAY_CLASS_ID,
  ctor: Float64Array,
}

function setUint32(raw: RawKernelExports, context: LoweringContext, pointer: number, value: number): void {
  try {
    context.dataView.setUint32(pointer, value, true)
  } catch {
    context.dataView = new DataView(raw.memory.buffer)
    context.dataView.setUint32(pointer, value, true)
  }
}

function getUint32(raw: RawKernelExports, context: LoweringContext, pointer: number): number {
  try {
    return context.dataView.getUint32(pointer, true)
  } catch {
    context.dataView = new DataView(raw.memory.buffer)
    return context.dataView.getUint32(pointer, true)
  }
}

function lowerTypedArray<T extends TypedArrayValue>(
  raw: RawKernelExports,
  context: LoweringContext,
  values: T,
  spec: LoweredArraySpec<T>,
): number {
  const byteLength = values.length << spec.align
  const bufferPtr = raw.__pin(raw.__new(byteLength, ARRAY_BUFFER_CLASS_ID))
  const headerPtr = raw.__pin(raw.__new(12, spec.classId))
  try {
    setUint32(raw, context, headerPtr, bufferPtr)
    setUint32(raw, context, headerPtr + 4, bufferPtr)
    setUint32(raw, context, headerPtr + 8, byteLength)
    new spec.ctor(raw.memory.buffer, bufferPtr, values.length).set(values)
    return headerPtr
  } finally {
    raw.__unpin(bufferPtr)
  }
}

function copyLoweredTypedArray<T extends TypedArrayValue>(
  raw: RawKernelExports,
  context: LoweringContext,
  pointer: number,
  target: T,
  spec: LoweredArraySpec<T>,
): void {
  target.set(new spec.ctor(raw.memory.buffer, getUint32(raw, context, pointer + 4), target.length))
}

export function evalUniformNumericLookupBatchRaw(
  raw: RawKernelExports,
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
  const context: LoweringContext = { dataView: new DataView(raw.memory.buffer) }
  const kindsPtr = lowerTypedArray(raw, context, kinds, uint8Spec)
  const matchModesPtr = lowerTypedArray(raw, context, matchModes, uint8Spec)
  const startsPtr = lowerTypedArray(raw, context, starts, float64Spec)
  const stepsPtr = lowerTypedArray(raw, context, steps, float64Spec)
  const lengthsPtr = lowerTypedArray(raw, context, lengths, uint32Spec)
  const repeatedRunLengthsPtr = lowerTypedArray(raw, context, repeatedRunLengths, uint32Spec)
  const lookupTagsPtr = lowerTypedArray(raw, context, lookupTags, uint8Spec)
  const lookupNumbersPtr = lowerTypedArray(raw, context, lookupNumbers, float64Spec)
  const outTagsPtr = lowerTypedArray(raw, context, outTags, uint8Spec)
  const outNumbersPtr = lowerTypedArray(raw, context, outNumbers, float64Spec)
  const outErrorsPtr = lowerTypedArray(raw, context, outErrors, uint16Spec)
  try {
    raw.evalUniformNumericLookupBatch(
      kindsPtr,
      matchModesPtr,
      startsPtr,
      stepsPtr,
      lengthsPtr,
      repeatedRunLengthsPtr,
      lookupTagsPtr,
      lookupNumbersPtr,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    copyLoweredTypedArray(raw, context, outTagsPtr, outTags, uint8Spec)
    copyLoweredTypedArray(raw, context, outNumbersPtr, outNumbers, float64Spec)
    copyLoweredTypedArray(raw, context, outErrorsPtr, outErrors, uint16Spec)
  } finally {
    raw.__unpin(kindsPtr)
    raw.__unpin(matchModesPtr)
    raw.__unpin(startsPtr)
    raw.__unpin(stepsPtr)
    raw.__unpin(lengthsPtr)
    raw.__unpin(repeatedRunLengthsPtr)
    raw.__unpin(lookupTagsPtr)
    raw.__unpin(lookupNumbersPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}
