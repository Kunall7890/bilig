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

export function evalDenseNumericRowAggregateBatchRaw(
  raw: RawKernelExports,
  aggregateKind: number,
  values: Float64Array,
  rowCount: number,
  prefixColCount: number,
  startColOffset: number,
  aggregateColCount: number,
  resultOffset: number,
  outNumbers: Float64Array,
): void {
  const context: LoweringContext = { dataView: new DataView(raw.memory.buffer) }
  const valuesPtr = lowerTypedArray(raw, context, values, float64Spec)
  const outNumbersPtr = lowerTypedArray(raw, context, outNumbers, float64Spec)
  try {
    raw.evalDenseNumericRowAggregateBatch(
      aggregateKind,
      valuesPtr,
      rowCount,
      prefixColCount,
      startColOffset,
      aggregateColCount,
      resultOffset,
      outNumbersPtr,
    )
    copyLoweredTypedArray(raw, context, outNumbersPtr, outNumbers, float64Spec)
  } finally {
    raw.__unpin(valuesPtr)
    raw.__unpin(outNumbersPtr)
  }
}

export function evalAnchoredPrefixAggregateBatchRaw(
  raw: RawKernelExports,
  aggregateKind: number,
  tags: Uint8Array,
  numbers: Float64Array,
  errors: Uint16Array,
  rowCount: number,
  colCount: number,
  formulaRowEnds: Uint32Array,
  resultOffsets: Float64Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  const context: LoweringContext = { dataView: new DataView(raw.memory.buffer) }
  const tagsPtr = lowerTypedArray(raw, context, tags, uint8Spec)
  const numbersPtr = lowerTypedArray(raw, context, numbers, float64Spec)
  const errorsPtr = lowerTypedArray(raw, context, errors, uint16Spec)
  const formulaRowEndsPtr = lowerTypedArray(raw, context, formulaRowEnds, uint32Spec)
  const resultOffsetsPtr = lowerTypedArray(raw, context, resultOffsets, float64Spec)
  const outTagsPtr = lowerTypedArray(raw, context, outTags, uint8Spec)
  const outNumbersPtr = lowerTypedArray(raw, context, outNumbers, float64Spec)
  const outErrorsPtr = lowerTypedArray(raw, context, outErrors, uint16Spec)
  try {
    raw.evalAnchoredPrefixAggregateBatch(
      aggregateKind,
      tagsPtr,
      numbersPtr,
      errorsPtr,
      rowCount,
      colCount,
      formulaRowEndsPtr,
      resultOffsetsPtr,
      outTagsPtr,
      outNumbersPtr,
      outErrorsPtr,
    )
    copyLoweredTypedArray(raw, context, outTagsPtr, outTags, uint8Spec)
    copyLoweredTypedArray(raw, context, outNumbersPtr, outNumbers, float64Spec)
    copyLoweredTypedArray(raw, context, outErrorsPtr, outErrors, uint16Spec)
  } finally {
    raw.__unpin(tagsPtr)
    raw.__unpin(numbersPtr)
    raw.__unpin(errorsPtr)
    raw.__unpin(formulaRowEndsPtr)
    raw.__unpin(resultOffsetsPtr)
    raw.__unpin(outTagsPtr)
    raw.__unpin(outNumbersPtr)
    raw.__unpin(outErrorsPtr)
  }
}
