import type { RawKernelExports } from './raw-kernel-exports.js'

type TypedArrayValue = Uint8Array | Uint32Array

const ARRAY_BUFFER_CLASS_ID = 1
const UINT8_ARRAY_CLASS_ID = 4
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

const uint32Spec: LoweredArraySpec<Uint32Array> = {
  align: 2,
  classId: UINT32_ARRAY_CLASS_ID,
  ctor: Uint32Array,
}

function setUint32(raw: RawKernelExports, context: LoweringContext, pointer: number, value: number): void {
  try {
    context.dataView.setUint32(pointer, value, true)
  } catch {
    context.dataView = new DataView(raw.memory.buffer)
    context.dataView.setUint32(pointer, value, true)
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

export function materializePivotTableRaw(
  raw: RawKernelExports,
  sourceRangeIndex: number,
  sourceWidth: number,
  groupByColumnIndices: Uint32Array,
  valueColumnIndices: Uint32Array,
  valueAggregations: Uint8Array,
): void {
  const context: LoweringContext = { dataView: new DataView(raw.memory.buffer) }
  const groupByPtr = lowerTypedArray(raw, context, groupByColumnIndices, uint32Spec)
  const valueColsPtr = lowerTypedArray(raw, context, valueColumnIndices, uint32Spec)
  const valueAggsPtr = lowerTypedArray(raw, context, valueAggregations, uint8Spec)
  try {
    raw.materializePivotTable(
      sourceRangeIndex,
      sourceWidth,
      groupByColumnIndices.length,
      groupByPtr,
      valueColumnIndices.length,
      valueColsPtr,
      valueAggsPtr,
    )
  } finally {
    raw.__unpin(groupByPtr)
    raw.__unpin(valueColsPtr)
    raw.__unpin(valueAggsPtr)
  }
}
