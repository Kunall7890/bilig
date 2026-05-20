import type { RawKernelExports } from './raw-kernel-exports.js'

const ARRAY_BUFFER_CLASS_ID = 1
const UINT32_ARRAY_CLASS_ID = 7

interface LoweringContext {
  dataView: DataView
}

function setUint32(raw: RawKernelExports, context: LoweringContext, pointer: number, value: number): void {
  try {
    context.dataView.setUint32(pointer, value, true)
  } catch {
    context.dataView = new DataView(raw.memory.buffer)
    context.dataView.setUint32(pointer, value, true)
  }
}

function lowerUint32Array(raw: RawKernelExports, context: LoweringContext, values: Uint32Array): number {
  const byteLength = values.length << 2
  const bufferPtr = raw.__pin(raw.__new(byteLength, ARRAY_BUFFER_CLASS_ID))
  const headerPtr = raw.__pin(raw.__new(12, UINT32_ARRAY_CLASS_ID))
  try {
    setUint32(raw, context, headerPtr, bufferPtr)
    setUint32(raw, context, headerPtr + 4, bufferPtr)
    setUint32(raw, context, headerPtr + 8, byteLength)
    new Uint32Array(raw.memory.buffer, bufferPtr, values.length).set(values)
    return headerPtr
  } finally {
    raw.__unpin(bufferPtr)
  }
}

export function evalBatchRaw(raw: RawKernelExports, cellIndices: Uint32Array): void {
  const context: LoweringContext = { dataView: new DataView(raw.memory.buffer) }
  const cellIndicesPtr = lowerUint32Array(raw, context, cellIndices)
  try {
    raw.evalBatch(cellIndicesPtr)
  } finally {
    raw.__unpin(cellIndicesPtr)
  }
}
