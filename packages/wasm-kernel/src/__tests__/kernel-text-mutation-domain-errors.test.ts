import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

const OUTPUT_STRING_BASE = 2147483648
const maxExcelCellTextLength = 32_767

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
}

function encodeRet(): number {
  return Opcode.Ret << 24
}

function packPrograms(programs: number[][]): {
  programs: Uint32Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0
  for (const program of programs) {
    offsets.push(offset)
    lengths.push(program.length)
    flat.push(...program)
    offset += program.length
  }
  return {
    programs: Uint32Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function packConstants(constantsByProgram: number[][]): {
  constants: Float64Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0
  for (const constants of constantsByProgram) {
    offsets.push(offset)
    lengths.push(constants.length)
    flat.push(...constants)
    offset += constants.length
  }
  return {
    constants: Float64Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

function readStringCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, pooledStrings: readonly string[]): string {
  expect(kernel.readTags()[index]).toBe(ValueTag.String)
  const raw = kernel.readStringIds()[index] ?? 0
  const outputIndex = raw >= OUTPUT_STRING_BASE ? raw - OUTPUT_STRING_BASE : -1
  return outputIndex >= 0 ? (kernel.readOutputStrings()[outputIndex] ?? '') : (pooledStrings[raw] ?? '')
}

function expectErrorCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

describe('wasm kernel REPT domain errors', () => {
  it('returns #VALUE when repeated text would exceed Excel cell text length', async () => {
    const kernel = await createKernel()
    const width = 8
    const pooledStrings = ['x', 'xx']
    kernel.init(32, pooledStrings.length, 3, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 1]), Uint32Array.from([1, 2]), Uint16Array.from([120, 120, 120]))
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Rept, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Rept, 2), encodeRet()],
      [encodePushString(1), encodePushNumber(0), encodeCall(BuiltinId.Rept, 2), encodeRet()],
    ])
    const targetCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targetCells)
    const constants = packConstants([[maxExcelCellTextLength], [maxExcelCellTextLength + 1], [16_384]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(targetCells)

    expect(readStringCell(kernel, cellIndex(1, 0, width), pooledStrings)).toHaveLength(maxExcelCellTextLength)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Value)
  })
})
