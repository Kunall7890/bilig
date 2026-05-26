import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
}

function encodePushError(code: ErrorCode): number {
  return (Opcode.PushError << 24) | code
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

function expectErrorCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

describe('wasm kernel text Unicode error semantics', () => {
  it('propagates scalar text conversion input errors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 0, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushError(ErrorCode.Ref), encodeCall(BuiltinId.Char, 1), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Code, 1), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodeCall(BuiltinId.Unicode, 1), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Unichar, 1), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 4 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputs)
    const constants = packConstants([[], [], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    expectErrorCell(kernel, outputs[0], ErrorCode.Ref)
    expectErrorCell(kernel, outputs[1], ErrorCode.Name)
    expectErrorCell(kernel, outputs[2], ErrorCode.Ref)
    expectErrorCell(kernel, outputs[3], ErrorCode.Name)
  })

  it('matches documented UNICODE and UNICHAR domain errors', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(24, 2, 10, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 1]), Uint32Array.from([1, 1]), Uint16Array.from([0xd800, 0xdc00]))
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Unicode, 1), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Unicode, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Unichar, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Unichar, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Unichar, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Unichar, 1), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputs)
    const constants = packConstants([[], [], [0], [0xd800], [0xdfff], [0x110000]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    expectErrorCell(kernel, outputs[0], ErrorCode.Value)
    expectErrorCell(kernel, outputs[1], ErrorCode.Value)
    expectErrorCell(kernel, outputs[2], ErrorCode.Value)
    expectErrorCell(kernel, outputs[3], ErrorCode.NA)
    expectErrorCell(kernel, outputs[4], ErrorCode.NA)
    expectErrorCell(kernel, outputs[5], ErrorCode.Value)
  })
})
