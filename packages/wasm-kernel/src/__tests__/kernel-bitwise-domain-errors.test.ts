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

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBe(expected)
}

function expectErrorCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

describe('wasm kernel bitwise error semantics', () => {
  it('uses full 48-bit bitwise semantics and supports negative shift reversal', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 0, 24, 1, 1)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const programs = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitand, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitxor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitlshift, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitrshift, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitlshift, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitrshift, 2), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 7 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)

    const high = 2 ** 40
    const constants = packConstants([
      [high + 3, high + 1],
      [high + 3, high + 5],
      [high + 3, 5],
      [high, 4],
      [high, 4],
      [8, -1],
      [8, -1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    expectNumberCell(kernel, outputs[0], high + 1)
    expectNumberCell(kernel, outputs[1], high + 7)
    expectNumberCell(kernel, outputs[2], high + 6)
    expectNumberCell(kernel, outputs[3], 2 ** 44)
    expectNumberCell(kernel, outputs[4], 2 ** 36)
    expectNumberCell(kernel, outputs[5], 4)
    expectNumberCell(kernel, outputs[6], 16)
  })

  it('separates bitwise numeric-domain errors from nonnumeric #VALUE! coercions', async () => {
    const kernel = await createKernel()
    const width = 20
    kernel.init(64, 1, 32, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([3]), Uint16Array.from(Array.from('bad', (char) => char.charCodeAt(0))))
    kernel.writeCells(new Uint8Array(64), new Float64Array(64), new Uint32Array(64), new Uint16Array(64))

    const programs = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitand, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitxor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitlshift, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitrshift, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitlshift, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitrshift, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Bitand, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodeCall(BuiltinId.Bitor, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Bitxor, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodeCall(BuiltinId.Bitlshift, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Bitrshift, 2), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 12 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)
    const constants = packConstants([[-1, 1], [2 ** 48, 1], [1.5, 1], [-1, 1], [2 ** 48, 1], [4, 54], [4, -54], [1], [1], [1], [4], [1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    for (let index = 0; index < 7; index += 1) {
      expectErrorCell(kernel, outputs[index], ErrorCode.Num)
    }
    for (let index = 7; index < outputs.length; index += 1) {
      expectErrorCell(kernel, outputs[index], ErrorCode.Value)
    }
  })

  it('rejects malformed variadic bitwise programs', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 0, 12, 1, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const programs = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Bitand, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Bitor, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Bitxor, 3), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 3 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)
    const constants = packConstants([
      [6, 3, 1],
      [6, 3, 1],
      [6, 3, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    for (const output of outputs) {
      expectErrorCell(kernel, output, ErrorCode.Value)
    }
  })
})
