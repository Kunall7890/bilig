import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushBoolean(value: boolean): number {
  return (Opcode.PushBoolean << 24) | (value ? 1 : 0)
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

describe('wasm kernel extended distribution error semantics', () => {
  it('returns #NUM for documented extended distribution numeric-domain errors', async () => {
    const kernel = await createKernel()
    const width = 20
    kernel.init(64, 2, 64, 64, 1)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([1]), Uint16Array.from([50]))
    kernel.writeCells(new Uint8Array(64), new Float64Array(64), new Uint32Array(64), new Uint16Array(64))

    const programs = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Chidist, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.ChisqInv, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.BetaDist, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.BetaInv, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushBoolean(true), encodeCall(BuiltinId.FDist, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.FInv, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BuiltinId.TDist, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TDist2T, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Tdist, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TInv, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.BinomDist, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.BinomDistRange, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Critbinom, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushBoolean(true),
        encodeCall(BuiltinId.HypgeomDist, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Negbinomdist, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.NegbinomDist, 4),
        encodeRet(),
      ],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 16 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)

    const constants = packConstants([
      [-1, 1],
      [-0.1, 1],
      [0.5, 0, 1],
      [-0.1, 1, 1],
      [-1, 1, 1],
      [-0.1, 1, 1],
      [1, 0],
      [-1, 1],
      [1, 1, 3],
      [-0.1, 1],
      [4, 3, 0.5],
      [3, -0.1, 1],
      [3, 0.5, 0],
      [1, 11, 3, 10],
      [-1, 3, 0.5],
      [1, 0, 0.5],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    for (const output of outputs) {
      expect(kernel.readTags()[output]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[output]).toBe(ErrorCode.Num)
    }
  })

  it('keeps nonnumeric extended distribution arguments as #VALUE', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(16, 2, 8, 8, 1)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([1]), Uint16Array.from([50]))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const programs = packPrograms([
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Chidist, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.BetaInv, 3), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.FInv, 3), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.TInv, 2), encodeRet()],
      [
        encodePushString(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushBoolean(true),
        encodeCall(BuiltinId.BinomDist, 4),
        encodeRet(),
      ],
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Negbinomdist, 3), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)

    const constants = packConstants([[1], [1, 1], [1, 1], [1], [3, 0.5], [3, 0.5]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    for (const output of outputs) {
      expect(kernel.readTags()[output]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[output]).toBe(ErrorCode.Value)
    }
  })
})
