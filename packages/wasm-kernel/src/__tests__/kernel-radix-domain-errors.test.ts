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

function packStrings(strings: readonly string[]): {
  data: Uint16Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const data: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  for (const value of strings) {
    offsets.push(data.length)
    lengths.push(value.length)
    for (let index = 0; index < value.length; index += 1) {
      data.push(value.charCodeAt(index))
    }
  }
  return {
    data: Uint16Array.from(data),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

describe('wasm kernel radix conversion error semantics', () => {
  it('returns #NUM! for signed-radix source, range, and places-domain errors', async () => {
    const kernel = await createKernel()
    const width = 20
    kernel.init(64, 10, 24, 8, 1)
    const strings = packStrings(['102', '11111111111', 'G', '8', '200', '20000000', '1000', '1110', '100'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(64), new Float64Array(64), new Uint32Array(64), new Uint16Array(64))

    const programs = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Bin2dec, 1), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Bin2dec, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Hex2dec, 1), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.Oct2dec, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Bin2hex, 1), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Bin2oct, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Hex2bin, 1), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.Oct2hex, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Dec2bin, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Dec2hex, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Dec2oct, 1), encodeRet()],
      [encodePushString(4), encodeCall(BuiltinId.Hex2bin, 1), encodeRet()],
      [encodePushString(5), encodeCall(BuiltinId.Hex2oct, 1), encodeRet()],
      [encodePushString(6), encodeCall(BuiltinId.Oct2bin, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dec2hex, 2), encodeRet()],
      [encodePushString(7), encodePushNumber(0), encodeCall(BuiltinId.Bin2hex, 2), encodeRet()],
      [encodePushString(8), encodePushNumber(0), encodeCall(BuiltinId.Oct2hex, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dec2bin, 2), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 18 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)
    const constants = packConstants([
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [512],
      [549755813888],
      [536870912],
      [],
      [],
      [],
      [64, 1],
      [0],
      [1],
      [10, -1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    for (const output of outputs) {
      expect(kernel.readTags()[output]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[output]).toBe(ErrorCode.Num)
    }
  })

  it('keeps nonnumeric signed-radix decimal inputs and places as #VALUE!', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 8, 4, 1)
    const strings = packStrings(['bad', '1010', 'F'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const programs = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Dec2bin, 1), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodeCall(BuiltinId.Dec2bin, 2), encodeRet()],
      [encodePushString(1), encodePushString(0), encodeCall(BuiltinId.Bin2hex, 2), encodeRet()],
      [encodePushString(2), encodePushString(0), encodeCall(BuiltinId.Hex2bin, 2), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 4 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)
    const constants = packConstants([[10], [10], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    for (const output of outputs) {
      expect(kernel.readTags()[output]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[output]).toBe(ErrorCode.Value)
    }
  })

  it('returns #VALUE! when DECIMAL text exceeds the documented 255-character limit on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 2, 4, 1, 1)
    const strings = packStrings(['1'.repeat(255), '1'.repeat(256)])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const programs = packPrograms([
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Decimal, 2), encodeRet()],
      [encodePushString(1), encodePushNumber(0), encodeCall(BuiltinId.Decimal, 2), encodeRet()],
    ])
    const outputs = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)])
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)
    const constants = packConstants([[2], [2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    expect(kernel.readTags()[outputs[0]]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[outputs[0]]).toBeCloseTo(5.78960446186581e76)
    expect(kernel.readTags()[outputs[1]]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[outputs[1]]).toBe(ErrorCode.Value)
  })

  it('preserves incoming radix conversion errors before coercion and domain checks', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 2, 8, 2, 1)
    const strings = packStrings(['10'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const programs = packPrograms([
      [encodePushError(ErrorCode.Ref), encodePushNumber(0), encodeCall(BuiltinId.Base, 2), encodeRet()],
      [encodePushString(0), encodePushError(ErrorCode.NA), encodeCall(BuiltinId.Decimal, 2), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Bin2dec, 1), encodeRet()],
      [encodePushString(0), encodePushError(ErrorCode.Ref), encodeCall(BuiltinId.Oct2hex, 2), encodeRet()],
      [encodePushError(ErrorCode.Div0), encodeCall(BuiltinId.Dec2bin, 1), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 5 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, outputs)
    const constants = packConstants([[2], [], [], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    const expectedErrors = [ErrorCode.Ref, ErrorCode.NA, ErrorCode.Name, ErrorCode.Ref, ErrorCode.Div0]
    for (let index = 0; index < expectedErrors.length; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(expectedErrors[index])
    }
  })
})
