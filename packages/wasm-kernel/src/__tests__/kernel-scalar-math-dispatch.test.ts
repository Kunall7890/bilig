import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushBoolean(value: boolean): number {
  return (Opcode.PushBoolean << 24) | (value ? 1 : 0)
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

function encodeBinary(opcode: Opcode): number {
  return opcode << 24
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

function packStrings(values: string[]): {
  offsets: Uint32Array
  lengths: Uint32Array
  data: Uint16Array
} {
  const offsets: number[] = []
  const lengths: number[] = []
  const data: number[] = []
  let offset = 0
  for (const value of values) {
    offsets.push(offset)
    lengths.push(value.length)
    for (const char of value) {
      data.push(char.charCodeAt(0))
    }
    offset += value.length
  }
  return {
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
    data: Uint16Array.from(data),
  }
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

function expectKernelError(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, code: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(code)
}

describe('wasm kernel scalar math dispatch', () => {
  it('returns value errors for malformed scalar math arities', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Pi, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Floor, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Ceiling, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Abs, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Round, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Power, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )
    const constants = packConstants([[2.5], [2.5], [2.5], [1, 2], [1, 2, 3], [2, 3, 4]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )

    expectKernelError(kernel, cellIndex(1, 0, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 1, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 2, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 3, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 4, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 5, width), ErrorCode.Value)
  })

  it('rejects nonnumeric SERIESSUM coefficients', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 2, 4, 1, 1)
    const strings = packStrings(['bad', '2'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Seriessum, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushString(0), encodeCall(BuiltinId.Seriessum, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushString(1), encodeCall(BuiltinId.Seriessum, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.Seriessum, 4),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    const constants = packConstants([
      [1, 1, 1, 2],
      [1, 1, 1],
      [1, 1, 1],
      [1, 1, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(2)
    expectKernelError(kernel, cellIndex(1, 1, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 2, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 3, width), ErrorCode.Value)
  })

  it('validates the left arithmetic operand before propagating right operand errors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    const strings = packStrings(['523a', '42'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushString(0), encodePushError(ErrorCode.Ref), encodeBinary(Opcode.Add), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodePushString(0), encodeBinary(Opcode.Add), encodeRet()],
      [encodePushString(1), encodePushError(ErrorCode.Ref), encodeBinary(Opcode.Add), encodeRet()],
      [encodePushNumber(0), encodePushError(ErrorCode.Ref), encodeBinary(Opcode.Add), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    const constants = packConstants([[], [], [], [42]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expectKernelError(kernel, cellIndex(1, 0, width), ErrorCode.Value)
    expectKernelError(kernel, cellIndex(1, 1, width), ErrorCode.Ref)
    expectKernelError(kernel, cellIndex(1, 2, width), ErrorCode.Ref)
    expectKernelError(kernel, cellIndex(1, 3, width), ErrorCode.Ref)
  })

  it('keeps exponentiation operator odd roots separate from POWER domain errors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeBinary(Opcode.Pow), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Power, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    )
    const constants = packConstants([
      [-32, 0.2],
      [-32, 0.2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(-2, 12)
    expectKernelError(kernel, cellIndex(1, 1, width), ErrorCode.Num)
  })

  it('keeps rounding and core scalar math dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 3, 8, 1, 1)
    const strings = packStrings(['bad', '0008', '41CA00300'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Abs, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Round, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.FloorMath, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.CeilingPrecise, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Int, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.RoundUp, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.RoundDown, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Trunc, 2), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Round, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Int, 1), encodeRet()],
      [encodePushString(2), encodePushNumber(0), encodeCall(BuiltinId.Right, 2), encodeCall(BuiltinId.Int, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 15 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [-3.5],
      [1.234, 2],
      [5.5, 2],
      [-5.5, 2],
      [-5.5, 2, 0],
      [-5.5, 2],
      [5, 2],
      [-1.2],
      [-1.23, 1],
      [-1.29, 1],
      [-1.29, 1],
      [],
      [10, 3],
      [],
      [4],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 15 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(3.5)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(1.23, 12)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(4)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(-6)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(-2)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBeCloseTo(-1.3, 12)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBeCloseTo(-1.2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBeCloseTo(-1.2, 12)
    expect(kernel.readTags()[cellIndex(1, 11, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 11, width)]).toBe(ErrorCode.Value)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(9)
    expect(kernel.readNumbers()[cellIndex(1, 13, width)]).toBe(8)
    expect(kernel.readNumbers()[cellIndex(1, 14, width)]).toBe(300)
  })

  it('matches Desktop Excel MOD sign semantics for negative operands on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    const constants = packConstants([
      [-3, 2],
      [3, -2],
      [-3, -2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(-1)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-1)
  })

  it('matches Microsoft Excel ATAN2 coordinate argument order and zero-origin error on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Atan2, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Atan2, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Atan2, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    const constants = packConstants([
      [-1, 1],
      [1, -1],
      [0, 0],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo((3 * Math.PI) / 4, 12)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(-Math.PI / 4, 12)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 2, width)]).toBe(ErrorCode.Div0)
  })

  it('matches Microsoft Excel FLOOR positive-number negative-significance error semantics on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    const constants = packConstants([
      [2.5, -2],
      [-2.5, 2],
      [-2.5, -2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 0, width)]).toBe(ErrorCode.Num)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-2)
  })

  it('matches spreadsheet FLOOR and CEILING zero-number zero-significance semantics on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    const strings = packStrings([''])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[0, 0], [0], [0], [0, 0], [0], [0], [2.5, 0], [2.5, 0]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 6; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Number)
      expect(kernel.readNumbers()[cellIndex(1, index, width)]).toBe(0)
    }
    expectKernelError(kernel, cellIndex(1, 6, width), ErrorCode.Div0)
    expectKernelError(kernel, cellIndex(1, 7, width), ErrorCode.Div0)
  })

  it('matches Microsoft Excel ROUND and CEILING negative-number edge semantics on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 7, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Round, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Round, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Round, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodeBinary(Opcode.Div),
        encodePushNumber(1),
        encodeBinary(Opcode.Div),
        encodePushNumber(2),
        encodeBinary(Opcode.Div),
        encodePushNumber(3),
        encodeCall(BuiltinId.Round, 2),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 7 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [-2.5, 0],
      [-1.475, 2],
      [-50.55, -2],
      [16150000, 1000, 0.1, 0],
      [2.5, -2],
      [-2.5, 2],
      [-2.5, -2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 7 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(-3)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(-1.48, 12)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-100)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(162)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 4, width)]).toBe(ErrorCode.Num)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(-2)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(-4)
  })

  it('matches Microsoft Excel MROUND sign-domain errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    const constants = packConstants([
      [-10, 3],
      [5, -2],
      [-10, -3],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 0, width)]).toBe(ErrorCode.Num)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 1, width)]).toBe(ErrorCode.Num)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-9)
  })

  it('snaps decimal quotient math to spreadsheet multiple semantics on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 5, 1, 1, 1)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.FloorMath, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.FloorPrecise, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.CeilingMath, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.CeilingPrecise, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.IsoCeiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Quotient, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Quotient, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_value, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [0.3, 0.1],
      [0.3, 0.1],
      [0.3, 0.1],
      [0.3, 0.1],
      [0.3, 0.1],
      [0.3, 0.1],
      [0.3, 0.1],
      [0.15, 0.1],
      [6.05, 0.1],
      [0.35, 0.1],
      [0.3, 0.1],
      [0.3, 0.1],
      [-0.3, 0.1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 13 }, (_value, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 7; index += 1) {
      expect(kernel.readNumbers()[cellIndex(1, index, width)]).toBeCloseTo(0.3, 12)
    }
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBeCloseTo(0.2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBeCloseTo(6.1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBeCloseTo(0.05, 12)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBeCloseTo(0, 12)
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(-3)
  })

  it('matches spreadsheet zero-multiple rounding semantics on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.FloorMath, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.FloorPrecise, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.CeilingMath, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.CeilingPrecise, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.IsoCeiling, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 8 }, (_value, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants(Array.from({ length: 8 }, () => [2.5, 0]))
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 8 }, (_value, index) => cellIndex(1, index, width))))

    expectKernelError(kernel, cellIndex(1, 0, width), ErrorCode.Div0)
    expectKernelError(kernel, cellIndex(1, 1, width), ErrorCode.Div0)
    for (let index = 2; index < 8; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Number)
      expect(kernel.readNumbers()[cellIndex(1, index, width)]).toBe(0)
    }
  })

  it('keeps trigonometric and transcendental dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Sin, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Cos, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Tan, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Asin, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Acos, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Atan, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Atan2, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Degrees, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Radians, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Exp, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Ln, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Log10, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Log, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Power, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sqrt, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Seriessum, 6),
        encodeRet(),
      ],
      [encodePushNumber(0), encodeCall(BuiltinId.Sqrtpi, 1), encodeRet()],
      [encodeCall(BuiltinId.Pi, 0), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Cot, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Csc, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sech, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sign, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Power, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 23 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [Math.PI / 2],
      [0],
      [Math.PI / 4],
      [1],
      [1],
      [1],
      [1, 1],
      [Math.PI],
      [180],
      [1],
      [Math.E],
      [100],
      [100, 10],
      [2, 3],
      [9],
      [1, 0, 1, 1, 2, 3],
      [4],
      [],
      [Math.PI / 4],
      [0],
      [0],
      [-7],
      [-32, 1 / 5],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 23 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBeCloseTo(Math.PI / 2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBeCloseTo(0, 12)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBeCloseTo(Math.PI / 4, 12)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBeCloseTo(Math.PI / 4, 12)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBeCloseTo(180, 12)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBeCloseTo(Math.PI, 12)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBeCloseTo(Math.E, 12)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBeCloseTo(2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBeCloseTo(2, 12)
    expect(kernel.readNumbers()[cellIndex(1, 13, width)]).toBeCloseTo(8, 12)
    expect(kernel.readNumbers()[cellIndex(1, 14, width)]).toBeCloseTo(3, 12)
    expect(kernel.readNumbers()[cellIndex(1, 15, width)]).toBeCloseTo(6, 12)
    expect(kernel.readNumbers()[cellIndex(1, 16, width)]).toBeCloseTo(2 * Math.sqrt(Math.PI), 12)
    expect(kernel.readNumbers()[cellIndex(1, 17, width)]).toBeCloseTo(Math.PI, 12)
    expect(kernel.readNumbers()[cellIndex(1, 18, width)]).toBeCloseTo(1, 12)
    expect(kernel.readTags()[cellIndex(1, 19, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 19, width)]).toBe(ErrorCode.Div0)
    expect(kernel.readNumbers()[cellIndex(1, 20, width)]).toBeCloseTo(1, 12)
    expect(kernel.readNumbers()[cellIndex(1, 21, width)]).toBe(-1)
    expectKernelError(kernel, cellIndex(1, 22, width), ErrorCode.Num)
  })

  it('matches Excel scalar math text coercion and overflow errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 5, 8, 1, 1)
    const strings = packStrings(['bad', '1', '2', '3', ''])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Sin, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Cos, 1), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Power, 2), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Sin, 1), encodeRet()],
      [encodePushString(2), encodePushString(3), encodeCall(BuiltinId.Power, 2), encodeRet()],
      [encodePushString(4), encodeCall(BuiltinId.Exp, 1), encodeRet()],
      [encodePushString(4), encodeCall(BuiltinId.Int, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Exp, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Power, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sinh, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Cosh, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 11 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[2], [], [], [], [], [], [], [1000], [10, 400], [1000], [1000]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 11 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 3; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Value)
    }
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBeCloseTo(Math.sin(1), 12)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(8)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(0)
    for (let index = 7; index < 11; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Num)
    }
  })

  it('coerces direct numeric text across scalar math dispatch', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 11, 8, 1, 1)
    const strings = packStrings(['-2', '2.5', '5.5', '2', '7', '10', '3', '100', '12.34', '4', '5'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Abs, 1), encodeRet()],
      [encodePushString(1), encodePushNumber(0), encodeCall(BuiltinId.Round, 2), encodeRet()],
      [encodePushString(2), encodePushString(3), encodeCall(BuiltinId.Floor, 2), encodeRet()],
      [encodePushString(2), encodePushString(3), encodeCall(BuiltinId.Ceiling, 2), encodeRet()],
      [encodePushString(4), encodePushString(3), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushString(4), encodePushString(3), encodeCall(BuiltinId.Quotient, 2), encodeRet()],
      [encodePushString(5), encodePushString(6), encodeCall(BuiltinId.Mround, 2), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.Ln, 1), encodeRet()],
      [encodePushString(7), encodePushString(5), encodeCall(BuiltinId.Log, 2), encodeRet()],
      [encodePushString(9), encodeCall(BuiltinId.Sqrt, 1), encodeRet()],
      [encodePushString(8), encodePushNumber(0), encodeCall(BuiltinId.RoundUp, 2), encodeRet()],
      [encodePushString(8), encodePushNumber(0), encodeCall(BuiltinId.RoundDown, 2), encodeRet()],
      [encodePushString(8), encodePushNumber(0), encodeCall(BuiltinId.Trunc, 2), encodeRet()],
      [encodePushString(6), encodeCall(BuiltinId.Even, 1), encodeRet()],
      [encodePushString(6), encodeCall(BuiltinId.Odd, 1), encodeRet()],
      [encodePushString(10), encodeCall(BuiltinId.Fact, 1), encodeRet()],
      [encodePushString(10), encodePushString(3), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushString(9), encodePushString(6), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushString(10), encodePushString(3), encodeCall(BuiltinId.Permut, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 19 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[], [0], [], [], [], [], [], [], [], [], [1], [1], [1], [], [], [], [], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 19 }, (_, index) => cellIndex(1, index, width))))

    const expected = [2, 3, 4, 6, 1, 3, 9, Math.log(2), 2, 2, 12.4, 12.3, 12.3, 4, 3, 120, 10, 20, 20]
    for (const [index, value] of expected.entries()) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Number)
      expect(kernel.readNumbers()[cellIndex(1, index, width)]).toBeCloseTo(value, 12)
    }
  })

  it('propagates scalar math argument errors before coercion and domain checks on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Abs, 1), encodeRet()],
      [encodePushNumber(0), encodePushError(ErrorCode.NA), encodeCall(BuiltinId.Round, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushError(ErrorCode.NA), encodeCall(BuiltinId.FloorMath, 3), encodeRet()],
      [encodePushError(ErrorCode.Name), encodePushNumber(1), encodeCall(BuiltinId.Mod, 2), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Int, 1), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Acot, 1), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Cot, 1), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Sign, 1), encodeRet()],
      [encodePushError(ErrorCode.Name), encodePushNumber(0), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushNumber(0), encodePushError(ErrorCode.NA), encodeCall(BuiltinId.Permut, 2), encodeRet()],
      [encodePushError(ErrorCode.Name), encodePushNumber(1), encodeCall(BuiltinId.Mround, 2), encodeRet()],
      [encodePushError(ErrorCode.Name), encodePushNumber(0), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushNumber(0), encodePushError(ErrorCode.NA), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushError(ErrorCode.NA),
        encodeCall(BuiltinId.Seriessum, 4),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 14 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[], [1], [1, 1], [0], [], [], [], [], [1], [1], [1], [1], [1], [1, 1, 1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 14 }, (_, index) => cellIndex(1, index, width))))

    const expectedErrors = [
      ErrorCode.Name,
      ErrorCode.NA,
      ErrorCode.NA,
      ErrorCode.Name,
      ErrorCode.Name,
      ErrorCode.Name,
      ErrorCode.Name,
      ErrorCode.Name,
      ErrorCode.Name,
      ErrorCode.NA,
      ErrorCode.Name,
      ErrorCode.Name,
      ErrorCode.NA,
      ErrorCode.NA,
    ]
    for (const [index, code] of expectedErrors.entries()) {
      expectKernelError(kernel, cellIndex(1, index, width), code)
    }
  })

  it('returns Excel-compatible log errors through wasm dispatch', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 2, 8, 1, 1)
    const strings = packStrings(['bad'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Ln, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Log10, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Log, 2), encodeRet()],
      [encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Log, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Log, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 5 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[-1], [0], [10, 1], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 5 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 0, width)]).toBe(ErrorCode.Num)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 1, width)]).toBe(ErrorCode.Num)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 2, width)]).toBe(ErrorCode.Num)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Name)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 4, width)]).toBe(ErrorCode.Value)
  })

  it('returns Excel-compatible square-root domain errors through wasm dispatch', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 1, 8, 1, 1)
    const strings = packStrings(['bad'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Sqrt, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Sqrtpi, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Sqrt, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Sqrtpi, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 4 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[-1], [-1], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 4 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 2; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Num)
    }
    for (let index = 2; index < 4; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Value)
    }
  })

  it('returns Excel-compatible inverse trigonometric domain errors through wasm dispatch', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 1, 8, 1, 1)
    const strings = packStrings(['bad'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Asin, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Acos, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Acosh, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Atanh, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Atanh, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Acoth, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Acoth, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Asin, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Acosh, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Acoth, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[2], [2], [0.5], [1], [-1], [0.5], [-0.5], [], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 7; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Num)
    }
    for (let index = 7; index < 10; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Value)
    }
  })

  it('keeps bessel and combinatorics dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besseli, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselj, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Fact, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Factdouble, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Quotient, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permut, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permutationa, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Even, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Odd, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[1, 0], [1, 0], [1, 0], [1, 0], [5], [6], [5, 2], [2, 3], [7, 3], [4, 2], [3, 2], [-3.2], [-3.2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(1.2660658777520084, 6)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(0.7651976865579666, 6)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(0.42102443824070834, 6)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBeCloseTo(0.088256964215677, 6)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(120)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(48)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(10)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(4)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBe(12)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBe(9)
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(-5)
  })

  it('keeps combinatoric cancellation finite and overflow errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Fact, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Factdouble, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permut, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permutationa, 2), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputs)

    const constants = packConstants([[171], [301], [171, 1], [171, 2], [171, 1], [100, 100], [200, 170], [200, 200]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputs)

    expectKernelError(kernel, outputs[0], ErrorCode.Num)
    expectKernelError(kernel, outputs[1], ErrorCode.Num)
    expect(kernel.readTags()[outputs[2]]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[outputs[2]]).toBe(171)
    expect(kernel.readTags()[outputs[3]]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[outputs[3]]).toBe(14535)
    expect(kernel.readTags()[outputs[4]]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[outputs[4]]).toBe(171)
    expect(kernel.readTags()[outputs[5]]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[outputs[5]] / 4.5274257328e58).toBeCloseTo(1, 11)
    expectKernelError(kernel, outputs[6], ErrorCode.Num)
    expectKernelError(kernel, outputs[7], ErrorCode.Num)
  })

  it('returns Excel-compatible numeric-domain errors for combinatorics through wasm dispatch', async () => {
    const kernel = await createKernel()
    const width = 32
    kernel.init(96, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Fact, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Factdouble, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combin, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Combina, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permut, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permut, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permutationa, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Permutationa, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 11 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[-1], [-1], [-1, 0], [2, 3], [-1, 1], [1, -1], [0, 1], [0, 1], [3, 4], [0, 1], [-1, 1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 11 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 11; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Num)
    }
  })

  it('returns Excel-compatible Bessel domain errors through wasm dispatch', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 1, 8, 1, 1)
    const strings = packStrings(['bad'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besseli, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselj, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Besseli, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodeCall(BuiltinId.Besselj, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[1, -1], [1, -1], [1, -1], [1, -1], [1], [1], [0, 1], [0, 1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 4; index += 1) {
      expectKernelError(kernel, cellIndex(1, index, width), ErrorCode.Num)
    }
    for (let index = 4; index < 6; index += 1) {
      expectKernelError(kernel, cellIndex(1, index, width), ErrorCode.Value)
    }
    for (let index = 6; index < 8; index += 1) {
      expectKernelError(kernel, cellIndex(1, index, width), ErrorCode.Num)
    }
  })

  it('coerces direct numeric and empty text for Bessel dispatch', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(32, 3, 8, 1, 1)
    const strings = packStrings(['1', '0', ''])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(32), new Float64Array(32), new Uint32Array(32), new Uint16Array(32))

    const packed = packPrograms([
      [encodePushString(0), encodePushString(1), encodeCall(BuiltinId.Besseli, 2), encodeRet()],
      [encodePushString(0), encodePushString(1), encodeCall(BuiltinId.Besselj, 2), encodeRet()],
      [encodePushString(0), encodePushString(0), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushString(0), encodePushString(0), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
      [encodePushString(2), encodePushString(1), encodeCall(BuiltinId.Besseli, 2), encodeRet()],
      [encodePushString(2), encodePushString(1), encodeCall(BuiltinId.Besselj, 2), encodeRet()],
      [encodePushString(2), encodePushString(0), encodeCall(BuiltinId.Besselk, 2), encodeRet()],
      [encodePushString(2), encodePushString(0), encodeCall(BuiltinId.Bessely, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants(Array.from({ length: 8 }, () => []))
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(1.266065878, 8)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(0.765197687, 8)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(0.60190723, 7)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBeCloseTo(-0.78121282, 8)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1)
    expectKernelError(kernel, cellIndex(1, 6, width), ErrorCode.Num)
    expectKernelError(kernel, cellIndex(1, 7, width), ErrorCode.Num)
  })
})
