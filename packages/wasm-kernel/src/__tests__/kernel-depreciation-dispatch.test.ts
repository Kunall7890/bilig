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

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

function expectErrorCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

describe('wasm kernel depreciation dispatch', () => {
  it('keeps depreciation formulas stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 6, 3, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Db, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Ddb, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Vdb, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Sln, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Syd, 4), encodeRet()],
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
      ]),
    )
    const constants = packConstants([
      [10000, 1000, 5, 1],
      [2400, 300, 10, 2],
      [2400, 300, 10, 1, 3],
      [10000, 1000, 9],
      [10000, 1000, 9, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(1, 0, width), 3690)
    expectNumberCell(kernel, cellIndex(1, 1, width), 384)
    expectNumberCell(kernel, cellIndex(1, 2, width), 691.2)
    expectNumberCell(kernel, cellIndex(1, 3, width), 1000)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1800)
  })

  it('uses #NUM! for DDB and VDB numeric-domain errors while preserving #VALUE! coercions', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 4, 2, 1)
    const strings = packStrings(['bad'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const programs = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Ddb, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Ddb, 5),
        encodeRet(),
      ],
      [
        encodePushString(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(5),
        encodeCall(BuiltinId.Ddb, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(6),
        encodePushNumber(7),
        encodePushNumber(5),
        encodeCall(BuiltinId.Vdb, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Vdb, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Vdb, 6),
        encodeRet(),
      ],
      [
        encodePushString(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Vdb, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushString(0),
        encodeCall(BuiltinId.Vdb, 7),
        encodeRet(),
      ],
    ])
    const targetCells = Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, targetCells)
    const constants = packConstants([
      [2400, 300, 10, 2, 0],
      [2400, 300, 0, 2, 2],
      [300, 10, 2, 2],
      [2400, 300, 10, 3, 1, 2],
      [2400, 300, 0, 1, 3, 2],
      [2400, 300, 10, 1, 3, 0],
      [300, 10, 1, 3, 2],
      [2400, 300, 10, 1, 3, 2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(targetCells)

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 5, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 6, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 7, width), ErrorCode.Value)
  })

  it('uses documented DB, DDB, SLN, and SYD domain errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 10, 3, 2, 1)
    const strings = packStrings(['bad'])
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const programs = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Db, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Db, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BuiltinId.Db, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Ddb, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Sln, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Syd, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Syd, 4), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Syd, 4), encodeRet()],
    ])
    const targetCells = Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(programs.programs, programs.offsets, programs.lengths, targetCells)
    const constants = packConstants([
      [10000, 1000, 5, 6],
      [10000, 1000, 5, 6, 12],
      [10000, 1000, 5, 1, 13],
      [2400, 300, 10, 11],
      [10000, 1000, 0],
      [10000, 1000, 9, 10],
      [10000, 1000, 9, 0],
      [1000, 9, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(targetCells)

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Div0)
    expectErrorCell(kernel, cellIndex(1, 5, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 6, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 7, width), ErrorCode.Value)
  })
})
