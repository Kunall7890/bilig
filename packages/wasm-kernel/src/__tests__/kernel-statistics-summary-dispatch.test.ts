import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
import { createKernel, type KernelInstance } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex
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

function decodeErrorCode(rawCode: number): ErrorCode {
  switch (rawCode) {
    case 1:
      return ErrorCode.Null
    case 2:
      return ErrorCode.Div0
    case 3:
      return ErrorCode.Value
    case 4:
      return ErrorCode.Ref
    case 5:
      return ErrorCode.Name
    case 6:
      return ErrorCode.Num
    case 7:
      return ErrorCode.NA
    case 8:
      return ErrorCode.Blocked
    default:
      throw new Error(`Unexpected error code: ${rawCode}`)
  }
}

function readSpillValues(kernel: KernelInstance, ownerCellIndex: number): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0
  const tags = kernel.readSpillTags()
  const values = kernel.readSpillNumbers()
  return Array.from({ length }, (_, index) => {
    const tag = tags[offset + index] ?? ValueTag.Empty
    const rawValue = values[offset + index] ?? 0
    if (tag === ValueTag.Number) {
      return { tag, value: rawValue }
    }
    if (tag === ValueTag.Error) {
      return { tag, code: decodeErrorCode(rawValue) }
    }
    if (tag === ValueTag.Empty) {
      return { tag }
    }
    throw new Error(`Unexpected spill tag: ${tag}`)
  })
}

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

function expectErrorCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

describe('wasm kernel ordered statistics dispatch slab', () => {
  it('keeps rank, percentiles, trimmean, and probability stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(120, 1, 6, 11, 32)

    const cellTags = new Uint8Array(120)
    const cellNumbers = new Float64Array(120)
    ;[10, 20, 20, 30].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    ;[2, 4, 4, 4, 5, 5, 7, 9].forEach((value, index) => {
      cellTags[10 + index] = ValueTag.Number
      cellNumbers[10 + index] = value
    })
    ;[79, 85, 78, 85, 50, 81].forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number
      cellNumbers[20 + index] = value
    })
    ;[60, 80, 90].forEach((value, index) => {
      cellTags[30 + index] = ValueTag.Number
      cellNumbers[30 + index] = value
    })
    ;[1, 2, 3].forEach((value, index) => {
      cellTags[40 + index] = ValueTag.Number
      cellNumbers[40 + index] = value
    })
    ;[0.2, 0.3, 0.5].forEach((value, index) => {
      cellTags[50 + index] = ValueTag.Number
      cellNumbers[50 + index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(120), new Uint16Array(120))

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 10, 11, 12, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24, 25, 30, 31, 32, 40, 41, 42, 50, 51, 52]),
      Uint32Array.from([0, 4, 12, 18, 21, 24]),
      Uint32Array.from([4, 8, 6, 3, 3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 8, 6, 3, 3, 3]), Uint32Array.from([1, 1, 1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushNumber(1), encodeCall(BuiltinId.RankAvg, 3), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.StdevS, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.VarP, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.Median, 1), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.Large, 2), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.PercentileInc, 2), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.QuartileExc, 2), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.PercentrankInc, 3), encodeRet()],
      [encodePushRange(4), encodePushRange(5), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Prob, 4), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.Trimmean, 2), encodeRet()],
      [encodePushRange(2), encodePushRange(3), encodeCall(BuiltinId.Frequency, 2), encodeRet()],
    ])
    const constants = packConstants([[20, 0], [], [], [], [2], [0.75], [3], [7, 4], [2, 3], [0.25], []])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
        cellIndex(5, 5, width),
        cellIndex(5, 6, width),
        cellIndex(5, 7, width),
        cellIndex(5, 8, width),
        cellIndex(5, 9, width),
        cellIndex(7, 0, width),
      ]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
        cellIndex(5, 5, width),
        cellIndex(5, 6, width),
        cellIndex(5, 7, width),
        cellIndex(5, 8, width),
        cellIndex(5, 9, width),
        cellIndex(7, 0, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(5, 0, width), 2.5, 12)
    expectNumberCell(kernel, cellIndex(5, 1, width), 2.138089935299395, 12)
    expectNumberCell(kernel, cellIndex(5, 2, width), 4, 12)
    expectNumberCell(kernel, cellIndex(5, 3, width), 4.5, 12)
    expectNumberCell(kernel, cellIndex(5, 4, width), 7, 12)
    expectNumberCell(kernel, cellIndex(5, 5, width), 5.5, 12)
    expectNumberCell(kernel, cellIndex(5, 6, width), 6.5, 12)
    expectNumberCell(kernel, cellIndex(5, 7, width), 0.8571, 12)
    expectNumberCell(kernel, cellIndex(5, 8, width), 0.8, 12)
    expectNumberCell(kernel, cellIndex(5, 9, width), 29 / 6, 12)
    expect(readSpillValues(kernel, cellIndex(7, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 0 },
    ])
  })

  it('ignores referenced non-numeric values for MEDIAN on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 0, 1, 2)

    const cellTags = new Uint8Array(8)
    const cellNumbers = new Float64Array(8)
    cellTags[0] = ValueTag.String
    cellTags[1] = ValueTag.Number
    cellNumbers[1] = 1
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(8), new Uint16Array(8))
    kernel.uploadRangeMembers(Uint32Array.from([0, 1]), Uint32Array.from([0]), Uint32Array.from([2]))
    kernel.uploadRangeShapes(Uint32Array.from([1]), Uint32Array.from([2]))

    const packed = packPrograms([[encodePushRange(0), encodeCall(BuiltinId.Median, 1), encodeRet()]])
    const target = cellIndex(1, 0, width)
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([target]))
    kernel.uploadConstants(Float64Array.from([]), Uint32Array.from([0]), Uint32Array.from([0]))
    kernel.evalBatch(Uint32Array.from([target]))

    expectNumberCell(kernel, target, 1, 12)
  })

  it('returns Excel-compatible division errors for insufficient variance samples', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(32, 5, 0, 2, 4)

    const cellTags = new Uint8Array(32)
    const cellNumbers = new Float64Array(32)
    cellTags[0] = ValueTag.Number
    cellNumbers[0] = 42
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(32), new Uint16Array(32))
    kernel.uploadRangeMembers(Uint32Array.from([0, 1]), Uint32Array.from([0, 1]), Uint32Array.from([1, 1]))
    kernel.uploadRangeShapes(Uint32Array.from([1, 1]), Uint32Array.from([1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BuiltinId.StdevS, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.VarS, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.StdevP, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BuiltinId.VarP, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.StdevP, 1), encodeRet()],
    ])
    const constants = packConstants([[], [], [], [], []])
    const targets = Uint32Array.from([
      cellIndex(2, 0, width),
      cellIndex(2, 1, width),
      cellIndex(2, 2, width),
      cellIndex(2, 3, width),
      cellIndex(2, 4, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targets)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(targets)

    expectErrorCell(kernel, cellIndex(2, 0, width), ErrorCode.Div0)
    expectErrorCell(kernel, cellIndex(2, 1, width), ErrorCode.Div0)
    expectErrorCell(kernel, cellIndex(2, 2, width), ErrorCode.Div0)
    expectErrorCell(kernel, cellIndex(2, 3, width), ErrorCode.Div0)
    expectNumberCell(kernel, cellIndex(2, 4, width), 0)
  })

  it('returns #DIV/0! for skew and kurt invalid-dispersion domains', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 6, 16, 0, 0)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Skew, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(0), encodePushNumber(0), encodeCall(BuiltinId.Skew, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.SkewP, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(0), encodePushNumber(0), encodeCall(BuiltinId.SkewP, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Kurt, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(0), encodePushNumber(0), encodePushNumber(0), encodeCall(BuiltinId.Kurt, 4), encodeRet()],
    ])
    const constants = packConstants([[1, 2], [1], [1, 2], [1], [1, 2, 3], [1]])
    const targets = Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targets)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(targets)

    for (const target of targets) {
      expectErrorCell(kernel, target, ErrorCode.Div0)
    }
  })

  it('returns Excel-compatible domain errors for ordered statistics helpers', async () => {
    const kernel = await createKernel()
    const width = 20
    const targetRow = 8
    kernel.init(220, 17, 17, 7, 22)

    const cellTags = new Uint8Array(220)
    const cellNumbers = new Float64Array(220)
    ;[1, 2, 4, 7, 8, 9, 10, 12].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    ;[1, 2, 3].forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number
      cellNumbers[20 + index] = value
    })
    ;[0.2, 0.2, -0.1].forEach((value, index) => {
      cellTags[30 + index] = ValueTag.Number
      cellNumbers[30 + index] = value
    })
    ;[0, 0.5, 0.5].forEach((value, index) => {
      cellTags[40 + index] = ValueTag.Number
      cellNumbers[40 + index] = value
    })
    ;[0.2, 0.2, 0.2].forEach((value, index) => {
      cellTags[50 + index] = ValueTag.Number
      cellNumbers[50 + index] = value
    })
    ;[0.5, 0.5].forEach((value, index) => {
      cellTags[60 + index] = ValueTag.Number
      cellNumbers[60 + index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(220), new Uint16Array(220))

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 20, 21, 22, 30, 31, 32, 40, 41, 42, 50, 51, 52, 60, 61]),
      Uint32Array.from([0, 8, 11, 14, 17, 20, 22]),
      Uint32Array.from([8, 3, 3, 3, 3, 2, 0]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([8, 3, 3, 3, 3, 2, 0]), Uint32Array.from([1, 1, 1, 1, 1, 1, 0]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Small, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Large, 2), encodeRet()],
      [encodePushRange(6), encodePushNumber(0), encodeCall(BuiltinId.Small, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Percentile, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.PercentileExc, 2), encodeRet()],
      [encodePushRange(6), encodePushNumber(0), encodeCall(BuiltinId.Percentile, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Quartile, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.QuartileExc, 2), encodeRet()],
      [encodePushRange(6), encodePushNumber(0), encodeCall(BuiltinId.Quartile, 2), encodeRet()],
      [encodePushRange(6), encodePushNumber(0), encodeCall(BuiltinId.Percentrank, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Percentrank, 3), encodeRet()],
      [encodePushRange(1), encodePushRange(2), encodePushNumber(0), encodeCall(BuiltinId.Prob, 3), encodeRet()],
      [encodePushRange(1), encodePushRange(3), encodePushNumber(0), encodeCall(BuiltinId.Prob, 3), encodeRet()],
      [encodePushRange(1), encodePushRange(4), encodePushNumber(0), encodeCall(BuiltinId.Prob, 3), encodeRet()],
      [encodePushRange(1), encodePushRange(5), encodePushNumber(0), encodeCall(BuiltinId.Prob, 3), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Trimmean, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Trimmean, 2), encodeRet()],
    ])
    const constants = packConstants([
      [0],
      [99],
      [1],
      [-0.01],
      [0.01],
      [0.5],
      [-1],
      [0],
      [1],
      [1],
      [8, 0],
      [1],
      [1],
      [1],
      [1],
      [-0.1],
      [1.1],
    ])
    const targets = Uint32Array.from(Array.from({ length: 17 }, (_, index) => cellIndex(targetRow, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targets)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(targets)

    const expectedErrors = [
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.Num,
      ErrorCode.NA,
      ErrorCode.Num,
      ErrorCode.Num,
    ]
    expectedErrors.forEach((expectedError, index) => {
      const target = cellIndex(targetRow, index, width)
      expect(kernel.readTags()[target]).toBe(ValueTag.Error)
      expect({ index, code: kernel.readErrors()[target] }).toEqual({ index, code: expectedError })
    })
  })
})
