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

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

function expectErrorCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

describe('wasm kernel finance/cashflow dispatch', () => {
  it('keeps annuity formulas stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 2, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Pv, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Pmt, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Nper, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Rate, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Ipmt, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Ppmt, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Ispmt, 4), encodeRet()],
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
        cellIndex(1, 6, width),
      ]),
    )
    const constants = packConstants([
      [0.1, 2, -576.1904761904761],
      [0.1, 2, 1000],
      [0.1, -576.1904761904761, 1000],
      [48, -200, 8000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(1, 0, width), 1000.0000000000006)
    expectNumberCell(kernel, cellIndex(1, 1, width), -576.1904761904758)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1.9999999999999982)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.007701472488246008)
    expectNumberCell(kernel, cellIndex(1, 4, width), -100)
    expectNumberCell(kernel, cellIndex(1, 5, width), -476.1904761904758)
    expectNumberCell(kernel, cellIndex(1, 6, width), -50)
  })

  it('keeps cumulative and growth formulas stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 2, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumipmt, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumprinc, 6),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Fv, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Npv, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Fvschedule, 4),
        encodeRet(),
      ],
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
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.1, 2, -100, -1000],
      [0.1, 100, 200, 300],
      [1000, 0.09, 0.11, 0.1],
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

    expectNumberCell(kernel, cellIndex(1, 0, width), -11135.232130750845, 9)
    expectNumberCell(kernel, cellIndex(1, 1, width), -934.1071234208765, 9)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1420)
    expectNumberCell(kernel, cellIndex(1, 3, width), 481.5927873779113)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1330.89)
  })

  it('matches Microsoft Excel CUMIPMT and CUMPRINC numeric domain errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 1, 7, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([3]), Uint16Array.from(Array.from('bad', (char) => char.charCodeAt(0))))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumipmt, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumipmt, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumprinc, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumipmt, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumprinc, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumipmt, 6),
        encodeRet(),
      ],
      [
        encodePushString(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Cumipmt, 6),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 7 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [0, 30 * 12, 125000, 13, 24, 0],
      [0.09 / 12, 0, 125000, 13, 24, 0],
      [0.09 / 12, 30 * 12, 0, 13, 24, 0],
      [0.09 / 12, 30 * 12, 125000, 24, 13, 0],
      [0.09 / 12, 30 * 12, 125000, 0, 24, 0],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 2],
      [30 * 12, 125000, 13, 24, 0],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 7 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 6; index += 1) {
      expectErrorCell(kernel, cellIndex(1, index, width), ErrorCode.Num)
    }
    expectErrorCell(kernel, cellIndex(1, 6, width), ErrorCode.Value)
  })

  it('keeps rate-conversion helpers stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 2, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Effect, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Nominal, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Pduration, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Rri, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    const constants = packConstants([
      [0.12, 12],
      [0.12682503013196977, 12],
      [0.1, 100, 121],
      [2, 100, 121],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.12682503013196977)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 2)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.1)
  })

  it('returns Excel-compatible domain errors for rate-conversion helpers', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(48, 13, 32, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([3]), Uint16Array.from([98, 97, 100]))
    kernel.writeCells(new Uint8Array(48), new Float64Array(48), new Uint32Array(48), new Uint16Array(48))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Effect, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Effect, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Effect, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Nominal, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Nominal, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodeCall(BuiltinId.Nominal, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Pduration, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Pduration, 3), encodeRet()],
      [encodePushNumber(0), encodePushString(0), encodePushNumber(1), encodeCall(BuiltinId.Pduration, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Rri, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Rri, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Rri, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushString(0), encodeCall(BuiltinId.Rri, 3), encodeRet()],
    ])
    const targets = Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targets)
    const constants = packConstants([
      [0, 12],
      [0.1, 0],
      [12],
      [0, 12],
      [0.1, 0],
      [0.1],
      [0, 100, 121],
      [0.1, -100, 121],
      [0.1, 121],
      [0, 100, 121],
      [2, 0, 121],
      [2, 100, -121],
      [2, 100],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(targets)

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 5, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 6, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 7, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 8, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 9, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 10, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 11, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 12, width), ErrorCode.Value)
  })
})
