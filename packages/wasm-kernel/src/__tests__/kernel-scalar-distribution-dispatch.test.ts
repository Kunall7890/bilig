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

describe('wasm kernel scalar distribution dispatch', () => {
  it('matches Excel normal CDF precision for option-pricing inputs', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 7, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Normsdist, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Normsdist, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.NormDist, 4), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    const constants = packConstants([[-0.8281017980432489], [-0.9281017980432489], [-0.8281017980432489, 0, 1, 1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]))

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.203806425664055, 12)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.17667738351319964, 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.203806425664055, 12)
  })

  it('keeps scalar distribution and special-function dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 24
    kernel.init(120, 4, 21, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 1, 2, 2]), Uint32Array.from([1, 1, 0, 1]), Uint16Array.from([50, 49, 48]))
    kernel.writeCells(new Uint8Array(120), new Float64Array(120), new Uint32Array(120), new Uint16Array(120))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Gauss, 1), encodeRet()],
      [encodePushString(0), encodeCall(BuiltinId.Phi, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Erf, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Erf, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Erfc, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Fisherinv, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Gammaln, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.ConfidenceNorm, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.ConfidenceT, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Standardize, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.NormDist, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Norminv, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.NormSDist, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Normsinv, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Loginv, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.LognormDist, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.GammaInv, 3), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Erf, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Erf, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Erfc, 1), encodeRet()],
      [encodePushString(3), encodePushString(1), encodeCall(BuiltinId.Erf, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 21 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [],
      [],
      [1],
      [0, 1],
      [0],
      [0.5],
      [5],
      [0.05, 2, 16],
      [0.05, 2, 16],
      [8, 6, 2],
      [42, 40, 1.5, 1],
      [0.9, 40, 1.5],
      [1.25, 0],
      [0.75],
      [0.8, 1, 0.5],
      [3, 1, 0.5, 1],
      [0.08030139707139418, 3, 2],
      [],
      [],
      [],
      [],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 21 }, (_, index) => cellIndex(1, index, width))))

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.477249868052, 12)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.053990966513, 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.842700689748, 10)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.842700689748, 10)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.46211715726, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 3.178053830348, 10)
    expectNumberCell(kernel, cellIndex(1, 7, width), 0.97998199306, 10)
    expectNumberCell(kernel, cellIndex(1, 8, width), 1.06572477278, 10)
    expectNumberCell(kernel, cellIndex(1, 9, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 10, width), 0.908788780274, 12)
    expectNumberCell(kernel, cellIndex(1, 11, width), 41.92232734621, 10)
    expectNumberCell(kernel, cellIndex(1, 12, width), 0.182649085389, 10)
    expectNumberCell(kernel, cellIndex(1, 13, width), 0.674489750223, 10)
    expectNumberCell(kernel, cellIndex(1, 14, width), 4.140475417393, 10)
    expectNumberCell(kernel, cellIndex(1, 15, width), 0.578174100803, 12)
    expectNumberCell(kernel, cellIndex(1, 16, width), 2, 10)
    expectNumberCell(kernel, cellIndex(1, 17, width), 0.842700689748, 10)
    expectNumberCell(kernel, cellIndex(1, 18, width), 0, 12)
    expectNumberCell(kernel, cellIndex(1, 19, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 20, width), 0.842700689748, 10)
  })

  it('returns Excel-compatible domain errors for scalar distribution dispatch', async () => {
    const kernel = await createKernel()
    const width = 20
    const cases = [
      {
        program: [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Expondist, 3), encodeRet()],
        constants: [-1, 2, 0],
        error: ErrorCode.Num,
      },
      {
        program: [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.ExponDist, 3), encodeRet()],
        constants: [1, 0, 1],
        error: ErrorCode.Num,
      },
      {
        program: [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Expondist, 3), encodeRet()],
        constants: [2, 0],
        error: ErrorCode.Value,
      },
      {
        program: [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Poisson, 3), encodeRet()],
        constants: [-1, 2.5, 0],
        error: ErrorCode.Num,
      },
      {
        program: [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.PoissonDist, 3), encodeRet()],
        constants: [3, -1, 1],
        error: ErrorCode.Num,
      },
      {
        program: [encodePushNumber(0), encodePushString(0), encodePushNumber(1), encodeCall(BuiltinId.Poisson, 3), encodeRet()],
        constants: [3, 0],
        error: ErrorCode.Value,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushNumber(1),
          encodePushNumber(2),
          encodePushNumber(3),
          encodeCall(BuiltinId.Weibull, 4),
          encodeRet(),
        ],
        constants: [-1, 2, 3, 0],
        error: ErrorCode.Num,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushNumber(1),
          encodePushNumber(2),
          encodePushNumber(3),
          encodeCall(BuiltinId.WeibullDist, 4),
          encodeRet(),
        ],
        constants: [1, 0, 3, 1],
        error: ErrorCode.Num,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushNumber(1),
          encodePushNumber(2),
          encodePushNumber(3),
          encodeCall(BuiltinId.Weibull, 4),
          encodeRet(),
        ],
        constants: [1, 2, 0, 0],
        error: ErrorCode.Num,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushString(0),
          encodePushNumber(1),
          encodePushNumber(2),
          encodeCall(BuiltinId.Weibull, 4),
          encodeRet(),
        ],
        constants: [1, 2, 0],
        error: ErrorCode.Value,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushNumber(1),
          encodePushNumber(2),
          encodePushNumber(3),
          encodeCall(BuiltinId.Gammadist, 4),
          encodeRet(),
        ],
        constants: [-1, 3, 2, 0],
        error: ErrorCode.Num,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushNumber(1),
          encodePushNumber(2),
          encodePushNumber(3),
          encodeCall(BuiltinId.GammaDist, 4),
          encodeRet(),
        ],
        constants: [2, 0, 2, 1],
        error: ErrorCode.Num,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushNumber(1),
          encodePushNumber(2),
          encodePushNumber(3),
          encodeCall(BuiltinId.Gammadist, 4),
          encodeRet(),
        ],
        constants: [2, 3, 0, 0],
        error: ErrorCode.Num,
      },
      {
        program: [
          encodePushNumber(0),
          encodePushNumber(1),
          encodePushString(0),
          encodePushNumber(2),
          encodeCall(BuiltinId.GammaDist, 4),
          encodeRet(),
        ],
        constants: [2, 3, 1],
        error: ErrorCode.Value,
      },
    ]

    kernel.init(64, 1, cases.length, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([3]), Uint16Array.from([98, 97, 100]))
    kernel.writeCells(new Uint8Array(64), new Float64Array(64), new Uint32Array(64), new Uint16Array(64))

    const packed = packPrograms(cases.map((item) => item.program))
    const targets = Uint32Array.from(cases.map((_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targets)
    const constants = packConstants(cases.map((item) => item.constants))
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(targets)

    for (const [index, item] of cases.entries()) {
      expectErrorCell(kernel, cellIndex(1, index, width), item.error)
    }
  })
})
