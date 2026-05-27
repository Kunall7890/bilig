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

describe('wasm kernel date/finance helpers', () => {
  it('keeps US DAYS360 February month-end handling on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 0, 1, 1)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodeCall(BuiltinId.Days360, 2),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([cellIndex(1, 0, width)]))
    const constants = packConstants([[2024, 2, 29, 3, 31]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width)]))

    expectNumberCell(kernel, cellIndex(1, 0, width), 31)
  })

  it('evaluates WEEKNUM return type 21 as ISO system 2 on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 0, 1, 1)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Weeknum, 2),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([cellIndex(1, 0, width)]))
    const constants = packConstants([[2021, 1, 1, 21]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width)]))

    expectNumberCell(kernel, cellIndex(1, 0, width), 53)
  })

  it('keeps ISOWEEKNUM on the calendar path instead of the 1900 WEEKDAY compatibility path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 0, 1, 1)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([[encodePushNumber(0), encodeCall(BuiltinId.Isoweeknum, 1), encodeRet()]])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([cellIndex(1, 0, width)]))
    const constants = packConstants([[1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width)]))

    expectNumberCell(kernel, cellIndex(1, 0, width), 1)
  })

  it('preserves Excel 1900-system WEEKDAY compatibility before March 1 1900 on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 7, 0, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Weekday, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Weekday, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Weekday, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Weekday, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Weekday, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Weekday, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Weekday, 2), encodeRet()],
    ])
    const targetCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targetCells)
    const constants = packConstants([[0], [1], [59], [60], [61], [1, 2], [1, 3]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(targetCells)

    expectNumberCell(kernel, cellIndex(1, 0, width), 7)
    expectNumberCell(kernel, cellIndex(1, 1, width), 1)
    expectNumberCell(kernel, cellIndex(1, 2, width), 3)
    expectNumberCell(kernel, cellIndex(1, 3, width), 4)
    expectNumberCell(kernel, cellIndex(1, 4, width), 5)
    expectNumberCell(kernel, cellIndex(1, 5, width), 7)
    expectNumberCell(kernel, cellIndex(1, 6, width), 6)
  })

  it('returns #NUM for invalid WEEKDAY and WEEKNUM date domains on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 6, 0, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Weekday, 2),
        encodeRet(),
      ],
      [encodePushNumber(0), encodeCall(BuiltinId.Weekday, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Weekday, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Weeknum, 2),
        encodeRet(),
      ],
      [encodePushNumber(0), encodeCall(BuiltinId.Weeknum, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Weeknum, 1), encodeRet()],
    ])
    const targetCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targetCells)
    const constants = packConstants([[2024, 1, 1, 0], [-1], [2_958_466], [2024, 1, 1, 3], [-1], [2_958_466]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(targetCells)

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 5, width), ErrorCode.Num)
  })

  it('returns documented date-domain errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 7, 0, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(1), encodeCall(BuiltinId.Date, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(1), encodeCall(BuiltinId.Date, 3), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Isoweeknum, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Isoweeknum, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Eomonth, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Eomonth, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Edate, 2), encodeRet()],
    ])
    const targetCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targetCells)
    const constants = packConstants([[-1, 1], [10000, 1], [-1], [2_958_466], [-1, 0], [2_958_465, 1], [-1, 1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(targetCells)

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 5, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 6, width), ErrorCode.Value)
  })

  it('keeps date-basis helper behavior stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 3, 2, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 1]),
      Uint32Array.from([1, 2]),
      Uint16Array.from(Array.from('DYM', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(1),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(2),
        encodePushNumber(1),
        encodeCall(BuiltinId.Date, 3),
        encodePushString(0),
        encodeCall(BuiltinId.Datedif, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(1),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(2),
        encodePushNumber(1),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Days360, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(1),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(2),
        encodePushNumber(1),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Yearfrac, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodeCall(BuiltinId.Weeknum, 2),
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
      [2023, 1, 3],
      [2023, 1, 3, 0],
      [2023, 1, 3, 2],
      [2024, 1, 1, 2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expectNumberCell(kernel, cellIndex(1, 0, width), 59)
    expectNumberCell(kernel, cellIndex(1, 1, width), 60)
    expectNumberCell(kernel, cellIndex(1, 2, width), 59 / 360)
    expectNumberCell(kernel, cellIndex(1, 3, width), 1)
  })

  it('returns #NUM for invalid YEARFRAC basis values on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 2, 0, 1, 1)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Yearfrac, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Yearfrac, 3), encodeRet()],
    ])
    const targetCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targetCells)
    const constants = packConstants([
      [45292, 45474, 5],
      [45292, 45474, -1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(targetCells)

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
  })

  it('keeps depreciation helper behavior stable across refactors', async () => {
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

  it('keeps discounted-security helper behavior stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 5, 2, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BuiltinId.Disc, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BuiltinId.Pricedisc, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BuiltinId.Yielddisc, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodeCall(BuiltinId.Tbillprice, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodeCall(BuiltinId.Tbillyield, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodeCall(BuiltinId.Tbilleq, 3),
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
        cellIndex(1, 5, width),
      ]),
    )
    const constants = packConstants([
      [2023, 1, 1, 4, 97, 100, 2],
      [2008, 2, 16, 3, 1, 0.0525, 100, 2, 99.795],
      [2008, 2, 16, 3, 1, 99.795, 100, 2],
      [2008, 3, 31, 6, 1, 0.09],
      [2008, 3, 31, 6, 1, 98.45],
      [2008, 3, 31, 6, 1, 0.0914],
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
      ]),
    )

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.12)
    expectNumberCell(kernel, cellIndex(1, 1, width), 99.79583333333333)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.05282257198685834)
    expectNumberCell(kernel, cellIndex(1, 3, width), 98.45)
    expectNumberCell(kernel, cellIndex(1, 4, width), 0.09141696292534264)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.09415149356594302)
  })

  it('keeps coupon-duration helper behavior stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 2, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(6),
        encodePushNumber(7),
        encodePushNumber(8),
        encodePushNumber(9),
        encodeCall(BuiltinId.Duration, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(6),
        encodePushNumber(7),
        encodePushNumber(8),
        encodePushNumber(9),
        encodeCall(BuiltinId.Mduration, 6),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    )
    const constants = packConstants([
      [2018, 7, 1, 2048, 1, 1, 0.08, 0.09, 2, 1],
      [2008, 1, 1, 2016, 1, 1, 0.08, 0.09, 2, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]))

    expectNumberCell(kernel, cellIndex(1, 0, width), 10.919145281591925)
    expectNumberCell(kernel, cellIndex(1, 1, width), 5.735669813918838)
  })
})
