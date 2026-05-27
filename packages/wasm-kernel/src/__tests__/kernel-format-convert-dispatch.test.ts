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

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

describe('wasm kernel format and conversion dispatch', () => {
  it('keeps address, dollar, radix, and unit conversion dispatch stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(96, 8, 13, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 7, 11, 21, 23, 25, 26, 27, 30, 33]),
      Uint32Array.from([7, 4, 10, 2, 2, 1, 1, 3, 3, 3]),
      Uint16Array.from(Array.from("O'Brien00FF1111111111mikmFCDEMEURFRF", (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Address, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushString(0),
        encodeCall(BuiltinId.Address, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollar, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarde, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarfr, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Base, 3), encodeRet()],
      [encodePushString(1), encodePushNumber(0), encodeCall(BuiltinId.Decimal, 2), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Bin2dec, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dec2hex, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(3), encodePushString(4), encodeCall(BuiltinId.Convert, 3), encodeRet()],
      [encodePushNumber(0), encodePushString(5), encodePushString(6), encodeCall(BuiltinId.Convert, 3), encodeRet()],
      [encodePushNumber(0), encodePushString(7), encodePushString(8), encodeCall(BuiltinId.Euroconvert, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushString(9),
        encodePushString(7),
        encodePushBoolean(true),
        encodePushNumber(1),
        encodeCall(BuiltinId.Euroconvert, 5),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [12, 3],
      [2, 28, 3, 1],
      [-1234.5, 1],
      [1.08, 16],
      [1.5, 16],
      [255, 16, 4],
      [16],
      [],
      [255, 4],
      [6],
      [68],
      [1.2],
      [1, 3],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(1, 3, width), 1.5)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1.08)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(1, 6, width), 255)
    expectNumberCell(kernel, cellIndex(1, 7, width), -1)
    expect(kernel.readTags()[cellIndex(1, 8, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(1, 9, width), 9.656064)
    expectNumberCell(kernel, cellIndex(1, 10, width), 20)
    expectNumberCell(kernel, cellIndex(1, 11, width), 0.61)
    expectNumberCell(kernel, cellIndex(1, 12, width), 0.29728616, 8)
    expect(kernel.readOutputStrings()).toEqual(['$C$12', "'O''Brien'!$AB2", '-$1,234.5', '00FF', '00FF'])
  })

  it('uses the documented logical ADDRESS A1 flag on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 1, 8, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodeCall(BuiltinId.Address, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Address, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.Address, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Address, 4), encodeRet()],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 4 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputs)
    const constants = packConstants([
      [2, 3, 2],
      [2, 3, 2, 0],
      [2, 3, 2],
      [2, 3, 2, 2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputs)

    expect(kernel.readOutputStrings()).toEqual(['R2C[3]', 'R2C[3]', 'C$2', 'C$2'])
  })

  it('quotes ADDRESS sheet names only when formula syntax requires quoting on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 4, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 6, 17]),
      Uint32Array.from([6, 11, 7]),
      Uint16Array.from(Array.from("Sheet2EXCEL SHEETO'Brien", (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodePushString(0),
        encodeCall(BuiltinId.Address, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodePushString(1),
        encodeCall(BuiltinId.Address, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(0),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodePushString(2),
        encodeCall(BuiltinId.Address, 5),
        encodeRet(),
      ],
    ])
    const outputs = Uint32Array.from(Array.from({ length: 3 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputs)
    const constants = packConstants([
      [1, 1, 1],
      [2, 3, 1],
      [1, 1, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputs)

    expect(kernel.readOutputStrings()).toEqual(['Sheet2!$A$1', "'EXCEL SHEET'!R2C3", "'O''Brien'!$A$1"])
  })

  it('formats currency text with the maximum supported decimal width on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 2, 1, 1, 1)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))
    const packed = packPrograms([[encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollar, 2), encodeRet()]])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([cellIndex(1, 0, width)]))
    const constants = packConstants([[1, 127]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual([`$1.${'0'.repeat(127)}`])
  })

  it('preserves error codes for invalid conversion inputs', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 3, 2, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 2, 5]),
      Uint32Array.from([2, 3, 3]),
      Uint16Array.from(Array.from('ftsecBAD', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushString(0), encodePushString(1), encodeCall(BuiltinId.Convert, 3), encodeRet()],
      [encodePushNumber(1), encodePushString(2), encodePushString(1), encodeCall(BuiltinId.Euroconvert, 3), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodePushString(0), encodePushString(1), encodeCall(BuiltinId.Convert, 3), encodeRet()],
      [encodePushNumber(1), encodePushError(ErrorCode.Name), encodePushString(1), encodeCall(BuiltinId.Convert, 3), encodeRet()],
      [encodePushNumber(1), encodePushError(ErrorCode.NA), encodePushString(1), encodeCall(BuiltinId.Euroconvert, 3), encodeRet()],
      [
        encodePushNumber(1),
        encodePushString(2),
        encodePushString(1),
        encodePushBoolean(false),
        encodePushError(ErrorCode.Name),
        encodeCall(BuiltinId.Euroconvert, 5),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[2.5], [1], [], [1], [1], [1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 0, width)]).toBe(ErrorCode.NA)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 1, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 2, width)]).toBe(ErrorCode.Ref)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Name)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 4, width)]).toBe(ErrorCode.NA)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 5, width)]).toBe(ErrorCode.Name)
  })

  it('preserves error codes for address and dollar-format inputs on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 0, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushError(ErrorCode.Ref), encodePushNumber(0), encodeCall(BuiltinId.Address, 2), encodeRet()],
      [encodePushNumber(0), encodePushError(ErrorCode.NA), encodeCall(BuiltinId.Address, 2), encodeRet()],
      [encodePushError(ErrorCode.Name), encodePushNumber(1), encodeCall(BuiltinId.Dollar, 2), encodeRet()],
      [encodePushNumber(0), encodePushError(ErrorCode.Ref), encodeCall(BuiltinId.Dollar, 2), encodeRet()],
      [encodePushError(ErrorCode.NA), encodePushNumber(2), encodeCall(BuiltinId.Dollarde, 2), encodeRet()],
      [encodePushNumber(3), encodePushError(ErrorCode.Name), encodeCall(BuiltinId.Dollarfr, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[1], [2], [16], [1.5]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))))

    const expectedErrors = [ErrorCode.Ref, ErrorCode.NA, ErrorCode.Name, ErrorCode.Ref, ErrorCode.NA, ErrorCode.Name]
    for (let index = 0; index < expectedErrors.length; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(expectedErrors[index])
    }
  })

  it('matches Microsoft Excel BASE numeric domain errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 0, 6, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Base, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Base, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Base, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Base, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Base, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Base, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [-1, 16],
      [2 ** 53, 16],
      [31, 1],
      [31, 37],
      [31, 16, -1],
      [31, 16, 256],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))))

    for (let index = 0; index < 6; index += 1) {
      expect(kernel.readTags()[cellIndex(1, index, width)]).toBe(ValueTag.Error)
      expect(kernel.readErrors()[cellIndex(1, index, width)]).toBe(ErrorCode.Num)
    }
  })

  it('matches Microsoft Excel DOLLARDE and DOLLARFR denominator semantics on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 0, 8, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarde, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarfr, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarde, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarfr, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarde, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarfr, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarde, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarfr, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [1.06, 12],
      [1.5, 12],
      [1.02, 12.9],
      [1.5, 12.9],
      [1.5, -1],
      [1.5, -1],
      [1.5, 0.5],
      [1.5, 0.5],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 8 }, (_, index) => cellIndex(1, index, width))))

    expectNumberCell(kernel, cellIndex(1, 0, width), 1.5)
    expectNumberCell(kernel, cellIndex(1, 1, width), 1.06)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1 + 2 / 12)
    expectNumberCell(kernel, cellIndex(1, 3, width), 1.06)

    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 4, width)]).toBe(ErrorCode.Num)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 5, width)]).toBe(ErrorCode.Num)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 6, width)]).toBe(ErrorCode.Div0)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 7, width)]).toBe(ErrorCode.Div0)
  })

  it('renders NULL error labels on the wasm VALUETOTEXT path', async () => {
    const kernel = await createKernel()
    const width = 2
    kernel.init(4, 1, 0, 1, 1)
    kernel.writeCells(new Uint8Array(4), new Float64Array(4), new Uint32Array(4), new Uint16Array(4))

    const packed = packPrograms([[encodePushError(ErrorCode.Null), encodeCall(BuiltinId.Valuetotext, 1), encodeRet()]])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([cellIndex(1, 0, width)]))
    kernel.uploadConstants(Float64Array.from([]), Uint32Array.from([0]), Uint32Array.from([0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['#NULL!'])
  })
})
