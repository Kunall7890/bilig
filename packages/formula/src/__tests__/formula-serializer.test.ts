import { ErrorCode } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { serializeFormula } from '../formula-serializer.js'

describe('formula serializer', () => {
  it('serializes literals, sheet-qualified refs, and binary precedence', () => {
    expect(serializeFormula({ kind: 'StringLiteral', value: 'a"b' })).toBe('"a""b"')
    expect(serializeFormula({ kind: 'ErrorLiteral', code: ErrorCode.Spill })).toBe('#SPILL!')
    expect(serializeFormula({ kind: 'CellRef', sheetName: 'Sheet 2', ref: '$B3' })).toBe("'Sheet 2'!$B3")
    expect(
      serializeFormula({
        kind: 'BinaryExpr',
        operator: '*',
        left: {
          kind: 'BinaryExpr',
          operator: '+',
          left: { kind: 'NumberLiteral', value: 1 },
          right: { kind: 'NumberLiteral', value: 2 },
        },
        right: { kind: 'NumberLiteral', value: 3 },
      }),
    ).toBe('(1+2)*3')
  })

  it('serializes omitted arguments, array constants, fallback errors, and 3D range edge prefixes', () => {
    expect(serializeFormula({ kind: 'OmittedArgument' })).toBe('')
    expect(serializeFormula({ kind: 'ErrorLiteral', code: 999 })).toBe('#ERROR!')
    expect(
      serializeFormula({
        kind: 'ArrayConstant',
        rows: [
          [
            { kind: 'NumberLiteral', value: 1 },
            { kind: 'BooleanLiteral', value: true },
            { kind: 'StringLiteral', value: 'x' },
          ],
          [{ kind: 'ErrorLiteral', code: ErrorCode.Div0 }, { kind: 'OmittedArgument' }],
        ],
      }),
    ).toBe('{1,TRUE,"x";#DIV/0!,}')
    expect(
      serializeFormula({
        kind: 'RangeRef',
        sheetEndName: 'Sheet 3',
        start: 'A1',
        end: 'B2',
      }),
    ).toBe('A1:B2')
    expect(
      serializeFormula({
        kind: 'RangeRef',
        sheetName: 'Sheet 1',
        sheetEndName: 'Sheet 3',
        start: 'A1',
        end: 'B2',
      }),
    ).toBe("'Sheet 1':'Sheet 3'!A1:B2")
  })
})
