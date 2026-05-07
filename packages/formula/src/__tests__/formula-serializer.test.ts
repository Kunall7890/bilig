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
})
