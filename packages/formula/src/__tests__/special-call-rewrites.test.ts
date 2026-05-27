import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@bilig/protocol'
import { rewriteSpecialCall } from '../special-call-rewrites.js'

describe('special call rewrites', () => {
  it('rewrites missing IF false branches to explicit FALSE literals', () => {
    expect(
      rewriteSpecialCall({
        kind: 'CallExpr',
        callee: 'IF',
        args: [
          { kind: 'CellRef', ref: 'A1' },
          { kind: 'NumberLiteral', value: 1 },
        ],
      }),
    ).toEqual({
      kind: 'CallExpr',
      callee: 'IF',
      args: [
        { kind: 'CellRef', ref: 'A1' },
        { kind: 'NumberLiteral', value: 1 },
        { kind: 'BooleanLiteral', value: false },
      ],
    })
  })

  it('rewrites TRUE and FALSE into literals or VALUE errors', () => {
    expect(rewriteSpecialCall({ kind: 'CallExpr', callee: 'TRUE', args: [] })).toEqual({
      kind: 'BooleanLiteral',
      value: true,
    })
    expect(
      rewriteSpecialCall({
        kind: 'CallExpr',
        callee: 'FALSE',
        args: [{ kind: 'NumberLiteral', value: 1 }],
      }),
    ).toEqual({ kind: 'ErrorLiteral', code: ErrorCode.Value })
  })

  it('rewrites IFS and SWITCH into nested core expressions', () => {
    expect(
      rewriteSpecialCall({
        kind: 'CallExpr',
        callee: 'IFS',
        args: [
          { kind: 'BooleanLiteral', value: false },
          { kind: 'NumberLiteral', value: 1 },
          { kind: 'BooleanLiteral', value: true },
          { kind: 'NumberLiteral', value: 2 },
        ],
      }),
    ).toEqual({
      kind: 'CallExpr',
      callee: 'IF',
      args: [
        { kind: 'BooleanLiteral', value: false },
        { kind: 'NumberLiteral', value: 1 },
        {
          kind: 'CallExpr',
          callee: 'IF',
          args: [
            { kind: 'BooleanLiteral', value: true },
            { kind: 'NumberLiteral', value: 2 },
            { kind: 'ErrorLiteral', code: ErrorCode.NA },
          ],
        },
      ],
    })

    const switchRewrite = rewriteSpecialCall({
      kind: 'CallExpr',
      callee: 'SWITCH',
      args: [
        { kind: 'StringLiteral', value: 'b' },
        { kind: 'StringLiteral', value: 'a' },
        { kind: 'NumberLiteral', value: 1 },
        { kind: 'StringLiteral', value: 'b' },
        { kind: 'NumberLiteral', value: 2 },
        { kind: 'NumberLiteral', value: 9 },
      ],
    })
    expect(switchRewrite).toMatchObject({
      kind: 'CallExpr',
      callee: 'IF',
      args: [
        { kind: 'CallExpr', callee: 'OR' },
        { kind: 'NumberLiteral', value: 1 },
        { kind: 'CallExpr', callee: 'IF' },
      ],
    })
    expect(switchRewrite?.kind === 'CallExpr' ? switchRewrite.args[2] : undefined).toMatchObject({
      kind: 'CallExpr',
      callee: 'IF',
      args: [
        { kind: 'CallExpr', callee: 'OR' },
        { kind: 'NumberLiteral', value: 2 },
        { kind: 'NumberLiteral', value: 9 },
      ],
    })
    expect(
      rewriteSpecialCall({
        kind: 'CallExpr',
        callee: 'XOR',
        args: [
          { kind: 'BooleanLiteral', value: true },
          { kind: 'BooleanLiteral', value: false },
        ],
      }),
    ).toBeUndefined()
  })

  it('returns VALUE or NA errors for invalid special-call arities', () => {
    expect(
      rewriteSpecialCall({
        kind: 'CallExpr',
        callee: 'IFS',
        args: [{ kind: 'BooleanLiteral', value: true }],
      }),
    ).toEqual({ kind: 'ErrorLiteral', code: ErrorCode.Value })
    expect(
      rewriteSpecialCall({
        kind: 'CallExpr',
        callee: 'SWITCH',
        args: [
          { kind: 'NumberLiteral', value: 1 },
          { kind: 'NumberLiteral', value: 2 },
        ],
      }),
    ).toEqual({ kind: 'ErrorLiteral', code: ErrorCode.Value })
    expect(rewriteSpecialCall({ kind: 'CallExpr', callee: 'XOR', args: [] })).toBeUndefined()
    expect(rewriteSpecialCall({ kind: 'CallExpr', callee: 'UNRELATED', args: [] })).toBeUndefined()
  })
})
