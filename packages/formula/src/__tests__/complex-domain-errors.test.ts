import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })

describe('complex formula error precedence', () => {
  it('preserves incoming errors before complex parsing and domain checks', () => {
    expect(getBuiltin('IMABS')?.(err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(getBuiltin('IMREAL')?.(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(getBuiltin('IMDIV')?.(text('1+i'), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(getBuiltin('COMPLEX')?.(num(1), num(2), err(ErrorCode.Div0))).toEqual(err(ErrorCode.Div0))
  })
})
