import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getTextBuiltin } from '../builtins/text.js'

const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })

const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

describe('text Unicode domain errors', () => {
  it('returns #VALUE! for partial surrogate text in UNICODE', () => {
    expect(getTextBuiltin('UNICODE')?.(text('\ud800'))).toEqual(valueError)
    expect(getTextBuiltin('UNICODE')?.(text('\udc00'))).toEqual(valueError)
  })
})
