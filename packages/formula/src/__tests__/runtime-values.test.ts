import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { isArrayValue, scalarFromEvaluationResult, type EvaluationResult } from '../runtime-values.js'

describe('runtime value helpers', () => {
  it('detects array results and extracts their first scalar', () => {
    const arrayResult: EvaluationResult = {
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'first' },
        { tag: ValueTag.Number, value: 2 },
      ],
    }

    expect(isArrayValue(arrayResult)).toBe(true)
    expect(scalarFromEvaluationResult(arrayResult)).toEqual({ tag: ValueTag.String, value: 'first' })
  })

  it('passes scalar results through and treats empty arrays as empty cells', () => {
    const scalar = { tag: ValueTag.Boolean, value: true } as const

    expect(isArrayValue(scalar)).toBe(false)
    expect(scalarFromEvaluationResult(scalar)).toBe(scalar)
    expect(scalarFromEvaluationResult({ kind: 'array', rows: 0, cols: 0, values: [] })).toEqual({ tag: ValueTag.Empty })
  })
})
