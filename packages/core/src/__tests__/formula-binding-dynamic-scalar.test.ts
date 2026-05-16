import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { branchChoice, criterionMatches } from '../engine/services/formula-binding-dynamic-scalar.js'

const numberValue = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const booleanValue = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const stringValue = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })

describe('formula binding dynamic scalar helpers', () => {
  it('classifies branch choices without treating text or errors as decisive branches', () => {
    expect(branchChoice(booleanValue(true))).toBe('truthy')
    expect(branchChoice(numberValue(42))).toBe('truthy')
    expect(branchChoice(booleanValue(false))).toBe('falsy')
    expect(branchChoice(numberValue(0))).toBe('falsy')
    expect(branchChoice({ tag: ValueTag.Empty })).toBe('falsy')
    expect(branchChoice(stringValue('TRUE'))).toBe('no-branch')
    expect(branchChoice({ tag: ValueTag.Error, code: ErrorCode.Value })).toBe('no-branch')
    expect(branchChoice(undefined)).toBe('unknown')
  })

  it('matches criteria strings against numeric and text values', () => {
    expect(criterionMatches(numberValue(4), stringValue('>3'))).toBe(true)
    expect(criterionMatches(numberValue(4), stringValue('<=3'))).toBe(false)
    expect(criterionMatches(stringValue('West'), stringValue('west'))).toBe(true)
    expect(criterionMatches(stringValue('East'), stringValue('<>West'))).toBe(true)
    expect(criterionMatches(stringValue(''), stringValue('='))).toBe(true)
  })

  it('does not match criteria when either side is an error', () => {
    expect(criterionMatches({ tag: ValueTag.Error, code: ErrorCode.Value }, stringValue('>0'))).toBeUndefined()
    expect(criterionMatches(numberValue(1), { tag: ValueTag.Error, code: ErrorCode.Value })).toBeUndefined()
  })
})
