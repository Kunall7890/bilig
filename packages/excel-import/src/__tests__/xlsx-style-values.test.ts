import { describe, expect, it } from 'vitest'
import { asArray, normalizeRgbColor, numberValue, recordChild, stringValue, toArgbColor } from '../xlsx-style-values.js'

describe('xlsx style value helpers', () => {
  it('normalizes scalar XML parser values conservatively', () => {
    expect(asArray(undefined)).toEqual([])
    expect(asArray('one')).toEqual(['one'])
    expect(asArray(['one', 'two'])).toEqual(['one', 'two'])
    expect(stringValue('font')).toBe('font')
    expect(stringValue(12)).toBeNull()
    expect(numberValue(' 12.5 ')).toBe(12.5)
    expect(numberValue('')).toBeNull()
    expect(numberValue('abc')).toBeNull()
  })

  it('reads child records without treating primitives as objects', () => {
    expect(recordChild({ color: { rgb: 'FF112233' } }, 'color')).toEqual({ rgb: 'FF112233' })
    expect(recordChild({ color: 'red' }, 'color')).toBeNull()
    expect(recordChild(null, 'color')).toBeNull()
  })

  it('normalizes RGB and ARGB color values used by imports and exports', () => {
    expect(normalizeRgbColor('112233')).toBe('#112233')
    expect(normalizeRgbColor('#AABBCC')).toBe('#aabbcc')
    expect(normalizeRgbColor('FF112233')).toBe('#112233')
    expect(normalizeRgbColor('bad')).toBeNull()
    expect(toArgbColor('#112233')).toBe('FF112233')
    expect(toArgbColor('bad')).toBeNull()
  })
})
