import { ValueTag } from './protocol'
import { toNumberExact } from './operands'

export function coerceLength(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue
  }
  const numeric = toNumberExact(tag, value)
  if (isNaN(numeric)) {
    return i32.MIN_VALUE
  }
  const truncated = <i32>numeric
  return truncated >= 0 ? truncated : i32.MIN_VALUE
}

export function coercePositiveStart(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue
  }
  const numeric = toNumberExact(tag, value)
  if (isNaN(numeric)) {
    return i32.MIN_VALUE
  }
  const truncated = <i32>numeric
  return truncated >= 1 ? truncated : i32.MIN_VALUE
}

export function coerceNonNegativeLength(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value)
  if (isNaN(numeric)) {
    return i32.MIN_VALUE
  }
  const truncated = <i32>numeric
  return truncated >= 0 ? truncated : i32.MIN_VALUE
}

export function excelTrim(input: string): string {
  let start = 0
  let end = input.length
  while (start < end && input.charCodeAt(start) == 32) {
    start += 1
  }
  while (end > start && input.charCodeAt(end - 1) == 32) {
    end -= 1
  }
  let result = ''
  let previousSpace = false
  for (let index = start; index < end; index += 1) {
    const char = input.charCodeAt(index)
    if (char == 32) {
      if (!previousSpace) {
        result += ' '
      }
      previousSpace = true
      continue
    }
    previousSpace = false
    result += String.fromCharCode(char)
  }
  return result
}

function hasSearchSyntax(pattern: string): bool {
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern.charCodeAt(index)
    if (char == 126 || char == 42 || char == 63) {
      return true
    }
  }
  return false
}

function isHighSurrogate(char: i32): bool {
  return char >= 0xd800 && char <= 0xdbff
}

function isLowSurrogate(char: i32): bool {
  return char >= 0xdc00 && char <= 0xdfff
}

function scalarCodeUnitStep(text: string, index: i32): i32 {
  if (index < 0 || index >= text.length) {
    return 0
  }
  const char = text.charCodeAt(index)
  if (isHighSurrogate(char) && index + 1 < text.length && isLowSurrogate(text.charCodeAt(index + 1))) {
    return 2
  }
  return 1
}

function sameScalarAt(left: string, leftIndex: i32, right: string, rightIndex: i32): bool {
  if (leftIndex >= left.length || rightIndex >= right.length) {
    return false
  }
  const leftStep = scalarCodeUnitStep(left, leftIndex)
  const rightStep = scalarCodeUnitStep(right, rightIndex)
  if (leftStep != rightStep) {
    return false
  }
  for (let offset = 0; offset < leftStep; offset += 1) {
    if (left.charCodeAt(leftIndex + offset) != right.charCodeAt(rightIndex + offset)) {
      return false
    }
  }
  return true
}

export function scalarTextLength(text: string): i32 {
  let length = 0
  let index = 0
  while (index < text.length) {
    index += scalarCodeUnitStep(text, index)
    length += 1
  }
  return length
}

export function codeUnitIndexFromScalarIndex(text: string, scalarIndex: i32): i32 {
  if (scalarIndex <= 0) {
    return 0
  }
  let index = 0
  let currentScalarIndex = 0
  while (index < text.length && currentScalarIndex < scalarIndex) {
    index += scalarCodeUnitStep(text, index)
    currentScalarIndex += 1
  }
  return index
}

function scalarIndexFromCodeUnitIndex(text: string, codeUnitIndex: i32): i32 {
  const target = max<i32>(0, min<i32>(codeUnitIndex, text.length))
  let index = 0
  let scalarIndex = 0
  while (index < target) {
    index += scalarCodeUnitStep(text, index)
    scalarIndex += 1
  }
  return scalarIndex
}

export function sliceTextScalars(text: string, startIndex: i32, count: i32): string {
  if (count <= 0) {
    return ''
  }
  const startCodeUnitIndex = codeUnitIndexFromScalarIndex(text, startIndex)
  const endCodeUnitIndex = codeUnitIndexFromScalarIndex(text, startIndex + count)
  return text.slice(startCodeUnitIndex, endCodeUnitIndex)
}

function wildcardMatchAt(pattern: string, haystack: string, patternIndex: i32, haystackIndex: i32): bool {
  let p = patternIndex
  let h = haystackIndex
  while (p < pattern.length) {
    const char = pattern.charCodeAt(p)
    if (char == 126) {
      const nextIndex = p + 1
      if (h >= haystack.length || !sameScalarAt(pattern, nextIndex < pattern.length ? nextIndex : p, haystack, h)) {
        return false
      }
      p = nextIndex < pattern.length ? nextIndex + scalarCodeUnitStep(pattern, nextIndex) : nextIndex
      h += scalarCodeUnitStep(haystack, h)
      continue
    }
    if (char == 42) {
      let nextPatternIndex = p + 1
      while (nextPatternIndex < pattern.length && pattern.charCodeAt(nextPatternIndex) == 42) {
        nextPatternIndex += 1
      }
      if (nextPatternIndex >= pattern.length) {
        return true
      }
      let scan = h
      while (scan <= haystack.length) {
        if (wildcardMatchAt(pattern, haystack, nextPatternIndex, scan)) {
          return true
        }
        if (scan >= haystack.length) {
          break
        }
        scan += scalarCodeUnitStep(haystack, scan)
      }
      return false
    }
    if (h >= haystack.length) {
      return false
    }
    if (char == 63) {
      p += 1
      h += scalarCodeUnitStep(haystack, h)
      continue
    }
    if (!sameScalarAt(pattern, p, haystack, h)) {
      return false
    }
    p += scalarCodeUnitStep(pattern, p)
    h += scalarCodeUnitStep(haystack, h)
  }
  return true
}

function wildcardMatchAtCodeUnits(pattern: string, haystack: string, patternIndex: i32, haystackIndex: i32): bool {
  let p = patternIndex
  let h = haystackIndex
  while (p < pattern.length) {
    const char = pattern.charCodeAt(p)
    if (char == 126) {
      const nextIndex = p + 1
      const nextChar = nextIndex < pattern.length ? pattern.charCodeAt(nextIndex) : 126
      if (h >= haystack.length || haystack.charCodeAt(h) != nextChar) {
        return false
      }
      p = nextIndex < pattern.length ? nextIndex + 1 : nextIndex
      h += 1
      continue
    }
    if (char == 42) {
      let nextPatternIndex = p + 1
      while (nextPatternIndex < pattern.length && pattern.charCodeAt(nextPatternIndex) == 42) {
        nextPatternIndex += 1
      }
      if (nextPatternIndex >= pattern.length) {
        return true
      }
      for (let scan = h; scan <= haystack.length; scan += 1) {
        if (wildcardMatchAtCodeUnits(pattern, haystack, nextPatternIndex, scan)) {
          return true
        }
      }
      return false
    }
    if (h >= haystack.length) {
      return false
    }
    if (char == 63) {
      p += 1
      h += 1
      continue
    }
    if (haystack.charCodeAt(h) != char) {
      return false
    }
    p += 1
    h += 1
  }
  return true
}

function findPositionInternal(
  needle: string,
  haystack: string,
  start: i32,
  caseSensitive: bool,
  wildcardAware: bool,
  scalarPositions: bool,
): i32 {
  const textLength = scalarPositions ? scalarTextLength(haystack) : haystack.length
  if (start > textLength) {
    return i32.MIN_VALUE
  }
  if (needle.length == 0) {
    return start
  }
  const startIndex = scalarPositions ? codeUnitIndexFromScalarIndex(haystack, start - 1) : start - 1
  const normalizedHaystack = caseSensitive ? haystack : haystack.toLowerCase()
  const normalizedNeedle = caseSensitive ? needle : needle.toLowerCase()
  if (wildcardAware && hasSearchSyntax(normalizedNeedle)) {
    for (let offset = start - 1; offset <= textLength; offset += 1) {
      const index = scalarPositions ? codeUnitIndexFromScalarIndex(normalizedHaystack, offset) : offset
      if (
        scalarPositions
          ? wildcardMatchAt(normalizedNeedle, normalizedHaystack, 0, index)
          : wildcardMatchAtCodeUnits(normalizedNeedle, normalizedHaystack, 0, index)
      ) {
        return scalarPositions ? offset + 1 : index + 1
      }
    }
    return i32.MIN_VALUE
  }
  const found = normalizedHaystack.indexOf(normalizedNeedle, startIndex)
  return found < 0 ? i32.MIN_VALUE : scalarPositions ? scalarIndexFromCodeUnitIndex(haystack, found) + 1 : found + 1
}

export function findPosition(needle: string, haystack: string, start: i32, caseSensitive: bool, wildcardAware: bool): i32 {
  return findPositionInternal(needle, haystack, start, caseSensitive, wildcardAware, true)
}

export function findPositionCodeUnits(needle: string, haystack: string, start: i32, caseSensitive: bool, wildcardAware: bool): i32 {
  return findPositionInternal(needle, haystack, start, caseSensitive, wildcardAware, false)
}
