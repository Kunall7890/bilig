import { describe, expect, it } from 'vitest'

import { readLargeSimpleCellValueFromTextRange } from '../xlsx-large-simple-cell-value-scan.js'

const encoder = new TextEncoder()

describe('large simple XLSX cell value byte scan', () => {
  it('parses common decimal numbers with the same value as JavaScript Number', () => {
    const samples = ['0.1', '-0.1', '+0.1', '.5', '-.5', '1.', '-0', '-0.0', '123456789012.34', '-9876543210.1234', '  42.125  ']

    for (const sample of samples) {
      expectSameNumber(readNumber(sample), Number(sample.trim()))
    }
  })

  it('keeps decimal parsing exact for a deterministic broad sample below the safe fast-path digit limit', () => {
    let state = 0x13579bdf
    for (let sampleIndex = 0; sampleIndex < 2_000; sampleIndex += 1) {
      state = nextState(state)
      const digitLength = 1 + (state % 15)
      let digits = ''
      for (let digitIndex = 0; digitIndex < digitLength; digitIndex += 1) {
        state = nextState(state)
        digits += String(state % 10)
      }
      state = nextState(state)
      const decimalOffset = state % (digitLength + 1)
      const sign = state & 1 ? '-' : ''
      const raw = `${sign}${digits.slice(0, decimalOffset)}.${digits.slice(decimalOffset)}`

      expectSameNumber(readNumber(raw), Number(raw))
    }
  })

  it('preserves fallback semantics for exponents, long decimals, and XML-escaped values', () => {
    const samples = ['1.25e3', '-6.02214076E23', '1234567890123456.7', '1&#46;5']

    for (const sample of samples) {
      expectSameNumber(readNumber(sample), Number(sample.replace('&#46;', '.')))
    }
  })
})

function readNumber(raw: string): unknown {
  const bytes = encoder.encode(raw)
  return readLargeSimpleCellValueFromTextRange(bytes, { start: 0, end: bytes.length }, null, [])
}

function nextState(state: number): number {
  return (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0
}

function expectSameNumber(actual: unknown, expected: number): void {
  expect(typeof actual).toBe('number')
  expect(Object.is(actual, expected)).toBe(true)
}
