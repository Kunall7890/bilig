import { describe, expect, it } from 'vitest'

import { parseSnapshotRevisionHeader } from '../runtime-session.js'

describe('runtime session snapshot revision parsing', () => {
  it('accepts safe non-negative integer snapshot cursor headers', () => {
    expect(parseSnapshotRevisionHeader('0')).toBe(0)
    expect(parseSnapshotRevisionHeader('37')).toBe(37)
    expect(parseSnapshotRevisionHeader(' 451 ')).toBe(451)
    expect(parseSnapshotRevisionHeader(String(Number.MAX_SAFE_INTEGER))).toBe(Number.MAX_SAFE_INTEGER)
  })

  it.each([null, '', ' ', '-1', '+1', '01', '1.5', '37abc', String(Number.MAX_SAFE_INTEGER + 1)])(
    'rejects malformed snapshot cursor header %s',
    (value) => {
      expect(parseSnapshotRevisionHeader(value)).toBeNull()
    },
  )
})
