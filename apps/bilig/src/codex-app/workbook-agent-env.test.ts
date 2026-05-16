import { describe, expect, it } from 'vitest'

import { parsePositiveIntegerEnv } from './workbook-agent-env.js'

describe('workbook agent env parsing', () => {
  it('uses the fallback when no limit override is configured', () => {
    expect(parsePositiveIntegerEnv(undefined, 4, 'BILIG_CODEX_MAX_CLIENTS')).toBe(4)
    expect(parsePositiveIntegerEnv('', 4, 'BILIG_CODEX_MAX_CLIENTS')).toBe(4)
  })

  it('accepts explicit positive safe integer limits', () => {
    expect(parsePositiveIntegerEnv('1', 4, 'BILIG_CODEX_MAX_CLIENTS')).toBe(1)
    expect(parsePositiveIntegerEnv('16', 4, 'BILIG_CODEX_MAX_CLIENTS')).toBe(16)
  })

  it.each(['0', '-1', '4abc', '1.5'])('rejects malformed workbook agent limit %s', (value) => {
    expect(() => parsePositiveIntegerEnv(value, 4, 'BILIG_CODEX_MAX_CLIENTS')).toThrow(
      `BILIG_CODEX_MAX_CLIENTS must be a positive integer, got ${value}`,
    )
  })

  it('rejects unsafe workbook agent limits', () => {
    expect(() => parsePositiveIntegerEnv('9007199254740992', 4, 'BILIG_CODEX_MAX_CLIENTS')).toThrow(
      'BILIG_CODEX_MAX_CLIENTS must be a safe integer, got 9007199254740992',
    )
  })
})
