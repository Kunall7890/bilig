import { describe, expect, it } from 'vitest'

import { parseStrictBooleanEnvFlag } from '../strict-env.js'

describe('strict environment parsing', () => {
  it('uses the provided fallback for unset and empty boolean flags', () => {
    expect(parseStrictBooleanEnvFlag(undefined, 'BILIG_TEST_FLAG', false)).toBe(false)
    expect(parseStrictBooleanEnvFlag('', 'BILIG_TEST_FLAG', true)).toBe(true)
  })

  it('accepts the explicit boolean values used by repo scripts', () => {
    expect(parseStrictBooleanEnvFlag('1', 'BILIG_TEST_FLAG', false)).toBe(true)
    expect(parseStrictBooleanEnvFlag('true', 'BILIG_TEST_FLAG', false)).toBe(true)
    expect(parseStrictBooleanEnvFlag('0', 'BILIG_TEST_FLAG', true)).toBe(false)
    expect(parseStrictBooleanEnvFlag('false', 'BILIG_TEST_FLAG', true)).toBe(false)
  })

  it('rejects ambiguous boolean values with the flag name', () => {
    expect(() => parseStrictBooleanEnvFlag('yes', 'BILIG_TEST_FLAG', false)).toThrow(
      'BILIG_TEST_FLAG must be "1", "true", "0", or "false" when set, got yes',
    )
  })
})
