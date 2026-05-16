import { describe, expect, it } from 'vitest'

import { resolveBrowserTestCiMode, resolveBrowserTestConfiguredPort, resolveBrowserTestTimeoutMs } from '../browser-test-runner-config.ts'

describe('browser test runner config', () => {
  it('parses CI mode explicitly for browser runner behavior', () => {
    expect(resolveBrowserTestCiMode(undefined)).toBe(false)
    expect(resolveBrowserTestCiMode('')).toBe(false)
    expect(resolveBrowserTestCiMode('0')).toBe(false)
    expect(resolveBrowserTestCiMode('false')).toBe(false)
    expect(resolveBrowserTestCiMode('1')).toBe(true)
    expect(resolveBrowserTestCiMode('true')).toBe(true)
    expect(() => resolveBrowserTestCiMode('treu')).toThrow('CI must be "1", "true", "0", or "false" when set, got treu')
  })

  it('parses decimal startup timeouts without truncating malformed values', () => {
    expect(resolveBrowserTestTimeoutMs('120000', 300000)).toBe(120000)
    expect(resolveBrowserTestTimeoutMs('120000ms', 300000)).toBe(300000)
    expect(resolveBrowserTestTimeoutMs('0', 300000)).toBe(300000)
    expect(resolveBrowserTestTimeoutMs(undefined, 300000)).toBe(300000)
  })

  it('validates explicit browser test TCP ports before starting services', () => {
    expect(resolveBrowserTestConfiguredPort('4180', 'BILIG_E2E_WEB_PORT')).toBe('4180')
    expect(() => resolveBrowserTestConfiguredPort('4180abc', 'BILIG_E2E_WEB_PORT')).toThrow(
      'BILIG_E2E_WEB_PORT must be a TCP port between 1 and 65535, got 4180abc',
    )
    expect(() => resolveBrowserTestConfiguredPort('0', 'BILIG_E2E_WEB_PORT')).toThrow(
      'BILIG_E2E_WEB_PORT must be a TCP port between 1 and 65535, got 0',
    )
    expect(() => resolveBrowserTestConfiguredPort('70000', 'BILIG_E2E_WEB_PORT')).toThrow(
      'BILIG_E2E_WEB_PORT must be a TCP port between 1 and 65535, got 70000',
    )
  })
})
