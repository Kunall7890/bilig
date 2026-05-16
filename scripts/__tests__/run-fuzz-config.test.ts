import { describe, expect, it } from 'vitest'
import { buildVitestFuzzCommand, parseFuzzMode, resolveVitestFuzzMaxWorkers } from '../run-fuzz-config.js'

describe('run fuzz config', () => {
  it('resolves explicit fuzz modes without silently downgrading unknown values', () => {
    expect(parseFuzzMode(undefined)).toBe('default')
    expect(parseFuzzMode('default')).toBe('default')
    expect(parseFuzzMode('main')).toBe('main')
    expect(parseFuzzMode('nightly')).toBe('nightly')
    expect(parseFuzzMode('replay')).toBe('replay')
    expect(() => parseFuzzMode('mainn')).toThrow('Fuzz mode must be "default", "main", "nightly", or "replay", got mainn')
  })

  it('caps vitest fuzz workers to a conservative subset of host parallelism', () => {
    expect(resolveVitestFuzzMaxWorkers(1)).toBe(1)
    expect(resolveVitestFuzzMaxWorkers(2)).toBe(1)
    expect(resolveVitestFuzzMaxWorkers(3)).toBe(2)
    expect(resolveVitestFuzzMaxWorkers(8)).toBe(2)
    expect(resolveVitestFuzzMaxWorkers(32)).toBe(2)
  })

  it('appends the maxWorkers flag to the vitest fuzz command', () => {
    expect(buildVitestFuzzCommand(['packages/core/src/__tests__/snapshot-wire-parity.fuzz.test.ts'], 8)).toEqual([
      'pnpm',
      'exec',
      'vitest',
      'run',
      'packages/core/src/__tests__/snapshot-wire-parity.fuzz.test.ts',
      '--maxWorkers',
      '2',
    ])
  })
})
