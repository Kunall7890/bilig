import { describe, expect, it } from 'vitest'

import { buildVitestArgBatches, buildVitestArgs, readVitestBatchCooldownMs } from '../run-vitest.ts'

describe('run-vitest wrapper arguments', () => {
  it('bounds Vitest workers in CI by default', () => {
    expect(buildVitestArgs(['--run', 'sample.test.ts'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      'sample.test.ts',
      '--maxWorkers',
      '1',
    ])
  })

  it('preserves an explicit maxWorkers flag', () => {
    expect(buildVitestArgs(['--run', '--maxWorkers=1'], { BILIG_CI_PROFILE: 'fast' })).toEqual(['--run', '--maxWorkers=1'])
  })

  it('allows CI worker limit overrides', () => {
    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '3',
      }),
    ).toEqual(['--run', '--maxWorkers', '3'])
  })

  it('splits large CI run file lists into serial batches', () => {
    const files = Array.from({ length: 13 }, (_, index) => `test-${index + 1}.test.ts`)

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
      }),
    ).toEqual([
      ['--run', ...files.slice(0, 12), '--maxWorkers', '1'],
      ['--run', files[12], '--maxWorkers', '1'],
    ])
  })

  it('allows CI file chunk size overrides', () => {
    expect(
      buildVitestArgBatches(['--run', 'a.test.ts', 'b.test.ts', 'c.test.ts'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '2',
      }),
    ).toEqual([
      ['--run', 'a.test.ts', 'b.test.ts', '--maxWorkers', '1'],
      ['--run', 'c.test.ts', '--maxWorkers', '1'],
    ])
  })

  it('does not split run arguments that include flags', () => {
    expect(
      buildVitestArgBatches(['--run', 'sample.test.ts', '--reporter=dot'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '1',
      }),
    ).toEqual([['--run', 'sample.test.ts', '--reporter=dot', '--maxWorkers', '1']])
  })

  it('adds a short CI-only cooldown between split batches', () => {
    expect(readVitestBatchCooldownMs({})).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast' })).toBe(1000)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '0' })).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '2500' })).toBe(2500)
  })
})
