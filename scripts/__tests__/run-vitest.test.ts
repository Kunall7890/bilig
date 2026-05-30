import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  buildVitestArgBatches,
  buildVitestArgs,
  isBroadCorpusVitestRun,
  readVitestBatchCooldownMs,
  resolveVitestBin,
} from '../run-vitest.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readPackageScripts(packageJsonPath: string): Record<string, string> {
  const manifest: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  if (!isRecord(manifest) || !isRecord(manifest.scripts)) {
    throw new Error(`${packageJsonPath} must define package scripts`)
  }
  for (const [name, script] of Object.entries(manifest.scripts)) {
    if (typeof script !== 'string') {
      throw new Error(`${packageJsonPath} script ${name} must be a string`)
    }
  }
  return manifest.scripts
}

describe('run-vitest wrapper arguments', () => {
  it('bounds Vitest workers in CI by default', () => {
    expect(buildVitestArgs(['--run', 'sample.test.ts'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      'sample.test.ts',
      '--maxWorkers',
      '1',
      '--pool',
      'forks',
      '--configLoader',
      'runner',
      '--reporter',
      'verbose',
    ])
  })

  it('preserves an explicit maxWorkers flag', () => {
    expect(buildVitestArgs(['--run', '--maxWorkers=1'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      '--maxWorkers=1',
      '--pool',
      'forks',
      '--configLoader',
      'runner',
      '--reporter',
      'verbose',
    ])
  })

  it('allows CI worker limit overrides', () => {
    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '3',
      }),
    ).toEqual(['--run', '--maxWorkers', '3', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'])
  })

  it('ignores malformed CI worker limit overrides instead of forwarding them', () => {
    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '3abc',
      }),
    ).toEqual(['--run', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'])

    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '0',
      }),
    ).toEqual(['--run', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'])
  })

  it('uses the verbose reporter in CI unless a reporter is already explicit', () => {
    expect(buildVitestArgs(['--run', 'sample.test.ts', '--reporter=dot'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      'sample.test.ts',
      '--reporter=dot',
      '--maxWorkers',
      '1',
      '--pool',
      'forks',
      '--configLoader',
      'runner',
    ])
  })

  it('splits large CI run file lists into serial batches', () => {
    const files = Array.from({ length: 7 }, (_, index) => `test-${index + 1}.test.ts`)

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
      }),
    ).toEqual([
      ['--run', ...files.slice(0, 3), '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', ...files.slice(3, 6), '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', files[6], '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('keeps broad corpus CI runs in one worker-bounded batch', () => {
    const files = [
      'scripts/__tests__/public-workbook-corpus.test.ts',
      'scripts/__tests__/public-workbook-corpus-cli.test.ts',
      'scripts/__tests__/public-workbook-corpus-evidence-refresh.test.ts',
      'scripts/__tests__/public-workbook-corpus-completion-audit.test.ts',
      'scripts/__tests__/public-workbook-corpus-completion-audit-roundtrip.test.ts',
      'scripts/__tests__/public-workbook-corpus-feature-witness-plan.test.ts',
      'scripts/__tests__/public-workbook-corpus-financial-plan.test.ts',
      'scripts/__tests__/public-workbook-corpus-links.test.ts',
      'scripts/__tests__/public-workbook-corpus-resource-limit-plan.test.ts',
      'scripts/__tests__/public-workbook-corpus-missing.test.ts',
      'scripts/__tests__/public-workbook-corpus-verify-checkpoint.test.ts',
      'scripts/__tests__/public-workbook-corpus-workbook.test.ts',
      'packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts',
      'packages/excel-import/src/__tests__/xlsx-table-sort-state-roundtrip.test.ts',
    ]

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
      }),
    ).toEqual([['--run', ...files, '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose']])
  })

  it('allows CI file chunk size overrides', () => {
    expect(
      buildVitestArgBatches(['--run', 'a.test.ts', 'b.test.ts', 'c.test.ts'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '2',
      }),
    ).toEqual([
      ['--run', 'a.test.ts', 'b.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
      ['--run', 'c.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('does not mix jsdom and node environment files in one CI batch', () => {
    expect(
      buildVitestArgBatches(
        [
          '--run',
          'apps/web/src/__tests__/workbook-editor-conflict.test.tsx',
          'apps/web/src/__tests__/worker-runtime-state.test.ts',
          'apps/web/src/__tests__/worker-workbook-app-model.test.ts',
        ],
        { BILIG_CI_PROFILE: 'fast' },
      ),
    ).toEqual([
      [
        '--run',
        'apps/web/src/__tests__/workbook-editor-conflict.test.tsx',
        '--maxWorkers',
        '1',
        '--pool',
        'forks',
        '--configLoader',
        'runner',
        '--reporter',
        'verbose',
      ],
      [
        '--run',
        'apps/web/src/__tests__/worker-runtime-state.test.ts',
        'apps/web/src/__tests__/worker-workbook-app-model.test.ts',
        '--maxWorkers',
        '1',
        '--pool',
        'forks',
        '--configLoader',
        'runner',
        '--reporter',
        'verbose',
      ],
    ])
  })

  it('ignores malformed CI file chunk size overrides instead of truncating them', () => {
    const files = ['a.test.ts', 'b.test.ts', 'c.test.ts', 'd.test.ts']

    expect(
      buildVitestArgBatches(['--run', ...files], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '2abc',
      }),
    ).toEqual([
      [
        '--run',
        'a.test.ts',
        'b.test.ts',
        'c.test.ts',
        '--maxWorkers',
        '1',
        '--pool',
        'forks',
        '--configLoader',
        'runner',
        '--reporter',
        'verbose',
      ],
      ['--run', 'd.test.ts', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner', '--reporter', 'verbose'],
    ])
  })

  it('does not split run arguments that include flags', () => {
    expect(
      buildVitestArgBatches(['--run', 'sample.test.ts', '--reporter=dot'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_FILE_CHUNK_SIZE: '1',
      }),
    ).toEqual([['--run', 'sample.test.ts', '--reporter=dot', '--maxWorkers', '1', '--pool', 'forks', '--configLoader', 'runner']])
  })

  it('adds a short CI-only cooldown between split batches', () => {
    expect(readVitestBatchCooldownMs({})).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast' })).toBe(1000)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '0' })).toBe(0)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '2500' })).toBe(2500)
    expect(readVitestBatchCooldownMs({ BILIG_CI_PROFILE: 'fast', BILIG_VITEST_BATCH_COOLDOWN_MS: '2500ms' })).toBe(1000)
  })

  it('resolves the Vitest binary from the workspace root', () => {
    expect(resolveVitestBin(repoRoot, 'darwin')).toBe(resolve(repoRoot, 'node_modules/.bin/vitest'))
    expect(resolveVitestBin(repoRoot, 'win32')).toBe(resolve(repoRoot, 'node_modules/.bin/vitest.cmd'))
  })

  it('classifies the public workbook corpus correctness lane as broad', () => {
    expect(
      isBroadCorpusVitestRun([
        '--run',
        'scripts/__tests__/public-workbook-corpus.test.ts',
        'scripts/__tests__/public-workbook-corpus-cli.test.ts',
        'scripts/__tests__/public-workbook-corpus-completion-audit.test.ts',
        'scripts/__tests__/public-workbook-corpus-links.test.ts',
      ]),
    ).toBe(true)
  })

  it('allows focused public workbook corpus Vitest checks', () => {
    expect(isBroadCorpusVitestRun(['--run', 'scripts/__tests__/public-workbook-corpus-links.test.ts'])).toBe(false)
  })

  it('classifies mixed public-corpus and xlsx import correctness as broad', () => {
    expect(
      isBroadCorpusVitestRun([
        '--run',
        'scripts/__tests__/public-workbook-corpus.test.ts',
        'packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts',
      ]),
    ).toBe(true)
  })

  it('runs package Vitest wrappers through tsx instead of bun', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const headlessPackageScripts = readPackageScripts(resolve(repoRoot, 'packages/headless/package.json'))
    const xlsxFormulaRecalcPackageScripts = readPackageScripts(resolve(repoRoot, 'packages/xlsx-formula-recalc/package.json'))
    const runVitestSource = readFileSync(resolve(repoRoot, 'scripts/run-vitest.ts'), 'utf8')

    expect(packageJson).toContain('"test": "tsx scripts/run-vitest.ts --run"')
    expect(packageJson).toContain('"coverage": "tsx scripts/run-vitest.ts --run --coverage')
    expect(packageJson).toContain('"test:watch": "tsx scripts/run-vitest.ts"')
    expect(packageJson).not.toContain('bun scripts/run-vitest.ts')
    expect(headlessPackageScripts['test:excel-oracle']).toContain('tsx ../../scripts/run-vitest.ts --root ../.. --run')
    expect(headlessPackageScripts['test:excel-oracle']).not.toContain('vitest run --root ../..')
    expect(xlsxFormulaRecalcPackageScripts['test:excel-oracle']).toContain('tsx ../../scripts/run-vitest.ts --root ../.. --run')
    expect(xlsxFormulaRecalcPackageScripts['test:excel-oracle']).not.toContain('vitest run --root ../..')
    expect(runVitestSource).toContain('process.stderr.write(`${error instanceof Error ? error.message : String(error)}\\n`)')
  })
})
