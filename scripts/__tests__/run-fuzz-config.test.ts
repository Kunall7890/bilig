import { describe, expect, it } from 'vitest'
import {
  buildVitestFuzzCommand,
  fuzzSourceMayContainReplaySuite,
  parseFuzzMode,
  resolveSkipBrowserFuzz,
  resolveVitestFuzzMaxWorkers,
  selectReplayVitestFuzzFiles,
} from '../run-fuzz-config.js'

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

  it('resolves the browser fuzz skip flag strictly', () => {
    expect(resolveSkipBrowserFuzz({})).toBe(false)
    expect(resolveSkipBrowserFuzz({ BILIG_FUZZ_SKIP_BROWSER: '1' })).toBe(true)
    expect(resolveSkipBrowserFuzz({ BILIG_FUZZ_SKIP_BROWSER: 'true' })).toBe(true)
    expect(resolveSkipBrowserFuzz({ BILIG_FUZZ_SKIP_BROWSER: '0' })).toBe(false)
    expect(resolveSkipBrowserFuzz({ BILIG_FUZZ_SKIP_BROWSER: 'false' })).toBe(false)
    expect(() => resolveSkipBrowserFuzz({ BILIG_FUZZ_SKIP_BROWSER: 'yes' })).toThrow(
      'BILIG_FUZZ_SKIP_BROWSER must be "1", "true", "0", or "false" when set, got yes',
    )
  })

  it('matches replay suites against exact suite literals in fuzz sources', () => {
    const source = "await runProperty({ suite: 'excel-import/xlsx/generated-workbooks', arbitrary, predicate })"

    expect(fuzzSourceMayContainReplaySuite(source, 'excel-import/xlsx/generated-workbooks')).toBe(true)
    expect(fuzzSourceMayContainReplaySuite(source, 'formula/parse/invalid-input')).toBe(false)
  })

  it('matches replay suites against dynamic suite prefixes in fuzz sources', () => {
    const source = 'await runProperty({ suite: `core/history/${seedName}`, arbitrary, predicate })'

    expect(fuzzSourceMayContainReplaySuite(source, 'core/history/formula-graph')).toBe(true)
    expect(fuzzSourceMayContainReplaySuite(source, 'core/snapshot/formula-graph')).toBe(false)
  })

  it('selects only vitest fuzz files that can contain a replay suite', () => {
    const sources = new Map([
      ['formula.fuzz.test.ts', "await runProperty({ suite: 'formula/parse/valid-input', arbitrary, predicate })"],
      ['history.fuzz.test.ts', 'await runProperty({ suite: `core/history/${seedName}`, arbitrary, predicate })'],
    ])

    expect(
      selectReplayVitestFuzzFiles(
        [...sources.keys()],
        { suite: 'core/history/formula-graph', kind: 'model' },
        (filePath) => sources.get(filePath) ?? '',
      ),
    ).toEqual(['history.fuzz.test.ts'])
  })

  it('allows browser replay fixtures to skip vitest file selection', () => {
    expect(selectReplayVitestFuzzFiles(['formula.fuzz.test.ts'], { suite: 'browser/selection', kind: 'browser' }, () => '')).toEqual([])
  })

  it('fails non-browser replay fixtures when no vitest fuzz file can contain the suite', () => {
    expect(() => selectReplayVitestFuzzFiles(['formula.fuzz.test.ts'], { suite: 'core/history/missing', kind: 'model' }, () => '')).toThrow(
      'No Vitest fuzz file contains replay suite core/history/missing',
    )
  })
})
