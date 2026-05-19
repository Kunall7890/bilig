import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import * as fc from 'fast-check'
import type { RunDetails } from 'fast-check'
import { afterEach, describe, expect, it } from 'vitest'
import {
  BYTE_FUZZ_DICTIONARY_PATH,
  BYTE_FUZZ_TARGETS_DIR,
  TEST_FUZZ_PACKAGE_ROOT,
  captureCounterexample,
  extractReplayPathForTest,
  resolveFuzzCaptureEnabled,
  runProperty,
  shouldRunFuzzSuite,
} from '../index.js'

const tempDirs: string[] = []

afterEach(() => {
  while (tempDirs.length > 0) {
    const tempDir = tempDirs.pop()
    if (tempDir) {
      rmSync(tempDir, { force: true, recursive: true })
    }
  }
})

function withTempCwd<T>(run: () => T): T {
  const previousCwd = process.cwd()
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-test-fuzz-'))
  tempDirs.push(tempDir)
  process.chdir(tempDir)
  try {
    return run()
  } finally {
    process.chdir(previousCwd)
  }
}

async function withEnv<T>(updates: Record<string, string | undefined>, run: () => Promise<T>): Promise<T> {
  const previous = new Map<string, string | undefined>()
  for (const key of Object.keys(updates)) {
    previous.set(key, process.env[key])
    const value = updates[key]
    if (value === undefined) {
      delete process.env[key]
    } else {
      process.env[key] = value
    }
  }
  try {
    return await run()
  } finally {
    for (const [key, value] of previous) {
      if (value === undefined) {
        delete process.env[key]
      } else {
        process.env[key] = value
      }
    }
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function createRunDetails(seed: number, counterexamplePath: string, counterexample: unknown[]): RunDetails<unknown[]> {
  return {
    failed: true,
    interrupted: false,
    numRuns: 1,
    numSkips: 0,
    numShrinks: 0,
    seed,
    counterexample,
    errorInstance: null,
    counterexamplePath,
    failures: [],
    executionSummary: [],
    verbose: 0,
    runConfiguration: {},
  }
}

function readReplayPath(artifactPath: string): string | null {
  const raw = JSON.parse(readFileSync(artifactPath, 'utf8')) as unknown
  if (!isRecord(raw)) {
    return null
  }
  return typeof raw['replayPath'] === 'string' ? raw['replayPath'] : null
}

describe('test-fuzz replay-path extraction', () => {
  it('owns byte-fuzz assets from the reusable test-fuzz package', () => {
    expect(TEST_FUZZ_PACKAGE_ROOT.endsWith('packages/test-fuzz')).toBe(true)
    expect(BYTE_FUZZ_TARGETS_DIR).toBe(join(TEST_FUZZ_PACKAGE_ROOT, 'byte-targets'))
    expect(BYTE_FUZZ_DICTIONARY_PATH).toBe(join(TEST_FUZZ_PACKAGE_ROOT, 'dictionaries', 'workbook-byte.dict'))
    expect(existsSync(BYTE_FUZZ_TARGETS_DIR)).toBe(true)
    expect(existsSync(BYTE_FUZZ_DICTIONARY_PATH)).toBe(true)
  })

  it('extracts replayPath from string counterexamples', () => {
    expect(extractReplayPathForTest(['cmd replayPath="0:1:2"'])).toBe('0:1:2')
  })

  it('extracts replayPath from nested command objects without string coercion', () => {
    const artifactPath = withTempCwd(() =>
      captureCounterexample({
        suite: 'grid/replay-path',
        kind: 'byte',
        details: createRunDetails(123, '0:0', [[{ kind: 'replay', replayPath: '/tmp/replay.json' }]]),
      }),
    )

    expect(readReplayPath(artifactPath)).toBe('/tmp/replay.json')
  })

  it('extracts replayPath from stringified counterexample fragments', () => {
    const artifactPath = withTempCwd(() =>
      captureCounterexample({
        suite: 'grid/replay-string',
        kind: 'byte',
        details: createRunDetails(456, '0:1', ['Command(replayPath="/tmp/from-string.json")']),
      }),
    )

    expect(readReplayPath(artifactPath)).toBe('/tmp/from-string.json')
  })

  it('writes replay commands through the unified fuzz entrypoint', () => {
    const artifactPath = withTempCwd(() =>
      captureCounterexample({
        suite: 'grid/replay-command',
        kind: 'byte',
        details: createRunDetails(789, '0:1', ['Command(replayPath="/tmp/from-string.json")']),
      }),
    )
    const raw: unknown = JSON.parse(readFileSync(artifactPath, 'utf8'))

    expect(isRecord(raw) ? raw['reproductionCommand'] : null).toBe(`pnpm test:fuzz -- replay ${artifactPath}`)
  })

  it('resolves fuzz capture flags strictly', () => {
    expect(resolveFuzzCaptureEnabled({})).toBe(false)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: '1' })).toBe(true)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: 'true' })).toBe(true)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: '0' })).toBe(false)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: 'false' })).toBe(false)
    expect(() => resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: 'yes' })).toThrow(
      'BILIG_FUZZ_CAPTURE must be "1", "true", "0", or "false" when set, got yes',
    )
  })

  it('fails malformed capture config before reporting a fuzz failure', async () => {
    await expect(
      withEnv({ BILIG_FUZZ_CAPTURE: 'yes' }, () =>
        runProperty({
          suite: 'capture/malformed-env',
          arbitrary: fc.constant('value'),
          predicate: () => {
            throw new Error('predicate failure')
          },
          parameters: { numRuns: 1 },
        }),
      ),
    ).rejects.toThrow('BILIG_FUZZ_CAPTURE must be "1", "true", "0", or "false" when set, got yes')
  })

  it('rejects retired fuzz profile names instead of silently downgrading them', async () => {
    await expect(
      withEnv({ BILIG_FUZZ_PROFILE: 'main' }, () =>
        runProperty({
          suite: 'profile/retired-name',
          arbitrary: fc.constant('value'),
          predicate: () => {},
          parameters: { numRuns: 1 },
        }),
      ),
    ).rejects.toThrow('BILIG_FUZZ_PROFILE must be "fuzz" or "replay" when set, got main')
  })

  it('fails selected suites that run zero generated cases', async () => {
    await expect(
      runProperty({
        suite: 'zero-runs/guard',
        arbitrary: fc.constant('value'),
        predicate: () => {},
        parameters: { numRuns: 0 },
      }),
    ).rejects.toThrow('Fuzz suite zero-runs/guard ran zero cases')
  })

  it('writes a replay hit marker only for the selected suite', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-test-fuzz-hit-'))
    tempDirs.push(tempDir)
    const replayFile = join(tempDir, 'replay.json')
    const hitFile = join(tempDir, 'hit.json')
    writeFileSync(replayFile, `${JSON.stringify({ suite: 'marker/suite', kind: 'property', seed: 123 })}\n`, 'utf8')
    await withEnv(
      {
        BILIG_FUZZ_REPLAY: replayFile,
        BILIG_FUZZ_REPLAY_HIT_FILE: hitFile,
      },
      async () => {
        expect(shouldRunFuzzSuite('other/suite', 'property')).toBe(false)
        expect(existsSync(hitFile)).toBe(false)
        expect(shouldRunFuzzSuite('marker/suite', 'property')).toBe(true)
        expect(existsSync(hitFile)).toBe(true)
      },
    )
  })
})
