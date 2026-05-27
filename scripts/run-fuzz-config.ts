import { parseStrictBooleanEnvFlag } from './strict-env.js'

export type FuzzMode = 'default' | 'main' | 'nightly' | 'replay'

export interface ReplayFixtureMetadata {
  readonly suite: string
  readonly kind: string | null
}

export function parseFuzzMode(value: string | undefined): FuzzMode {
  if (value === undefined || value === 'default') {
    return 'default'
  }
  if (value === 'main' || value === 'nightly' || value === 'replay') {
    return value
  }
  throw new Error(`Fuzz mode must be "default", "main", "nightly", or "replay", got ${value}`)
}

export function resolveVitestFuzzMaxWorkers(availableWorkers: number): number {
  if (!Number.isFinite(availableWorkers) || availableWorkers <= 0) {
    return 1
  }
  return Math.max(1, Math.min(2, Math.ceil(availableWorkers / 2)))
}

export function buildVitestFuzzCommand(files: readonly string[], availableWorkers: number): string[] {
  return ['pnpm', 'exec', 'vitest', 'run', ...files, '--maxWorkers', String(resolveVitestFuzzMaxWorkers(availableWorkers))]
}

export function fuzzSourceMayContainReplaySuite(source: string, suite: string): boolean {
  if (sourceContainsSuiteLiteral(source, suite)) {
    return true
  }

  const separatorIndex = suite.lastIndexOf('/')
  if (separatorIndex > 0 && sourceContainsDynamicSuitePrefix(source, suite.slice(0, separatorIndex))) {
    return true
  }

  return false
}

export function selectReplayVitestFuzzFiles(
  files: readonly string[],
  replayMetadata: ReplayFixtureMetadata | null,
  readSource: (filePath: string) => string,
): string[] {
  if (!replayMetadata) {
    return [...files]
  }

  const selected = files.filter((filePath) => fuzzSourceMayContainReplaySuite(readSource(filePath), replayMetadata.suite))
  if (selected.length === 0 && replayMetadata.kind !== 'browser') {
    throw new Error(`No Vitest fuzz file contains replay suite ${replayMetadata.suite}`)
  }
  return selected
}

export function resolveSkipBrowserFuzz(env: { BILIG_FUZZ_SKIP_BROWSER?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_FUZZ_SKIP_BROWSER, 'BILIG_FUZZ_SKIP_BROWSER', false)
}

function sourceContainsSuiteLiteral(source: string, suite: string): boolean {
  const escapedSuite = escapeRegExp(suite)
  return [`'${escapedSuite}'`, `"${escapedSuite}"`, `\`${escapedSuite}\``].some((quotedSuite) =>
    sourceContainsSuiteAssignment(source, quotedSuite),
  )
}

function sourceContainsDynamicSuitePrefix(source: string, prefix: string): boolean {
  const escapedPrefix = escapeRegExp(`${prefix}/`)
  return [`'${escapedPrefix}`, `"${escapedPrefix}`, `\`${escapedPrefix}`].some((quotedPrefix) =>
    sourceContainsSuiteAssignment(source, quotedPrefix),
  )
}

function sourceContainsSuiteAssignment(source: string, valuePattern: string): boolean {
  return new RegExp(`suite\\s*:\\s*${valuePattern}`, 'u').test(source)
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
