import { parseStrictBooleanEnvFlag } from './strict-env.js'

export type FuzzMode = 'default' | 'main' | 'nightly' | 'replay'

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

export function resolveSkipBrowserFuzz(env: { BILIG_FUZZ_SKIP_BROWSER?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_FUZZ_SKIP_BROWSER, 'BILIG_FUZZ_SKIP_BROWSER', false)
}
