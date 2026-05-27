#!/usr/bin/env bun

import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { availableParallelism } from 'node:os'
import { join, resolve } from 'node:path'
import {
  buildVitestFuzzCommand,
  parseFuzzMode,
  resolveSkipBrowserFuzz,
  selectReplayVitestFuzzFiles,
  type FuzzMode,
  type ReplayFixtureMetadata,
} from './run-fuzz-config.js'

function runCommand(command: string[], extraEnv: Record<string, string>): void {
  const result = Bun.spawnSync(command, {
    stdin: 'inherit',
    stdout: 'inherit',
    stderr: 'inherit',
    env: {
      ...process.env,
      ...extraEnv,
    },
  })
  if (result.exitCode !== 0) {
    process.exit(result.exitCode ?? 1)
  }
}

const DEFAULT_FUZZ_PATTERNS = [
  /^packages\/formula\/src\/__tests__\/.+\.fuzz\.test\.ts$/,
  /^packages\/core\/src\/__tests__\/(engine-history|engine-structure|engine-replica|engine-snapshot|engine-import-export|snapshot-wire-parity|literal-loader-parity|formula-runtime-differential)\.fuzz\.test\.ts$/,
  /^packages\/(storage-server|excel-import|headless|binary-protocol|runtime-kernel|workbook)\/src\/__tests__\/.+\.fuzz\.test\.ts$/,
  /^packages\/grid\/src\/__tests__\/gridSelection\.fuzz\.test\.ts$/,
  /^packages\/renderer\/src\/__tests__\/commit-log\.fuzz\.test\.ts$/,
  /^packages\/wasm-kernel\/src\/__tests__\/kernel-bridge\.fuzz\.test\.ts$/,
  /^packages\/zero-sync\/src\/__tests__\/.+\.fuzz\.test\.ts$/,
  /^apps\/bilig\/src\/zero\/__tests__\/(projection|reconnect-replay|sync-relay|sync-relay-scheduled)\.fuzz\.test\.ts$/,
  /^apps\/web\/src\/__tests__\/(projected-viewport|runtime-sync|runtime-sync-scheduled|selection-command-parity|worker-workbook-app-model)\.fuzz\.test\.ts$/,
  /^packages\/worker-transport\/src\/__tests__\/.+\.fuzz\.test\.ts$/,
]

function listVitestFuzzFiles(): string[] {
  return ['packages', 'apps'].flatMap((root) => walkFuzzFiles(root)).toSorted((left, right) => left.localeCompare(right))
}

function selectVitestFuzzFiles(mode: FuzzMode, files: readonly string[]): string[] {
  if (mode !== 'default') {
    return [...files]
  }
  return files.filter((filePath) => DEFAULT_FUZZ_PATTERNS.some((pattern) => pattern.test(filePath)))
}

function shouldRunBrowserFuzz(mode: FuzzMode, replayKind: string | null, hasReplayFixture: boolean): boolean {
  if (hasReplayFixture) {
    return replayKind === 'browser'
  }
  return mode === 'main' || mode === 'nightly'
}

function parseReplayFixtureMetadata(filePath: string): ReplayFixtureMetadata {
  if (!existsSync(filePath)) {
    throw new Error(`Replay fixture does not exist: ${filePath}`)
  }
  const parsed = JSON.parse(readFileSync(filePath, 'utf8')) as unknown
  if (!isRecord(parsed) || typeof parsed['suite'] !== 'string') {
    throw new Error(`Replay fixture must contain a suite string: ${filePath}`)
  }
  return {
    suite: parsed['suite'],
    kind: typeof parsed['kind'] === 'string' ? parsed['kind'] : null,
  }
}

const args = process.argv.slice(2)
const mode = parseFuzzModeOrExit(args[0])
const replayFixture = mode === 'replay' ? args.slice(1).find((value) => value !== '--') : undefined

if (mode === 'replay' && !replayFixture) {
  console.error('Usage: pnpm test:fuzz:replay -- <fixture-path>')
  process.exit(1)
}

const resolvedReplayFixture = replayFixture ? resolve(replayFixture) : null
const replayMetadata = resolvedReplayFixture ? parseReplayFixtureMetadata(resolvedReplayFixture) : null
const env = {
  BILIG_FUZZ_PROFILE: mode,
  BILIG_FUZZ_CAPTURE: '1',
  ...(resolvedReplayFixture ? { BILIG_FUZZ_REPLAY: resolvedReplayFixture } : {}),
}
const skipBrowserFuzz = resolveSkipBrowserFuzzOrExit()
if (replayMetadata?.kind === 'browser' && skipBrowserFuzz) {
  console.error('Cannot replay a browser fuzz fixture while BILIG_FUZZ_SKIP_BROWSER is enabled')
  process.exit(1)
}

const vitestFuzzFiles = selectReplayVitestFuzzFiles(selectVitestFuzzFiles(mode, listVitestFuzzFiles()), replayMetadata, (filePath) =>
  readFileSync(filePath, 'utf8'),
)
if (vitestFuzzFiles.length > 0) {
  runCommand(buildVitestFuzzCommand(vitestFuzzFiles, availableParallelism()), env)
}

if (!skipBrowserFuzz && shouldRunBrowserFuzz(mode, replayMetadata?.kind ?? null, resolvedReplayFixture !== null)) {
  runCommand(['bun', 'scripts/run-browser-tests.ts', '--grep', '@fuzz-browser'], {
    ...env,
    BILIG_FUZZ_BROWSER: '1',
  })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function walkFuzzFiles(root: string): string[] {
  if (!existsSync(root)) {
    return []
  }

  const files: string[] = []
  for (const entry of readdirSync(root, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name.startsWith('.')) {
      continue
    }

    const relativePath = join(root, entry.name)
    if (entry.isDirectory()) {
      files.push(...walkFuzzFiles(relativePath))
      continue
    }

    if (entry.isFile() && relativePath.endsWith('.fuzz.test.ts')) {
      files.push(relativePath)
    }
  }

  return files
}

function parseFuzzModeOrExit(value: string | undefined): FuzzMode {
  try {
    return parseFuzzMode(value)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
}

function resolveSkipBrowserFuzzOrExit(): boolean {
  try {
    return resolveSkipBrowserFuzz(process.env)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
}
