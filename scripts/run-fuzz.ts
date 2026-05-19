#!/usr/bin/env bun

import { existsSync, mkdirSync, readdirSync, rmSync } from 'node:fs'
import { availableParallelism } from 'node:os'
import { join, relative, resolve } from 'node:path'
import { BYTE_FUZZ_DICTIONARY_PATH, BYTE_FUZZ_TARGETS_DIR } from '@bilig/test-fuzz'
import { buildVitestFuzzCommand, parseFuzzMode, type FuzzMode } from './run-fuzz-config.js'

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

function listVitestFuzzFiles(): string[] {
  return ['packages', 'apps'].flatMap((root) => walkFuzzFiles(root)).toSorted((left, right) => left.localeCompare(right))
}

function assertReplayFixtureExists(filePath: string): void {
  if (!existsSync(filePath)) {
    throw new Error(`Replay fixture does not exist: ${filePath}`)
  }
}

const args = process.argv.slice(2)
const mode = parseFuzzModeOrExit(args[0])
const replayFixture = mode === 'replay' ? args.slice(1).find((value) => value !== '--') : undefined

if (mode === 'replay' && !replayFixture) {
  console.error('Usage: pnpm test:fuzz -- replay <fixture-path>')
  process.exit(1)
}

const resolvedReplayFixture = replayFixture ? resolve(replayFixture) : null
if (resolvedReplayFixture) {
  assertReplayFixtureExists(resolvedReplayFixture)
}
const env = {
  BILIG_FUZZ_PROFILE: mode,
  BILIG_FUZZ_CAPTURE: '1',
  ...(resolvedReplayFixture ? { BILIG_FUZZ_REPLAY: resolvedReplayFixture } : {}),
}

const vitestFuzzFiles = listVitestFuzzFiles()
if (vitestFuzzFiles.length === 0) {
  console.error('No Vitest fuzz files were discovered.')
  process.exit(1)
}

const replayHitFile = resolvedReplayFixture
  ? resolve('artifacts/fuzz/replay-hit', `${String(process.pid)}-${Date.now().toString(36)}.json`)
  : null
if (replayHitFile) {
  rmSync(replayHitFile, { force: true })
}

runCommand(buildVitestFuzzCommand(vitestFuzzFiles, availableParallelism()), {
  ...env,
  ...(replayHitFile ? { BILIG_FUZZ_REPLAY_HIT_FILE: replayHitFile } : {}),
})

if (replayHitFile && !existsSync(replayHitFile)) {
  console.error(`Replay fixture did not match any executed fuzz suite: ${resolvedReplayFixture ?? ''}`)
  process.exit(1)
}

if (mode === 'fuzz') {
  const byteTargets = listByteFuzzTargets()
  if (byteTargets.length > 0) {
    mkdirSync(resolve('artifacts/fuzz/jazzer-corpus'), { recursive: true })
    runCommand(['pnpm', '--filter', '@bilig/app^...', 'run', 'build'], {})
    for (const target of byteTargets) {
      const targetName = relative(BYTE_FUZZ_TARGETS_DIR, target)
      const corpusDir = resolve('artifacts/fuzz/jazzer-corpus', targetName.replace(/[^A-Za-z0-9_.-]/gu, '_'))
      mkdirSync(corpusDir, { recursive: true })
      runCommand(
        [
          'pnpm',
          'exec',
          'jazzer',
          target,
          corpusDir,
          '-i',
          'packages/',
          '-i',
          'apps/',
          '--',
          `-dict=${BYTE_FUZZ_DICTIONARY_PATH}`,
          '-runs=2048',
          '-max_total_time=10',
        ],
        {},
      )
    }
  }
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

    if (entry.isFile() && (relativePath.endsWith('.fuzz.test.ts') || relativePath.endsWith('.fuzz.test.tsx'))) {
      files.push(relativePath)
    }
  }

  return files
}

function listByteFuzzTargets(): string[] {
  return walkByteFuzzTargets(BYTE_FUZZ_TARGETS_DIR).toSorted((left, right) => left.localeCompare(right))
}

function walkByteFuzzTargets(root: string): string[] {
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
      files.push(...walkByteFuzzTargets(relativePath))
      continue
    }

    if (entry.isFile() && relativePath.endsWith('.mjs')) {
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
