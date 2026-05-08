#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

const DEFAULT_CI_FILE_CHUNK_SIZE = 12

export function buildVitestArgs(args: readonly string[], env: NodeJS.ProcessEnv = process.env): string[] {
  if (!env['BILIG_CI_PROFILE'] || hasArg(args, '--maxWorkers')) {
    return [...args]
  }
  return [...args, '--maxWorkers', env['BILIG_VITEST_MAX_WORKERS'] ?? '1']
}

export function buildVitestArgBatches(args: readonly string[], env: NodeJS.ProcessEnv = process.env): string[][] {
  return splitVitestRunArgsForCi(args, env).map((batchArgs) => buildVitestArgs(batchArgs, env))
}

function splitVitestRunArgsForCi(args: readonly string[], env: NodeJS.ProcessEnv): string[][] {
  if (!env['BILIG_CI_PROFILE']) {
    return [[...args]]
  }

  const runIndex = args.indexOf('--run')
  if (runIndex < 0) {
    return [[...args]]
  }

  const prefixArgs = args.slice(0, runIndex + 1)
  const runArgs = args.slice(runIndex + 1)
  if (runArgs.length === 0 || runArgs.some((arg) => arg.startsWith('-'))) {
    return [[...args]]
  }

  const chunkSize = readPositiveInt(env['BILIG_VITEST_FILE_CHUNK_SIZE']) ?? DEFAULT_CI_FILE_CHUNK_SIZE
  if (runArgs.length <= chunkSize) {
    return [[...args]]
  }

  const batches: string[][] = []
  for (let start = 0; start < runArgs.length; start += chunkSize) {
    batches.push([...prefixArgs, ...runArgs.slice(start, start + chunkSize)])
  }
  return batches
}

function hasArg(args: readonly string[], flag: string): boolean {
  return args.some((arg) => arg === flag || arg.startsWith(`${flag}=`))
}

function readPositiveInt(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }

  const parsed = Number.parseInt(value, 10)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : undefined
}

function main(): never {
  ensureWasmKernelArtifact()

  const vitestBin = process.platform === 'win32' ? 'node_modules\\.bin\\vitest.cmd' : 'node_modules/.bin/vitest'
  for (const args of buildVitestArgBatches(process.argv.slice(2))) {
    const result = spawnSync(vitestBin, args, {
      cwd: process.cwd(),
      env: process.env,
      stdio: 'inherit',
    })

    if (result.error) {
      throw result.error
    }

    if (result.signal) {
      process.stderr.write(`vitest terminated by signal ${result.signal}\n`)
    }

    if (result.status !== 0) {
      process.exit(result.status ?? 1)
    }
  }

  process.exit(0)
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
