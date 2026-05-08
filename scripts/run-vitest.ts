#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

ensureWasmKernelArtifact()

const vitestBin = process.platform === 'win32' ? 'node_modules\\.bin\\vitest.cmd' : 'node_modules/.bin/vitest'
const result = spawnSync(vitestBin, process.argv.slice(2), {
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

process.exit(result.status ?? 1)
