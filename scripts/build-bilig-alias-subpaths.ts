#!/usr/bin/env bun

import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const headlessDistDir = join(rootDir, 'packages', 'headless', 'dist')
const biligDistDir = join(rootDir, 'packages', 'bilig', 'dist')

if (!existsSync(join(headlessDistDir, 'xlsx.js')) || !existsSync(join(headlessDistDir, 'xlsx.d.ts'))) {
  throw new Error('Build @bilig/headless before building the bilig XLSX subpath')
}

mkdirSync(biligDistDir, { recursive: true })

writeFileSync(join(biligDistDir, 'xlsx.js'), "export * from '@bilig/headless/xlsx'\n")
writeFileSync(join(biligDistDir, 'xlsx.d.ts'), "export * from '@bilig/headless/xlsx'\n")
