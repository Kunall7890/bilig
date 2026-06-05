import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

const repoRoot = fileURLToPath(new URL('../..', import.meta.url))

describe('public workbook corpus workbook SheetJS boundary', () => {
  it('does not load xlsx when importing the native corpus helper module', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
import('./scripts/public-workbook-corpus-workbook.ts')
  .then(() => {
    const loaded = Object.keys(requireForCache.cache).filter((path) => /(?:^|[\\\\/])xlsx(?:[\\\\/]|$)|[\\\\/]\\.pnpm[\\\\/]xlsx@/u.test(path))
    process.stdout.write(JSON.stringify({ loaded }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isLoadedModuleOutput(output)).toBe(true)
    if (!isLoadedModuleOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
  }, 15_000)
})

function isLoadedModuleOutput(value: unknown): value is { readonly loaded: readonly string[] } {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return Array.isArray(loaded) && loaded.every((entry) => typeof entry === 'string')
}

function isRecord(value: unknown): value is Record<PropertyKey, unknown> {
  return typeof value === 'object' && value !== null
}
