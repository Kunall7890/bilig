import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('run-ci', () => {
  it('serializes generated-source pnpm checks to avoid workspace-state races', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')

    expect(source).toContain('const generatedSourceChecks: readonly CiTask[] = [')
    expect(source).toContain('parallel pnpm invocations can race on .pnpm-workspace-state-v1.json')
    expect(source).toContain("await runSequential('generated-source checks', generatedSourceChecks)")
    expect(source).not.toContain("await runStage('generated-source checks'")
  })
})
