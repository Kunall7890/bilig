import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import {
  assertLocalCiResourceGuardAllowsRun,
  localCiResourceGuardOverrideEnv,
  readLocalCiResourceGuardStatus,
} from '../ci-local-resource-guard.ts'

describe('local CI resource guard', () => {
  it('allows broad CI when no local coordination guard exists', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ci-guard-empty-'))

    expect(readLocalCiResourceGuardStatus(rootDir).activeMarkerPaths).toEqual([])
    expect(() => assertLocalCiResourceGuardAllowsRun(rootDir, {})).not.toThrow()
  })

  it('blocks broad CI while an active stop marker exists', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ci-guard-active-'))
    const coordinationDir = join(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(
      join(coordinationDir, '20260507T074946Z-codex-stop-interactive-corpus-runs.md'),
      '# Stop interactive corpus runs\n\nStatus: active on 2026-05-07T07:49:46Z.\n',
    )

    expect(readLocalCiResourceGuardStatus(rootDir).activeMarkerPaths).toEqual([
      '.agent-coordination/20260507T074946Z-codex-stop-interactive-corpus-runs.md',
    ])
    expect(() => assertLocalCiResourceGuardAllowsRun(rootDir, {})).toThrow(/Refusing to start broad CI/u)
  })

  it('uses the caller label in the refusal message', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ci-guard-label-'))
    const coordinationDir = join(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(
      join(coordinationDir, '20260507T074946Z-codex-stop-interactive-corpus-runs.md'),
      '# Stop interactive corpus runs\n\nStatus: active on 2026-05-07T07:49:46Z.\n',
    )

    expect(() => assertLocalCiResourceGuardAllowsRun(rootDir, {}, { runLabel: 'pre-push lint' })).toThrow(
      /Refusing to start pre-push lint/u,
    )
  })

  it('blocks broad CI while the memory guard marker is active', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ci-guard-memory-'))
    const coordinationDir = join(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(join(coordinationDir, '20260508T072657Z-codex-memory-guard.md'), '# Memory guard active\n')

    expect(readLocalCiResourceGuardStatus(rootDir).activeMarkerPaths).toEqual([
      '.agent-coordination/20260508T072657Z-codex-memory-guard.md',
    ])
    expect(() => assertLocalCiResourceGuardAllowsRun(rootDir, {})).toThrow(localCiResourceGuardOverrideEnv)
  })

  it('allows broad CI with an explicit local guard override', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ci-guard-override-'))
    const coordinationDir = join(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(
      join(coordinationDir, '20260508T092619Z-codex-memory-pressure-stop.md'),
      '# Memory pressure stop\n\nStatus: active on 2026-05-08T09:26:19Z.\n',
    )

    expect(() => assertLocalCiResourceGuardAllowsRun(rootDir, { [localCiResourceGuardOverrideEnv]: '1' })).not.toThrow()
  })

  it('rejects malformed local guard override values instead of ignoring them', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ci-guard-bad-override-'))

    expect(() => assertLocalCiResourceGuardAllowsRun(rootDir, { [localCiResourceGuardOverrideEnv]: 'true' })).toThrow(
      `${localCiResourceGuardOverrideEnv} must be "1" when set, got true`,
    )
  })

  it('ignores inactive guard-like files', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-ci-guard-inactive-'))
    const coordinationDir = join(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(
      join(coordinationDir, '20260507T074946Z-codex-stop-interactive-corpus-runs.md'),
      '# Stop interactive corpus runs\n\nStatus: resolved.\n',
    )

    expect(readLocalCiResourceGuardStatus(rootDir).activeMarkerPaths).toEqual([])
    expect(() => assertLocalCiResourceGuardAllowsRun(rootDir, {})).not.toThrow()
  })
})
