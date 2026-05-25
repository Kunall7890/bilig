import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { acquireMacosExcelOracleLock, macosExcelOracleLockDir, withMacosExcelOracleLock } from '../macos-excel-oracle-lock.js'

describe('macOS Desktop Excel oracle interprocess lock', () => {
  it('serializes independent Excel oracle callers through one filesystem lock', () => {
    const lockRoot = mkdtempSync(join(tmpdir(), 'bilig-excel-oracle-lock-test-'))
    let now = 1_000
    const clock = () => now
    const sleep = (ms: number) => {
      now += ms
    }

    try {
      const release = acquireMacosExcelOracleLock({
        lockRoot,
        now: clock,
        processAlive: (pid) => pid === 12_345,
        processId: 12_345,
        sleep,
      })

      expect(existsSync(macosExcelOracleLockDir(lockRoot))).toBe(true)
      expect(() =>
        acquireMacosExcelOracleLock({
          lockRoot,
          now: clock,
          processAlive: (pid) => pid === 12_345,
          processId: 67_890,
          sleep,
        }),
      ).toThrow('Timed out waiting for macOS Desktop Excel oracle lock')

      release()

      const nextRelease = acquireMacosExcelOracleLock({
        lockRoot,
        now: clock,
        processAlive: (pid) => pid === 67_890,
        processId: 67_890,
        sleep,
      })
      nextRelease()

      expect(existsSync(macosExcelOracleLockDir(lockRoot))).toBe(false)
    } finally {
      rmSync(lockRoot, { recursive: true, force: true })
    }
  })

  it('recovers a lock left by a dead process', () => {
    const lockRoot = mkdtempSync(join(tmpdir(), 'bilig-excel-oracle-stale-lock-test-'))
    const lockDir = macosExcelOracleLockDir(lockRoot)

    try {
      mkdirSync(lockDir, { recursive: true })
      writeFileSync(
        join(lockDir, 'owner.json'),
        JSON.stringify({
          acquiredAt: new Date(1_000).toISOString(),
          acquiredAtMs: 1_000,
          pid: 12_345,
        }),
      )

      const release = acquireMacosExcelOracleLock({
        lockRoot,
        now: () => 2_000,
        processAlive: () => false,
        processId: 67_890,
        sleep: () => {},
      })
      const owner: unknown = JSON.parse(readFileSync(join(lockDir, 'owner.json'), 'utf8'))

      expect(owner).toMatchObject({ pid: 67_890 })
      release()
    } finally {
      rmSync(lockRoot, { recursive: true, force: true })
    }
  })

  it('releases the Excel oracle lock when the protected operation throws', () => {
    const lockRoot = mkdtempSync(join(tmpdir(), 'bilig-excel-oracle-release-lock-test-'))

    try {
      expect(() =>
        withMacosExcelOracleLock(
          {
            lockRoot,
            now: () => 1_000,
            processId: 12_345,
            sleep: () => {},
          },
          () => {
            throw new Error('oracle failure')
          },
        ),
      ).toThrow('oracle failure')

      expect(existsSync(macosExcelOracleLockDir(lockRoot))).toBe(false)
    } finally {
      rmSync(lockRoot, { recursive: true, force: true })
    }
  })
})
