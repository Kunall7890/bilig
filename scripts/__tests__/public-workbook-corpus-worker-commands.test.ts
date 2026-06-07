import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, describe, expect, it, vi } from 'vitest'

import { writeFingerprintArtifactWorkerResult, writeFootprintWorkerResult } from '../public-workbook-corpus-worker-commands.ts'

describe('public workbook corpus worker commands', () => {
  const tempDirs: string[] = []

  afterEach(() => {
    vi.restoreAllMocks()
    for (const dir of tempDirs.splice(0)) {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('fails oversized fingerprint materialized fallback before importing workbook bytes', () => {
    const dir = mkdtempSync(join(tmpdir(), 'bilig-corpus-worker-fingerprint-'))
    tempDirs.push(dir)
    const filePath = join(dir, 'oversized.xlsx')
    writeFileSync(filePath, Buffer.alloc(1_000_001, 0))
    const stderrSpy = vi.spyOn(process.stderr, 'write').mockImplementation(() => true)
    const stdoutSpy = vi.spyOn(process.stdout, 'write').mockImplementation(() => true)
    const previousExitCode = process.exitCode
    process.exitCode = undefined

    try {
      writeFingerprintArtifactWorkerResult({
        filePath,
        fileName: 'oversized.xlsx',
        fingerprintMaxRssBytes: 1024 * 1024 * 1024,
      })

      expect(process.exitCode).toBe(1)
      expect(stdoutSpy).not.toHaveBeenCalled()
      expect(stderrSpy).toHaveBeenCalledWith(
        expect.stringContaining('Workbook fingerprinting worker materialized bytes fallback is small-workbook only'),
      )
    } finally {
      process.exitCode = previousExitCode
    }
  })

  it('fails oversized footprint materialized fallback before importing workbook bytes', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'bilig-corpus-worker-footprint-'))
    tempDirs.push(dir)
    const filePath = join(dir, 'oversized.xlsx')
    writeFileSync(filePath, Buffer.alloc(1_000_001, 0))

    await expect(
      writeFootprintWorkerResult({
        filePath,
        fileName: 'oversized.xlsx',
        verifyMaxRssBytes: 1024 * 1024 * 1024,
      }),
    ).rejects.toThrow(/Workbook footprint worker materialized bytes fallback is small-workbook only/u)
  })
})
