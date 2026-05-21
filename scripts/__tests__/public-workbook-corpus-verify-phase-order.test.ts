import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { exportXlsx } from '../../packages/excel-import/src/index.js'
import type { WorkbookSnapshot } from '../../packages/protocol/src/types.js'
import { roundTripsSupportedSemantics, verifyCachedWorkbookArtifact } from '../public-workbook-corpus-verify.ts'
import { sha256HexSync } from '../public-workbook-corpus-workbook.ts'

describe('public workbook corpus verification phase order', () => {
  it('runs round-trip before structural smoke to avoid stacking peak worker memory', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-phase-order-'))
    const cachePath = 'phase-order.xlsx'
    const bytes = exportXlsx(buildSmallWorkbook())
    writeFileSync(join(cacheDir, cachePath), bytes)
    const phases: string[] = []

    const result = await verifyCachedWorkbookArtifact(
      {
        id: 'workbook-phase-order',
        sourceId: 'source-phase-order',
        sourceUrl: 'https://example.com/phase-order.xlsx',
        downloadUrl: 'https://example.com/phase-order.xlsx',
        fileName: cachePath,
        cachePath,
        sha256: sha256HexSync(bytes),
        byteSize: bytes.byteLength,
        workbookFingerprint: 'phase-order',
        fetchedAt: '2026-05-14T00:00:00.000Z',
        license: { spdxId: 'MIT', title: 'MIT', evidenceUrl: null },
      },
      cacheDir,
      true,
      1_000,
      {
        timeoutMs: 30_000,
        maxRssBytes: 1536 * 1024 * 1024,
        onPhase: (phase) => {
          phases.push(phase)
        },
      },
    )

    expect(result.validation.roundTripPassed).toBe(true)
    expect(result.validation.structuralSmokePassed).toBe(true)
    expect(phases.indexOf('round-trip')).toBeLessThan(phases.indexOf('structural-smoke'))
  })

  it('uses the round-tripped snapshot for structural smoke instead of retaining the original through reimport', async () => {
    const snapshot = buildSmallWorkbook()

    const result = await roundTripsSupportedSemantics(snapshot, { retainRoundTrippedSnapshot: true })

    expect(result.passed).toBe(true)
    expect(result.structuralSmokeSnapshot).toBeDefined()
    expect(result.structuralSmokeSnapshot).not.toBe(snapshot)
    expect(result.structuralSmokeSnapshot?.sheets[0]?.name).toBe('Sheet1')
    expect(result.structuralSmokeSnapshot?.sheets[0]?.cells).not.toBe(snapshot.sheets[0]?.cells)
  })

  it('does not keep a structural-smoke snapshot when the caller does not need one', async () => {
    const result = await roundTripsSupportedSemantics(buildSmallWorkbook())

    expect(result.passed).toBe(true)
    expect(result.structuralSmokeSnapshot).toBeUndefined()
  })
})

function buildSmallWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'phase order' },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        cells: [{ address: 'A1', value: 1 }],
      },
    ],
  }
}
