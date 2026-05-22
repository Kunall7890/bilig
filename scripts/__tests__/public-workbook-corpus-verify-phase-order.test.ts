import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { exportXlsx } from '../../packages/excel-import/src/index.js'
import type { WorkbookSnapshot } from '../../packages/protocol/src/types.js'
import {
  cloneWorkbookSnapshotForStructuralSmoke,
  roundTripsSupportedSemantics,
  verifyCachedWorkbookArtifact,
} from '../public-workbook-corpus-verify.ts'
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

  it('strips lazy package artifact closures before structural smoke cloning', () => {
    const baseSheet = buildSmallWorkbook().sheets[0]
    if (!baseSheet) {
      throw new Error('Expected small workbook sheet')
    }
    const workbookMetadata: NonNullable<WorkbookSnapshot['workbook']['metadata']> = Object.assign(
      { definedNames: [{ name: 'Answer', value: '=Sheet1!$A$1' }] },
      {
        slicerConnectionArtifacts: {
          parts: [
            {
              path: 'xl/slicerCaches/slicerCache1.xml',
              readBytes: () => new Uint8Array([1, 2, 3]),
            },
          ],
        },
      },
    )
    const sheetMetadata: NonNullable<WorkbookSnapshot['sheets'][number]['metadata']> = Object.assign(
      { merges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }] },
      {
        drawingArtifacts: {
          relationshipTarget: '../drawings/drawing1.xml',
          readBytes: () => new Uint8Array([4, 5, 6]),
        },
      },
    )
    const snapshot: WorkbookSnapshot = {
      ...buildSmallWorkbook(),
      workbook: {
        name: 'lazy artifact smoke',
        metadata: workbookMetadata,
      },
      sheets: [
        {
          ...baseSheet,
          metadata: sheetMetadata,
        },
      ],
    }

    const clone = cloneWorkbookSnapshotForStructuralSmoke(snapshot)

    expect(() => structuredClone(clone)).not.toThrow()
    expect(clone.workbook.metadata).toMatchObject({ definedNames: [{ name: 'Answer', value: '=Sheet1!$A$1' }] })
    expect(clone.workbook.metadata).not.toHaveProperty('slicerConnectionArtifacts')
    expect(clone.sheets[0]?.metadata).toMatchObject({ merges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }] })
    expect(clone.sheets[0]?.metadata).not.toHaveProperty('drawingArtifacts')
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
