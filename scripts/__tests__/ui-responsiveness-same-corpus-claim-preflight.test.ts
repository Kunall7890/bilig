import { describe, expect, it } from 'vitest'

import {
  assertClaimGradeCaptureIncumbentsReady,
  parseCaptureArgs,
  preflightArgsForClaimGradeCapture,
} from '../capture-ui-responsiveness-same-corpus.ts'
import type { PreflightArgs } from '../ui-responsiveness-same-corpus-args.ts'
import type { PreflightEditableMutationProof, SameCorpusPreflight } from '../ui-responsiveness-same-corpus-preflight.ts'

describe('same-corpus claim-grade capture incumbent preflight', () => {
  it('maps claim-grade capture options to incumbent-only preflight options', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--serve-bilig-production',
      '--storage-state',
      'tmp/shared-state.json',
      '--google-sheets-storage-state',
      'tmp/google-state.json',
      '--microsoft-excel-web-storage-state',
      'tmp/microsoft-state.json',
      '--headed',
      '--samples',
      '5',
      '--ready-timeout-ms',
      '90000',
    ])

    expect(preflightArgsForClaimGradeCapture(args)).toMatchObject({
      corpusId: 'wide-mixed-250k',
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      googleSheetsStorageStatePath: expect.stringMatching(/tmp\/google-state\.json$/u),
      headless: false,
      microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      microsoftExcelWebStorageStatePath: expect.stringMatching(/tmp\/microsoft-state\.json$/u),
      outputPath: null,
      readyTimeoutMs: 90000,
      storageStatePath: expect.stringMatching(/tmp\/shared-state\.json$/u),
    })
  })

  it('keeps diagnostic incomplete-evidence captures out of claim-grade preflight', async () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--allow-incomplete-evidence',
    ])
    const calls: PreflightArgs[] = []

    await expect(
      assertClaimGradeCaptureIncumbentsReady(args, async (preflightArgs) => {
        calls.push(preflightArgs)
        throw new Error('diagnostic capture should not run claim-grade incumbent preflight')
      }),
    ).resolves.toBeUndefined()
    expect(preflightArgsForClaimGradeCapture(args)).toBeNull()
    expect(calls).toEqual([])
  })

  it('fails claim-grade capture on structured incumbent blockers before Bilig measurement', async () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--serve-bilig-production',
    ])
    const calls: PreflightArgs[] = []

    await expect(
      assertClaimGradeCaptureIncumbentsReady(args, async (preflightArgs) => {
        calls.push(preflightArgs)
        return blockedGoogleSheetsPreflight()
      }),
    ).rejects.toThrow('Google Sheets page is not authenticated')
    expect(calls).toHaveLength(1)
    expect(calls[0]).toMatchObject({
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      microsoftExcelWebUrl: null,
      outputPath: null,
    })
  })

  it('requires claim-grade incumbent preflight to prove editable write/readback/restore', async () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--serve-bilig-production',
    ])
    const readyPreflight = readyGoogleSheetsPreflight()

    await expect(
      assertClaimGradeCaptureIncumbentsReady(args, async () => ({
        ...readyPreflight,
        allCheckedProductsReady: true,
        products: [
          {
            ...readyPreflight.products[0],
            editableMutationProof: null,
          },
        ],
      })),
    ).rejects.toThrow('google-sheets is missing editable sentinel write/readback/restore proof')
  })

  it('allows claim-grade incumbent preflight only after editable restore proof is verified', async () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--serve-bilig-production',
    ])

    await expect(assertClaimGradeCaptureIncumbentsReady(args, async () => readyGoogleSheetsPreflight())).resolves.toBeUndefined()
  })
})

function blockedGoogleSheetsPreflight(): SameCorpusPreflight {
  return {
    mode: 'preflight',
    corpusCaseId: 'wide-mixed-250k',
    materializedCells: 250_000,
    requiredProductCount: 2,
    checkedProductCount: 1,
    readyProductCount: 0,
    blockedProductCount: 1,
    allCheckedProductsReady: false,
    products: [
      {
        product: 'google-sheets',
        source: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        finalUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        title: 'Google Sheets',
        status: 'blocked',
        blocker:
          'Cannot preflight same-corpus editable workloads on google-sheets: Google Sheets page is not authenticated; provide an authenticated storage state.',
        corpusVerification: null,
        editableMutationProof: null,
        limitations: ['storage state was not provided'],
      },
    ],
  }
}

function readyGoogleSheetsPreflight(): SameCorpusPreflight {
  return {
    mode: 'preflight',
    corpusCaseId: 'wide-mixed-250k',
    materializedCells: 250_000,
    requiredProductCount: 2,
    checkedProductCount: 1,
    readyProductCount: 1,
    blockedProductCount: 0,
    allCheckedProductsReady: true,
    products: [
      {
        product: 'google-sheets',
        source: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        finalUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        title: 'Google Sheets',
        status: 'ready',
        blocker: null,
        corpusVerification: {
          verified: true,
          method: 'google-sheets-xlsx-export',
          sheetName: 'Sheet1',
          materializedCells: 250_000,
          corpusFingerprint: {
            version: 'same-corpus-fingerprint-v1',
            corpusCaseId: 'wide-mixed-250k',
            workbookName: 'same-corpus-wide-mixed-250k',
            sheetCount: 1,
            materializedCells: 250_000,
            primaryViewport: {
              sheetName: 'Sheet1',
              rowStart: 1,
              rowEnd: 40,
              colStart: 1,
              colEnd: 12,
            },
            snapshotSha256: 'b'.repeat(64),
          },
          sourceWorkbookSha256: 'a'.repeat(64),
          checkedCells: [
            { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
            { address: 'B2', expected: '42', actual: '42' },
            { address: 'C3', expected: 'done', actual: 'done' },
          ],
        },
        editableMutationProof: verifiedEditableMutationProof(),
        limitations: [],
      },
    ],
  }
}

function verifiedEditableMutationProof(): PreflightEditableMutationProof {
  const before = {
    value: 'metric-1',
    formula: null,
    fillColor: null,
    visibleText: 'metric-1',
    source: 'visible-formula-bar',
  } as const
  const after = {
    ...before,
    value: 'same-corpus-preflight-google-sheets-c5',
    visibleText: 'same-corpus-preflight-google-sheets-c5',
  }
  return {
    product: 'google-sheets',
    captured: true,
    method: 'sentinel-write-readback-restore',
    sampleIndex: 0,
    sheetName: 'Sheet1',
    sheetId: 'google-sheet-id',
    targetRange: 'C5',
    intendedOperation: 'edit-visible-cell',
    intendedValue: 'same-corpus-preflight-google-sheets-c5',
    before,
    after,
    restored: before,
    authoritativeReadbackRevision: 'google-sheets-preflight-authoritative-readback-sha256:'.concat('c'.repeat(64)),
    visibleReadbackRevision: 'google-sheets-preflight-visible-readback-sha256:'.concat('d'.repeat(64)),
    undoRestoreStatus: 'verified',
    evidence: ['method=sentinel-write-readback-restore', 'targetRange=C5', 'writtenValueVerified=true', 'undoRestoreStatus=verified'],
  }
}
