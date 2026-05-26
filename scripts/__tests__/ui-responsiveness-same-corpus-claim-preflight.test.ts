import { describe, expect, it } from 'vitest'

import {
  assertClaimGradeCaptureIncumbentsReady,
  parseCaptureArgs,
  preflightArgsForClaimGradeCapture,
} from '../capture-ui-responsiveness-same-corpus.ts'
import type { PreflightArgs } from '../ui-responsiveness-same-corpus-args.ts'
import type { SameCorpusPreflight } from '../ui-responsiveness-same-corpus-preflight.ts'

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
        limitations: ['storage state was not provided'],
      },
    ],
  }
}
