import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import { existsSync, mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import {
  captureSameCorpusCommittedStatePhaseProof,
  type SameCorpusCommittedStatePage,
} from '../ui-responsiveness-same-corpus-committed-state-proof.ts'
import type { SameCorpusMutationTargetSelection } from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { SameCorpusMutationTargetReadback } from '../ui-responsiveness-same-corpus-semantic-proof.ts'

describe('same-corpus committed-state proof capture', () => {
  it('waits for Google Sheets XLSX export to match the expected target readback before accepting proof', async () => {
    const staleBytes = xlsxBytesForTargetValue('WideGrid', 'C5', 'stale-browser-only-value')
    const committedBytes = xlsxBytesForTargetValue('WideGrid', 'C5', 'same-corpus-edit-1')
    const page = mockGoogleSheetsExportPage([staleBytes, committedBytes])

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: sameCorpusGoogleReadback('same-corpus-edit-1'),
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'edit-visible-cell',
    })

    expect(page.requestCount()).toBe(2)
    expect(proof).toMatchObject({
      product: 'google-sheets',
      phase: 'after',
      targetRange: 'C5',
      readback: {
        value: 'same-corpus-edit-1',
        source: 'google-sheets-xlsx-export',
      },
    })
  })

  it('writes a per-phase committed-state JSON artifact for accepted Google Sheets export proof', async () => {
    const committedBytes = xlsxBytesForTargetValue('WideGrid', 'C5', 'same-corpus-edit-1')
    const page = mockGoogleSheetsExportPage([committedBytes])
    const artifactPath = join(mkdtempSync(join(tmpdir(), 'bilig-committed-proof-')), 'after.json')

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      artifactPath,
      expectedReadback: sameCorpusGoogleReadback('same-corpus-edit-1'),
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'edit-visible-cell',
    })

    expect(proof?.artifactPath).toBeTruthy()
    expect(proof?.artifactSha256).toMatch(/^[a-f0-9]{64}$/u)
    expect(existsSync(artifactPath)).toBe(true)
    expect(readFileSync(artifactPath, 'utf8')).toContain('"source": "google-sheets-xlsx-export"')
  })

  it('keeps cached formula results in XLSX readback instead of formula text only', async () => {
    const page = mockGoogleSheetsExportPage([xlsxBytesForFormulaResult('WideGrid', 'D5', '1+1', 2)])

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: { ...sameCorpusGoogleReadback('2'), formula: '=1+1' },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: {
        ...sameCorpusTargetSelection(),
        endAddress: 'D5',
        startAddress: 'D5',
        targetRange: 'D5',
      },
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'formula-edit',
    })

    expect(proof?.readback).toMatchObject({
      formula: '=1+1',
      value: '2',
      visibleText: '2',
      source: 'google-sheets-xlsx-export',
    })
  })

  it('fails when Google Sheets XLSX export never proves the expected committed target readback', async () => {
    const page = mockGoogleSheetsExportPage([xlsxBytesForTargetValue('WideGrid', 'C5', 'stale-browser-only-value')])

    await expect(
      captureSameCorpusCommittedStatePhaseProof({
        expectedReadback: sameCorpusGoogleReadback('same-corpus-edit-1'),
        page: page.page,
        phase: 'after',
        product: 'google-sheets',
        sampleIndex: 0,
        target: sameCorpusTargetSelection(),
        timeoutMs: 1,
        pollIntervalMs: 1,
        workload: 'edit-visible-cell',
      }),
    ).rejects.toThrow(/did not match expected after target readback/u)
  })
})

function sameCorpusGoogleReadback(value: string): SameCorpusMutationTargetReadback {
  return {
    value,
    formula: null,
    fillColor: null,
    visibleText: value,
    source: 'visible-formula-bar',
  }
}

function sameCorpusTargetSelection(): SameCorpusMutationTargetSelection {
  return {
    endAddress: 'C5',
    sheetId: '0',
    sheetName: 'WideGrid',
    startAddress: 'C5',
    targetRange: 'C5',
  }
}

function mockGoogleSheetsExportPage(responses: readonly Uint8Array[]): {
  readonly page: SameCorpusCommittedStatePage
  readonly requestCount: () => number
} {
  let requestCount = 0
  const page: SameCorpusCommittedStatePage = {
    context: () => ({
      request: {
        get: async () => {
          const body = responses[Math.min(requestCount, responses.length - 1)] ?? responses[0]
          requestCount += 1
          return {
            body: async () => Buffer.from(body),
            ok: () => true,
            status: () => 200,
            text: async () => '',
          }
        },
      },
    }),
    url: () => 'https://docs.google.com/spreadsheets/d/test-spreadsheet/edit#gid=0',
    waitForTimeout: async () => {},
  }
  return { page, requestCount: () => requestCount }
}

function xlsxBytesForTargetValue(sheetName: string, address: string, value: string): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet([[]])
  worksheet[address] = { t: 's', v: value }
  worksheet['!ref'] = `A1:${address}`
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function xlsxBytesForFormulaResult(sheetName: string, address: string, formula: string, value: number): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet([[]])
  worksheet[address] = { f: formula, t: 'n', v: value }
  worksheet['!ref'] = `A1:${address}`
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}
