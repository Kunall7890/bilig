import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
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
    expect(new Set(page.requestUrls()).size).toBe(2)
    expect(page.requestUrls().every((url) => url.includes('biligProofNonce='))).toBe(true)
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

  it('uses raw OOXML styles when Google Sheets fill color is not exposed by SheetJS style readback', async () => {
    const staleBytes = xlsxBytesForTargetValue('WideGrid', 'C5', 'segment-5')
    const committedBytes = xlsxBytesForTargetFill('WideGrid', 'C5', 'segment-5', '#c9daf8')
    const page = mockGoogleSheetsExportPage([staleBytes, committedBytes])

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#c9daf8',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(page.requestCount()).toBe(2)
    expect(proof?.readback).toMatchObject({
      value: 'segment-5',
      fillColor: '#c9daf8',
      source: 'google-sheets-xlsx-export',
    })
  })

  it('resolves Google Sheets theme fill colors from OOXML committed-state exports', async () => {
    const committedBytes = xlsxBytesForTargetThemeFill('WideGrid', 'C5', 'segment-5', '#c9daf8')
    const page = mockGoogleSheetsExportPage([committedBytes])

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#c9daf8',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(page.requestCount()).toBe(1)
    expect(proof?.readback).toMatchObject({
      value: 'segment-5',
      fillColor: '#c9daf8',
      source: 'google-sheets-xlsx-export',
    })
  })

  it('resolves Google Sheets indexed fill colors from OOXML committed-state exports', async () => {
    const committedBytes = xlsxBytesForTargetIndexedFill('WideGrid', 'C5', 'segment-5', 3)
    const page = mockGoogleSheetsExportPage([committedBytes])

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#00ff00',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(page.requestCount()).toBe(1)
    expect(proof?.readback).toMatchObject({
      value: 'segment-5',
      fillColor: '#00ff00',
      source: 'google-sheets-xlsx-export',
    })
  })

  it('resolves target fills inherited from OOXML row styles', async () => {
    const committedBytes = xlsxBytesForTargetRowFill('WideGrid', 'C5', 'segment-5', '#00ff00')
    const page = mockGoogleSheetsExportPage([committedBytes])

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#00ff00',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(proof?.readback).toMatchObject({
      value: 'segment-5',
      fillColor: '#00ff00',
      source: 'google-sheets-xlsx-export',
    })
  })

  it('resolves target fills inherited from OOXML column styles', async () => {
    const committedBytes = xlsxBytesForTargetColumnFill('WideGrid', 'C5', 'segment-5', '#a4c2f4')
    const page = mockGoogleSheetsExportPage([committedBytes])

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#a4c2f4',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 1_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(proof?.readback).toMatchObject({
      value: 'segment-5',
      fillColor: '#a4c2f4',
      source: 'google-sheets-xlsx-export',
    })
  })

  it('waits for Google Sheets to finish saving before reading the XLSX export when browser state is available', async () => {
    const committedBytes = xlsxBytesForTargetFill('WideGrid', 'C5', 'segment-5', '#c9daf8')
    let saveIdleWaitCount = 0
    const page = mockGoogleSheetsExportPage([committedBytes], {
      waitForFunction: async (_pageFunction, _arg, options) => {
        saveIdleWaitCount += 1
        expect(options?.timeout).toBeGreaterThan(0)
      },
    })

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#c9daf8',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 5_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(saveIdleWaitCount).toBe(1)
    expect(page.requestCount()).toBe(1)
    expect(proof?.readback.fillColor).toBe('#c9daf8')
  })

  it('uses XLSX export proof when Google Sheets save-idle text cannot settle', async () => {
    const committedBytes = xlsxBytesForTargetFill('WideGrid', 'C5', 'segment-5', '#c9daf8')
    const page = mockGoogleSheetsExportPage([committedBytes], {
      waitForFunction: async () => {
        throw new Error('Saving still visible')
      },
    })

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#c9daf8',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 5_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(page.requestCount()).toBe(1)
    expect(proof?.readback).toMatchObject({
      fillColor: '#c9daf8',
      source: 'google-sheets-xlsx-export',
      value: 'segment-5',
    })
  })

  it('uses page-context fetch when a live Google Sheets page is available', async () => {
    const committedBytes = xlsxBytesForTargetFill('WideGrid', 'C5', 'segment-5', '#c9daf8')
    let pageFetchCount = 0
    const page = mockGoogleSheetsExportPage([committedBytes], {
      requestError: new TypeError('"/spreadsheets/d/test-spreadsheet/export?format=xlsx" cannot be parsed as a URL.'),
      evaluate: async (_pageFunction, _arg) => {
        pageFetchCount += 1
        return {
          bodyText: '',
          bytes: Array.from(committedBytes),
          ok: true,
          status: 200,
        }
      },
    })

    const proof = await captureSameCorpusCommittedStatePhaseProof({
      expectedReadback: {
        ...sameCorpusGoogleReadback('segment-5'),
        fillColor: '#c9daf8',
      },
      page: page.page,
      phase: 'after',
      product: 'google-sheets',
      sampleIndex: 0,
      target: sameCorpusTargetSelection(),
      timeoutMs: 5_000,
      pollIntervalMs: 0,
      workload: 'fill-format-change',
    })

    expect(page.requestCount()).toBe(0)
    expect(pageFetchCount).toBe(1)
    expect(proof?.readback).toMatchObject({
      fillColor: '#c9daf8',
      source: 'google-sheets-xlsx-export',
      value: 'segment-5',
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

  it('carries the last committed export readback in mismatch diagnostics', async () => {
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
    ).rejects.toMatchObject({
      name: 'SameCorpusCommittedStateMismatchError',
      diagnostic: {
        expectedReadback: {
          value: 'same-corpus-edit-1',
        },
        lastReadback: {
          value: 'stale-browser-only-value',
          source: 'google-sheets-xlsx-export',
        },
        targetRange: 'C5',
      },
    })
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

function mockGoogleSheetsExportPage(
  responses: readonly Uint8Array[],
  options: {
    readonly evaluate?: NonNullable<SameCorpusCommittedStatePage['evaluate']>
    readonly requestError?: Error
    readonly waitForFunction?: NonNullable<SameCorpusCommittedStatePage['waitForFunction']>
  } = {},
): {
  readonly page: SameCorpusCommittedStatePage
  readonly requestCount: () => number
  readonly requestUrls: () => readonly string[]
} {
  let requestCount = 0
  const requestUrls: string[] = []
  const page: SameCorpusCommittedStatePage = {
    context: () => ({
      request: {
        get: async (url) => {
          requestUrls.push(url)
          if (options.requestError) {
            requestCount += 1
            throw options.requestError
          }
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
    ...(options.evaluate ? { evaluate: options.evaluate } : {}),
    ...(options.waitForFunction ? { waitForFunction: options.waitForFunction } : {}),
    waitForTimeout: async () => {},
  }
  return { page, requestCount: () => requestCount, requestUrls: () => [...requestUrls] }
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

function xlsxBytesForTargetFill(sheetName: string, address: string, value: string, fillColor: string): Uint8Array {
  const archive = unzipSync(xlsxBytesForTargetValue(sheetName, address, value))
  const stylesXml = strFromU8(archive['xl/styles.xml'] ?? new Uint8Array())
  const sheetXml = strFromU8(archive['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  archive['xl/styles.xml'] = strToU8(addTargetFillStyle(stylesXml, fillColor))
  archive['xl/worksheets/sheet1.xml'] = strToU8(addTargetCellStyle(sheetXml, address, 1))
  return zipSync(archive)
}

function xlsxBytesForTargetThemeFill(sheetName: string, address: string, value: string, fillColor: string): Uint8Array {
  const archive = unzipSync(xlsxBytesForTargetValue(sheetName, address, value))
  const stylesXml = strFromU8(archive['xl/styles.xml'] ?? new Uint8Array())
  const sheetXml = strFromU8(archive['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  const themeXml = strFromU8(archive['xl/theme/theme1.xml'] ?? new Uint8Array())
  archive['xl/styles.xml'] = strToU8(addTargetThemeFillStyle(stylesXml))
  archive['xl/worksheets/sheet1.xml'] = strToU8(addTargetCellStyle(sheetXml, address, 1))
  archive['xl/theme/theme1.xml'] = strToU8(setThemeAccent1(themeXml, fillColor))
  return zipSync(archive)
}

function xlsxBytesForTargetIndexedFill(sheetName: string, address: string, value: string, colorIndex: number): Uint8Array {
  const archive = unzipSync(xlsxBytesForTargetValue(sheetName, address, value))
  const stylesXml = strFromU8(archive['xl/styles.xml'] ?? new Uint8Array())
  const sheetXml = strFromU8(archive['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  archive['xl/styles.xml'] = strToU8(addTargetIndexedFillStyle(stylesXml, colorIndex))
  archive['xl/worksheets/sheet1.xml'] = strToU8(addTargetCellStyle(sheetXml, address, 1))
  return zipSync(archive)
}

function xlsxBytesForTargetRowFill(sheetName: string, address: string, value: string, fillColor: string): Uint8Array {
  const archive = unzipSync(xlsxBytesForTargetValue(sheetName, address, value))
  const stylesXml = strFromU8(archive['xl/styles.xml'] ?? new Uint8Array())
  const sheetXml = strFromU8(archive['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  archive['xl/styles.xml'] = strToU8(addTargetFillStyle(stylesXml, fillColor))
  archive['xl/worksheets/sheet1.xml'] = strToU8(addTargetRowStyle(sheetXml, address, 1))
  return zipSync(archive)
}

function xlsxBytesForTargetColumnFill(sheetName: string, address: string, value: string, fillColor: string): Uint8Array {
  const archive = unzipSync(xlsxBytesForTargetValue(sheetName, address, value))
  const stylesXml = strFromU8(archive['xl/styles.xml'] ?? new Uint8Array())
  const sheetXml = strFromU8(archive['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  archive['xl/styles.xml'] = strToU8(addTargetFillStyle(stylesXml, fillColor))
  archive['xl/worksheets/sheet1.xml'] = strToU8(addTargetColumnStyle(sheetXml, address, 1))
  return zipSync(archive)
}

function addTargetFillStyle(stylesXml: string, fillColor: string): string {
  const rgb = `FF${fillColor.replace(/^#/u, '').toUpperCase()}`
  return stylesXml
    .replace(
      /<fills count="2">[\s\S]*?<\/fills>/u,
      `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill></fills>`,
    )
    .replace(
      /<cellXfs count="1">([\s\S]*?)<\/cellXfs>/u,
      '<cellXfs count="2">$1<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>',
    )
}

function addTargetThemeFillStyle(stylesXml: string): string {
  return stylesXml
    .replace(
      /<fills count="2">[\s\S]*?<\/fills>/u,
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill></fills>',
    )
    .replace(
      /<cellXfs count="1">([\s\S]*?)<\/cellXfs>/u,
      '<cellXfs count="2">$1<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>',
    )
}

function addTargetIndexedFillStyle(stylesXml: string, colorIndex: number): string {
  return stylesXml
    .replace(
      /<fills count="2">[\s\S]*?<\/fills>/u,
      `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor indexed="${String(
        colorIndex,
      )}"/><bgColor indexed="64"/></patternFill></fill></fills>`,
    )
    .replace(
      /<cellXfs count="1">([\s\S]*?)<\/cellXfs>/u,
      '<cellXfs count="2">$1<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>',
    )
}

function setThemeAccent1(themeXml: string, fillColor: string): string {
  const rgb = fillColor.replace(/^#/u, '').toUpperCase()
  return themeXml.replace(/<a:accent1>[\s\S]*?<\/a:accent1>/u, `<a:accent1><a:srgbClr val="${rgb}"/></a:accent1>`)
}

function addTargetCellStyle(sheetXml: string, address: string, styleIndex: number): string {
  return sheetXml.replace(new RegExp(`<c r="${address}"`, 'u'), `<c r="${address}" s="${String(styleIndex)}"`)
}

function addTargetRowStyle(sheetXml: string, address: string, styleIndex: number): string {
  const rowNumber = address.match(/[0-9]+$/u)?.[0] ?? ''
  return sheetXml.replace(new RegExp(`<row r="${rowNumber}"`, 'u'), `<row r="${rowNumber}" s="${String(styleIndex)}" customFormat="1"`)
}

function addTargetColumnStyle(sheetXml: string, address: string, styleIndex: number): string {
  const columnIndex = columnIndexFromAddress(address)
  return sheetXml.replace(
    /<sheetData>/u,
    `<cols><col min="${String(columnIndex)}" max="${String(columnIndex)}" width="10" customWidth="1" style="${String(
      styleIndex,
    )}"/></cols><sheetData>`,
  )
}

function columnIndexFromAddress(address: string): number {
  const columnLetters = address.match(/^[A-Z]+/u)?.[0] ?? ''
  let columnIndex = 0
  for (const letter of columnLetters) {
    columnIndex = columnIndex * 26 + letter.charCodeAt(0) - 64
  }
  return columnIndex
}
