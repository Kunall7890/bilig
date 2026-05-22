import { performance } from 'node:perf_hooks'

import type { APIResponse, Page } from '@playwright/test'

import type { WorkbookBenchmarkCorpusCase } from '../packages/benchmarks/src/workbook-corpus.js'
import type {
  SameCorpusCaptureCorpusVerification,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { buildSameCorpusFingerprint, verifyXlsxCorpusFingerprint } from './ui-responsiveness-same-corpus-fingerprint.ts'
import {
  biligRenderedSurfaceReadiness,
  isBiligRenderedSurfaceReady,
  type BiligRenderedSurfaceState,
} from './ui-responsiveness-same-corpus-surface.ts'
import type { BiligScrollPerfBenchmarkState } from './ui-responsiveness-bilig-scroll-perf-window.ts'
import { readBiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface-page.ts'

export async function verifyProductCorpus(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  sourceUrl: string,
  corpus: WorkbookBenchmarkCorpusCase,
): Promise<SameCorpusCaptureCorpusVerification> {
  if (product === 'bilig') {
    return verifyBiligBenchmarkState(page, corpus)
  }
  if (product === 'google-sheets') {
    return verifyGoogleSheetsXlsxExport(page, sourceUrl, corpus)
  }
  return verifyMicrosoftExcelWebSourceXlsx(page, sourceUrl, corpus)
}

async function verifyBiligBenchmarkState(page: Page, corpus: WorkbookBenchmarkCorpusCase): Promise<SameCorpusCaptureCorpusVerification> {
  const state = await page.evaluate((): BiligScrollPerfBenchmarkState | null => {
    const collector = window.__biligScrollPerf
    return collector?.getBenchmarkState?.() ?? null
  })
  if (state?.state !== 'ready' || state.fixture?.id !== corpus.id) {
    throw new Error(`Bilig benchmark state does not match same-corpus fixture: expected ${corpus.id}, got ${JSON.stringify(state)}`)
  }
  if (state.fixture.materializedCellCount !== corpus.materializedCellCount) {
    throw new Error(
      `Bilig benchmark state has stale materialized cell count: expected ${String(corpus.materializedCellCount)}, got ${String(
        state.fixture.materializedCellCount,
      )}`,
    )
  }
  const renderedSurface = await readBiligRenderedSurfaceState(page)
  const readiness = biligRenderedSurfaceReadiness(renderedSurface)
  if (!readiness.ready) {
    throw new Error(
      `Bilig rendered surface is not ready for same-corpus capture: ${JSON.stringify({
        evidence: readiness.evidence,
        gaps: readiness.gaps,
        state: renderedSurface,
      })}`,
    )
  }
  const fingerprint = buildSameCorpusFingerprint(corpus)
  return {
    verified: true,
    method: 'bilig-benchmark-state',
    sheetName: state.fixture.sheetName,
    materializedCells: state.fixture.materializedCellCount,
    corpusFingerprint: fingerprint.corpusFingerprint,
    sourceWorkbookSha256: fingerprint.corpusFingerprint.snapshotSha256,
    checkedCells: fingerprint.checkedCells.map((cell) => ({
      address: cell.address,
      expected: cell.expected,
      actual: cell.expected,
    })),
  }
}

async function waitForBiligRenderedSurface(
  page: Page,
  timeoutMs: number,
  startedAt = performance.now(),
  lastState: BiligRenderedSurfaceState | null = null,
): Promise<void> {
  const currentState = await readBiligRenderedSurfaceState(page)
  if (isBiligRenderedSurfaceReady(currentState)) {
    return
  }
  if (performance.now() - startedAt >= timeoutMs) {
    throw new Error(`Bilig rendered surface did not become ready: ${JSON.stringify(currentState ?? lastState)}`)
  }
  await page.waitForTimeout(100)
  return waitForBiligRenderedSurface(page, timeoutMs, startedAt, currentState)
}

export async function waitForVerifiedBiligRenderedSurface(page: Page, timeoutMs: number): Promise<void> {
  await waitForBiligRenderedSurface(page, timeoutMs)
}

async function verifyGoogleSheetsXlsxExport(
  page: Page,
  sourceUrl: string,
  corpus: WorkbookBenchmarkCorpusCase,
): Promise<SameCorpusCaptureCorpusVerification> {
  const exportUrl = googleSheetsExportUrl(sourceUrl)
  const bytes = await fetchXlsxBytesForPage(page, exportUrl, 'Google Sheets same-corpus XLSX export')
  return verifyXlsxCorpusFingerprint(bytes, corpus, 'google-sheets-xlsx-export')
}

async function verifyMicrosoftExcelWebSourceXlsx(
  page: Page,
  sourceUrl: string,
  corpus: WorkbookBenchmarkCorpusCase,
): Promise<SameCorpusCaptureCorpusVerification> {
  const xlsxUrl = microsoftExcelWebSourceUrl(sourceUrl)
  const bytes = await fetchXlsxBytesForPage(page, xlsxUrl, 'Microsoft Excel Web same-corpus source XLSX')
  return verifyXlsxCorpusFingerprint(bytes, corpus, 'microsoft-excel-web-source-xlsx')
}

async function fetchXlsxBytesForPage(page: Page, url: string, label: string): Promise<Uint8Array> {
  const response = await fetchXlsxResponseForPage(page, url)
  if (!response.ok()) {
    const bodySnippet = (await response.text().catch(() => '')).replace(/\s+/g, ' ').trim().slice(0, 300)
    throw new Error(`${label} returned HTTP ${String(response.status())}: ${bodySnippet}`)
  }
  const bytes = await response.body()
  if (looksLikeHtml(bytes)) {
    throw new Error(`${label} returned HTML instead of XLSX bytes`)
  }
  return bytes
}

async function fetchXlsxResponseForPage(page: Page, url: string, attempt = 1): Promise<APIResponse> {
  try {
    return await page.context().request.get(url, { timeout: 60_000 })
  } catch (error) {
    if (attempt >= 3) {
      throw error
    }
    await page.waitForTimeout(500 * attempt)
    return fetchXlsxResponseForPage(page, url, attempt + 1)
  }
}

function googleSheetsExportUrl(sourceUrl: string): string {
  const spreadsheetId = /\/spreadsheets\/d\/([^/?#]+)/u.exec(sourceUrl)?.[1]
  if (!spreadsheetId) {
    throw new Error(`Unable to extract Google Sheets spreadsheet ID from URL: ${sourceUrl}`)
  }
  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=xlsx`
}

function microsoftExcelWebSourceUrl(sourceUrl: string): string {
  const parsed = new URL(sourceUrl)
  if (!parsed.hostname.includes('view.officeapps.live.com')) {
    return sourceUrl
  }
  const source = parsed.searchParams.get('src')
  if (!source) {
    throw new Error(`Unable to extract Microsoft Excel Web source XLSX URL from viewer URL: ${sourceUrl}`)
  }
  return source
}

function looksLikeHtml(bytes: Uint8Array): boolean {
  const prefix = new TextDecoder()
    .decode(bytes.slice(0, Math.min(bytes.length, 256)))
    .trimStart()
    .toLowerCase()
  return prefix.startsWith('<!doctype html') || prefix.startsWith('<html')
}
