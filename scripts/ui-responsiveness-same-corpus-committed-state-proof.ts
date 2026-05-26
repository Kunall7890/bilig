import { createHash } from 'node:crypto'
import { performance } from 'node:perf_hooks'

import * as XLSX from 'xlsx'

import type { SameCorpusMutationTargetSelection } from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type { SameCorpusMutationTargetProof, SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export interface SameCorpusCommittedStatePage {
  context(): {
    readonly request: {
      get(
        url: string,
        options?: { readonly timeout?: number },
      ): Promise<{
        body(): Promise<Uint8Array>
        ok(): boolean
        status(): number
        text(): Promise<string>
      }>
    }
  }
  url(): string
  waitForTimeout(timeoutMs: number): Promise<void>
}

export interface SameCorpusMutationTargetCommittedStateProof {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: 'google-sheets-xlsx-export'
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sheetName: string
  readonly sheetId: string | null
  readonly targetRange: string
  readonly before: SameCorpusMutationTargetCommittedStatePhaseProof
  readonly after: SameCorpusMutationTargetCommittedStatePhaseProof
  readonly restored: SameCorpusMutationTargetCommittedStatePhaseProof
}

export interface SameCorpusMutationTargetCommittedStatePhaseProof {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly phase: 'before' | 'after' | 'restored'
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sheetName: string
  readonly sheetId: string | null
  readonly targetRange: string
  readonly exportUrl: string
  readonly capturedAtMs: number
  readonly workbookByteSize: number
  readonly workbookSha256: string
  readonly readback: SameCorpusMutationTargetReadback
}

export async function captureSameCorpusCommittedStatePhaseProof(args: {
  readonly expectedReadback: SameCorpusMutationTargetReadback
  readonly page: SameCorpusCommittedStatePage
  readonly phase: SameCorpusMutationTargetCommittedStatePhaseProof['phase']
  readonly pollIntervalMs?: number
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection
  readonly timeoutMs?: number
  readonly workload: UiResponsivenessSameCorpusWorkload
}): Promise<SameCorpusMutationTargetCommittedStatePhaseProof | null> {
  if (args.product !== 'google-sheets') {
    return null
  }
  const exportUrl = googleSheetsExportUrl(args.page.url())
  const timeoutMs = Math.max(0, args.timeoutMs ?? 30_000)
  const pollIntervalMs = Math.max(0, args.pollIntervalMs ?? 500)
  const maxAttempts = Math.max(1, Math.ceil(timeoutMs / Math.max(1, pollIntervalMs)) + 1)
  return await captureGoogleSheetsCommittedStatePhaseProofUntilMatched(args, exportUrl, maxAttempts, pollIntervalMs)
}

async function captureGoogleSheetsCommittedStatePhaseProofUntilMatched(
  args: Parameters<typeof captureSameCorpusCommittedStatePhaseProof>[0],
  exportUrl: string,
  attemptsRemaining: number,
  pollIntervalMs: number,
): Promise<SameCorpusMutationTargetCommittedStatePhaseProof> {
  const proof = await captureGoogleSheetsCommittedStatePhaseProofSnapshot(args, exportUrl)
  if (sameCorpusCommittedReadbackMatches(args.workload, args.expectedReadback, proof.readback)) {
    return proof
  }
  if (attemptsRemaining <= 1) {
    throw new Error(
      `Google Sheets committed-state XLSX export did not match expected ${args.phase} target readback for ${args.workload} ${
        args.target.targetRange
      }: expected ${sameCorpusCommittedStateReadbackSummary(args.expectedReadback)}, last export ${sameCorpusCommittedStateReadbackSummary(
        proof.readback,
      )}`,
    )
  }
  await args.page.waitForTimeout(pollIntervalMs)
  return await captureGoogleSheetsCommittedStatePhaseProofUntilMatched(args, exportUrl, attemptsRemaining - 1, pollIntervalMs)
}

async function captureGoogleSheetsCommittedStatePhaseProofSnapshot(
  args: {
    readonly page: SameCorpusCommittedStatePage
    readonly phase: SameCorpusMutationTargetCommittedStatePhaseProof['phase']
    readonly product: UiResponsivenessSameCorpusProduct
    readonly sampleIndex: number
    readonly target: SameCorpusMutationTargetSelection
    readonly workload: UiResponsivenessSameCorpusWorkload
  },
  exportUrl: string,
): Promise<SameCorpusMutationTargetCommittedStatePhaseProof> {
  const bytes = await fetchGoogleSheetsXlsxExport(args.page, exportUrl)
  return {
    product: args.product,
    phase: args.phase,
    sampleIndex: args.sampleIndex,
    workload: args.workload,
    sheetName: args.target.sheetName,
    sheetId: args.target.sheetId,
    targetRange: args.target.targetRange,
    exportUrl,
    capturedAtMs: performance.now(),
    workbookByteSize: bytes.byteLength,
    workbookSha256: sha256Hex(bytes),
    readback: readGoogleSheetsExportTargetReadback(bytes, args.target),
  }
}

function sameCorpusCommittedStateReadbackSummary(readback: SameCorpusMutationTargetReadback | null): string {
  if (!readback) {
    return 'missing'
  }
  return JSON.stringify({
    fillColor: readback.fillColor,
    formula: readback.formula,
    source: readback.source,
    value: readback.value,
    visibleText: readback.visibleText,
  })
}

export function buildSameCorpusCommittedStateProof(args: {
  readonly after: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly before: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly product: UiResponsivenessSameCorpusProduct
  readonly restored: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection
  readonly workload: UiResponsivenessSameCorpusWorkload
}): SameCorpusMutationTargetCommittedStateProof | null {
  if (args.product !== 'google-sheets' || !args.before || !args.after || !args.restored) {
    return null
  }
  return {
    product: args.product,
    source: 'google-sheets-xlsx-export',
    sampleIndex: args.sampleIndex,
    workload: args.workload,
    sheetName: args.target.sheetName,
    sheetId: args.target.sheetId,
    targetRange: args.target.targetRange,
    before: args.before,
    after: args.after,
    restored: args.restored,
  }
}

export function sameCorpusMutationTargetCommittedStateInvalidReasons(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
): string[] {
  if (product !== 'google-sheets') {
    return []
  }
  const proof = sample.committedStateProof
  if (!proof) {
    return [`semantic UI mutation target proof for ${workload} is missing independent Google Sheets committed-state proof`]
  }
  const invalidReasons: string[] = []
  if (proof.product !== 'google-sheets' || proof.product !== sample.product || proof.source !== 'google-sheets-xlsx-export') {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state proof is not from Google Sheets XLSX export`)
  }
  if (proof.sampleIndex !== sample.sampleIndex) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state proof has mismatched sample`)
  }
  if (proof.workload !== workload || proof.workload !== sample.workload) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state proof has mismatched workload`)
  }
  if (proof.sheetName !== sample.sheetName || proof.sheetId !== sample.sheetId) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state proof has mismatched sheet identity`)
  }
  if (normalizeSameCorpusRange(proof.targetRange) !== normalizeSameCorpusRange(sample.targetRange)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state proof has mismatched target range`)
  }
  invalidReasons.push(...sameCorpusCommittedStatePhaseInvalidReasons(workload, sample, proof))
  invalidReasons.push(...sameCorpusCommittedStateReadbackInvalidReasons(workload, sample, proof))
  return invalidReasons
}

function sameCorpusCommittedStatePhaseInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  proof: SameCorpusMutationTargetCommittedStateProof,
): string[] {
  const invalidReasons: string[] = []
  for (const phase of ['before', 'after', 'restored'] as const) {
    const phaseProof = proof[phase]
    if (phaseProof.phase !== phase) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state ${phase} proof has mismatched phase`)
    }
    if (
      phaseProof.product !== proof.product ||
      phaseProof.sampleIndex !== sample.sampleIndex ||
      phaseProof.workload !== workload ||
      phaseProof.sheetName !== sample.sheetName ||
      phaseProof.sheetId !== sample.sheetId ||
      normalizeSameCorpusRange(phaseProof.targetRange) !== normalizeSameCorpusRange(sample.targetRange)
    ) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state ${phase} proof has mismatched identity`)
    }
    if (!phaseProof.exportUrl.includes('/spreadsheets/d/') || !phaseProof.exportUrl.includes('format=xlsx')) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state ${phase} proof is missing export URL`)
    }
    if (!sameGoogleSheetsExportWorkbook(proof.before.exportUrl, phaseProof.exportUrl)) {
      invalidReasons.push(
        `semantic UI mutation target proof for ${workload} committed-state ${phase} proof is from a different spreadsheet export URL`,
      )
    }
    if (!Number.isFinite(phaseProof.capturedAtMs) || phaseProof.capturedAtMs < 0) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state ${phase} proof is missing capture timing`)
    }
    if (!Number.isFinite(phaseProof.workbookByteSize) || phaseProof.workbookByteSize <= 0 || !isSha256(phaseProof.workbookSha256)) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state ${phase} proof is missing workbook bytes`)
    }
    if (phaseProof.readback.source !== 'google-sheets-xlsx-export') {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state ${phase} readback is not from XLSX export`)
    }
    if (sameCorpusCommittedStateReadbackCarriesBrowserMetadata(phaseProof.readback)) {
      invalidReasons.push(
        `semantic UI mutation target proof for ${workload} committed-state ${phase} readback carries browser-only proof metadata`,
      )
    }
  }
  invalidReasons.push(...sameCorpusCommittedStatePhaseTimingInvalidReasons(workload, sample, proof))
  if (proof.before.workbookSha256 === proof.after.workbookSha256) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state export reused the pre-mutation workbook hash`)
  }
  if (proof.after.workbookSha256 === proof.restored.workbookSha256) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} committed-state export reused the post-mutation workbook hash after restore`,
    )
  }
  return invalidReasons
}

function sameCorpusCommittedStateReadbackCarriesBrowserMetadata(readback: SameCorpusMutationTargetReadback): boolean {
  return (
    Object.hasOwn(readback, 'batchId') || Object.hasOwn(readback, 'capturedRevision') || Object.hasOwn(readback, 'visibleSceneProofSha256')
  )
}

function sameCorpusCommittedStatePhaseTimingInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  proof: SameCorpusMutationTargetCommittedStateProof,
): string[] {
  const invalidReasons: string[] = []
  if (proof.before.capturedAtMs >= sample.operationStartedAtMs) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} committed-state before export was not captured before mutation started`,
    )
  }
  if (proof.after.capturedAtMs < sample.operationStartedAtMs || proof.after.capturedAtMs > sample.postMutationProofCapturedAtMs) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} committed-state after export was not captured inside the post-mutation proof window`,
    )
  }
  if (proof.restored.capturedAtMs < sample.postMutationProofCapturedAtMs || proof.restored.capturedAtMs > sample.restoreProofCapturedAtMs) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} committed-state restored export was not captured inside the restore proof window`,
    )
  }
  return invalidReasons
}

function sameGoogleSheetsExportWorkbook(leftUrl: string, rightUrl: string): boolean {
  const leftId = googleSheetsSpreadsheetId(leftUrl)
  const rightId = googleSheetsSpreadsheetId(rightUrl)
  return leftId !== null && leftId === rightId
}

function googleSheetsSpreadsheetId(exportUrl: string): string | null {
  try {
    const url = new URL(exportUrl)
    const match = /^\/spreadsheets\/d\/([^/]+)\/export$/u.exec(url.pathname)
    const format = url.searchParams.get('format')
    return match && format === 'xlsx' ? decodeURIComponent(match[1] ?? '') : null
  } catch {
    return null
  }
}

function sameCorpusCommittedStateReadbackInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  proof: SameCorpusMutationTargetCommittedStateProof,
): string[] {
  const invalidReasons: string[] = []
  if (!sameCorpusCommittedReadbackMatches(workload, proof.before.readback, proof.after.readback)) {
    // Expected path; the export must prove the workbook changed.
  } else {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state export did not prove a before/after change`)
  }
  if (!sameCorpusCommittedReadbackMatches(workload, sample.after, proof.after.readback)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state after readback does not match target proof`)
  }
  if (!sameCorpusCommittedReadbackMatches(workload, sample.before, proof.before.readback)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state before readback does not match target proof`)
  }
  if (!sameCorpusCommittedReadbackMatches(workload, sample.restored, proof.restored.readback)) {
    invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state restored readback does not match target proof`)
  }
  if (!sameCorpusCommittedReadbackMatches(workload, proof.before.readback, proof.restored.readback)) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} committed-state restore readback does not match pre-mutation export`,
    )
  }
  return invalidReasons
}

function sameCorpusCommittedReadbackMatches(
  workload: UiResponsivenessSameCorpusWorkload,
  left: SameCorpusMutationTargetReadback,
  right: SameCorpusMutationTargetReadback,
): boolean {
  if (workload === 'formula-edit') {
    return left.formula === right.formula
  }
  if (workload === 'fill-format-change') {
    return normalizeSameCorpusColor(left.fillColor) === normalizeSameCorpusColor(right.fillColor)
  }
  return left.value === right.value || left.visibleText === right.visibleText
}

async function fetchGoogleSheetsXlsxExport(page: SameCorpusCommittedStatePage, exportUrl: string): Promise<Uint8Array> {
  const response = await page.context().request.get(exportUrl, { timeout: 60_000 })
  if (!response.ok()) {
    const bodySnippet = (await response.text().catch(() => '')).replace(/\s+/g, ' ').trim().slice(0, 300)
    throw new Error(`Google Sheets committed-state XLSX export returned HTTP ${String(response.status())}: ${bodySnippet}`)
  }
  const bytes = await response.body()
  if (looksLikeHtml(bytes)) {
    throw new Error('Google Sheets committed-state XLSX export returned HTML instead of XLSX bytes')
  }
  return bytes
}

function googleSheetsExportUrl(sourceUrl: string): string {
  const spreadsheetId = /\/spreadsheets\/d\/([^/?#]+)/u.exec(sourceUrl)?.[1]
  if (!spreadsheetId) {
    throw new Error(`Unable to extract Google Sheets spreadsheet ID from URL: ${sourceUrl}`)
  }
  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=xlsx`
}

function readGoogleSheetsExportTargetReadback(
  bytes: Uint8Array,
  target: SameCorpusMutationTargetSelection,
): SameCorpusMutationTargetReadback {
  const workbook = XLSX.read(Buffer.from(bytes), {
    cellFormula: true,
    cellStyles: true,
    cellText: false,
    type: 'buffer',
  })
  const worksheet = workbook.Sheets[target.sheetName]
  if (!worksheet) {
    throw new Error(`Google Sheets committed-state XLSX export is missing sheet ${target.sheetName}`)
  }
  const cell = worksheet[normalizeTargetStartAddress(target.startAddress)]
  const formula = typeof cell?.f === 'string' && cell.f.trim().length > 0 ? `=${cell.f.trim().replace(/^=/u, '')}` : null
  const value = formula ? null : normalizeSpreadsheetValue(cell?.v)
  return {
    value,
    formula,
    fillColor: readXlsxCellFillColor(cell),
    visibleText: formula ?? value,
    source: 'google-sheets-xlsx-export',
  }
}

function readXlsxCellFillColor(cell: XLSX.CellObject | undefined): string | null {
  const style = asRecord(readXlsxCellStyle(cell))
  const fill = asRecord(style?.fill)
  const fgColor = asRecord(fill?.fgColor)
  const bgColor = asRecord(fill?.bgColor)
  return normalizeXlsxRgb(fgColor?.rgb) ?? normalizeXlsxRgb(bgColor?.rgb) ?? null
}

function readXlsxCellStyle(cell: XLSX.CellObject | undefined): unknown {
  if (!cell || typeof cell !== 'object' || !('s' in cell)) {
    return null
  }
  return cell.s
}

function normalizeXlsxRgb(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const hex = value.trim().replace(/^#/u, '').toLowerCase()
  if (/^[0-9a-f]{8}$/u.test(hex)) {
    return `#${hex.slice(2)}`
  }
  return /^[0-9a-f]{6}$/u.test(hex) ? `#${hex}` : null
}

function normalizeSameCorpusColor(value: string | null | undefined): string | null {
  const trimmed = value?.trim().toLowerCase() ?? ''
  if (/^#[0-9a-f]{6}$/u.test(trimmed)) {
    return trimmed
  }
  const rgbMatch = trimmed.match(/^rgba?\((\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})(?:,\s*(?:0|0?\.\d+|1(?:\.0)?))?\)$/u)
  if (!rgbMatch) {
    return trimmed.length === 0 ? null : trimmed
  }
  const channels = rgbMatch.slice(1, 4).map((channel) => Number(channel))
  if (channels.some((channel) => !Number.isInteger(channel) || channel < 0 || channel > 255)) {
    return null
  }
  return `#${channels.map((channel) => channel.toString(16).padStart(2, '0')).join('')}`
}

function normalizeTargetStartAddress(address: string): string {
  return address.split('!').at(-1)?.replace(/\$/gu, '').trim().toUpperCase() ?? address
}

function normalizeSameCorpusRange(range: string): string {
  return range.split('!').at(-1)?.replace(/\$/gu, '').trim().toUpperCase() ?? ''
}

function normalizeSpreadsheetValue(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null
  }
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value)
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  if (value instanceof Date) {
    return value.toISOString()
  }
  const serialized = JSON.stringify(value)
  return serialized === undefined ? null : serialized
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === 'object' ? Object.fromEntries(Object.entries(value)) : null
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function isSha256(value: string): boolean {
  return /^[a-f0-9]{64}$/u.test(value)
}

function looksLikeHtml(bytes: Uint8Array): boolean {
  const prefix = new TextDecoder()
    .decode(bytes.slice(0, Math.min(bytes.length, 256)))
    .trimStart()
    .toLowerCase()
  return prefix.startsWith('<!doctype html') || prefix.startsWith('<html')
}
