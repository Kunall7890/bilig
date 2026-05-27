import { createHash } from 'node:crypto'
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import { strFromU8, unzipSync } from 'fflate'
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
  waitForFunction?(pageFunction: () => boolean, arg?: unknown, options?: { readonly timeout?: number }): Promise<unknown>
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
  readonly artifactPath?: string | null
  readonly artifactSha256?: string | null
  readonly workbookByteSize: number
  readonly workbookSha256: string
  readonly readback: SameCorpusMutationTargetReadback
}

export async function captureSameCorpusCommittedStatePhaseProof(args: {
  readonly expectedReadback: SameCorpusMutationTargetReadback
  readonly artifactPath?: string | null
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
  return await captureGoogleSheetsCommittedStatePhaseProofUntilMatched(args, exportUrl, performance.now() + timeoutMs, pollIntervalMs)
}

export function sameCorpusCommittedStateProofArtifactPath(args: {
  readonly caseId?: string
  readonly outputPath: string
  readonly phase: SameCorpusMutationTargetCommittedStatePhaseProof['phase']
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusWorkload
}): string {
  const caseId = args.caseId ?? `same-corpus-${args.workload}`
  return resolve(
    `${args.outputPath}.proof`,
    caseId,
    'committed-state',
    `google-sheets-sample-${String(args.sampleIndex + 1)}-${args.phase}.json`,
  )
}

async function captureGoogleSheetsCommittedStatePhaseProofUntilMatched(
  args: Parameters<typeof captureSameCorpusCommittedStatePhaseProof>[0],
  exportUrl: string,
  deadlineMs: number,
  pollIntervalMs: number,
): Promise<SameCorpusMutationTargetCommittedStatePhaseProof> {
  await waitForGoogleSheetsSaveIdle(args.page, Math.max(0, deadlineMs - performance.now()))
  const proof = await captureGoogleSheetsCommittedStatePhaseProofSnapshot(args, exportUrl)
  if (sameCorpusCommittedReadbackMatches(args.workload, args.expectedReadback, proof.readback)) {
    return proof
  }
  const remainingMs = deadlineMs - performance.now()
  if (remainingMs <= 0) {
    throwGoogleSheetsCommittedStateMismatch(args, proof)
  }
  await args.page.waitForTimeout(Math.min(pollIntervalMs, remainingMs))
  return await captureGoogleSheetsCommittedStatePhaseProofUntilMatched(args, exportUrl, deadlineMs, pollIntervalMs)
}

function throwGoogleSheetsCommittedStateMismatch(
  args: Parameters<typeof captureSameCorpusCommittedStatePhaseProof>[0],
  proof: SameCorpusMutationTargetCommittedStatePhaseProof,
): never {
  throw new Error(
    `Google Sheets committed-state XLSX export did not match expected ${args.phase} target readback for ${args.workload} ${
      args.target.targetRange
    }: expected ${sameCorpusCommittedStateReadbackSummary(args.expectedReadback)}, last export ${sameCorpusCommittedStateReadbackSummary(
      proof.readback,
    )}`,
  )
}

async function waitForGoogleSheetsSaveIdle(page: SameCorpusCommittedStatePage, timeoutMs: number): Promise<void> {
  if (!page.url().includes('docs.google.com/spreadsheets') || !page.waitForFunction || timeoutMs <= 0) {
    return
  }
  await page
    .waitForFunction(
      () => {
        const text = document.body?.innerText.replace(/\s+/gu, ' ').trim() ?? ''
        return !/\bSaving(?:\.\.\.|…)?\b/iu.test(text)
      },
      undefined,
      { timeout: Math.max(1, Math.min(10_000, timeoutMs)) },
    )
    .catch(() => undefined)
}

async function captureGoogleSheetsCommittedStatePhaseProofSnapshot(
  args: {
    readonly artifactPath?: string | null
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
  const proof = {
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
  return args.artifactPath ? writeCommittedStateProofArtifact(proof, args.artifactPath) : proof
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
    invalidReasons.push(...sameCorpusCommittedStateArtifactInvalidReasons(workload, sample, phaseProof))
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

function sameCorpusCommittedStateArtifactInvalidReasons(
  workload: UiResponsivenessSameCorpusWorkload,
  sample: SameCorpusMutationTargetProof,
  phaseProof: SameCorpusMutationTargetCommittedStatePhaseProof,
): string[] {
  const artifactPath = phaseProof.artifactPath?.trim() ?? ''
  const artifactSha256 = phaseProof.artifactSha256?.trim().toLowerCase() ?? ''
  const invalidReasons: string[] = []
  if (!artifactPath || !isSha256(artifactSha256)) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} committed-state ${phaseProof.phase} proof is missing archive JSON artifact`,
    )
    return invalidReasons
  }
  const normalizedPath = artifactPath.replaceAll('\\', '/')
  const expectedFile = `google-sheets-sample-${String(sample.sampleIndex + 1)}-${phaseProof.phase}.json`
  if (!normalizedPath.includes('/committed-state/') || !normalizedPath.endsWith(`/${expectedFile}`)) {
    invalidReasons.push(
      `semantic UI mutation target proof for ${workload} committed-state ${phaseProof.phase} artifact path is not tied to the sample`,
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
  if (workload === 'formula-edit') {
    if (sample.intendedPayload.kind !== 'formula' || proof.after.readback.formula !== sample.intendedPayload.formula) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state after readback did not prove formula text`)
    }
    const expectedRenderedValue = String(sample.sampleIndex + 2)
    if (proof.after.readback.value !== expectedRenderedValue && proof.after.readback.visibleText !== expectedRenderedValue) {
      invalidReasons.push(`semantic UI mutation target proof for ${workload} committed-state after readback did not prove formula result`)
    }
  }
  return invalidReasons
}

function sameCorpusCommittedReadbackMatches(
  workload: UiResponsivenessSameCorpusWorkload,
  left: SameCorpusMutationTargetReadback,
  right: SameCorpusMutationTargetReadback,
): boolean {
  if (workload === 'formula-edit') {
    if (left.formula !== null && right.formula !== null) {
      return left.formula === right.formula
    }
    return sameCorpusReadbackTextValue(left) !== null && sameCorpusReadbackTextValue(left) === sameCorpusReadbackTextValue(right)
  }
  if (workload === 'fill-format-change') {
    return normalizeSameCorpusColor(left.fillColor) === normalizeSameCorpusColor(right.fillColor)
  }
  return left.value === right.value || left.visibleText === right.visibleText
}

function sameCorpusReadbackTextValue(readback: SameCorpusMutationTargetReadback): string | null {
  return readback.value ?? readback.visibleText
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
  const value = normalizeSpreadsheetValue(cell?.v)
  return {
    value,
    formula,
    fillColor: readXlsxCellFillColor(cell) ?? readOoxmlCellFillColor(bytes, target),
    visibleText: value ?? formula,
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

function readOoxmlCellFillColor(bytes: Uint8Array, target: SameCorpusMutationTargetSelection): string | null {
  const archive = unzipSync(bytes)
  const sheetPath = readOoxmlSheetPath(archive, target.sheetName)
  if (!sheetPath) {
    return null
  }
  const sheetXml = readOoxmlText(archive, sheetPath)
  const stylesXml = readOoxmlText(archive, 'xl/styles.xml')
  if (!sheetXml || !stylesXml) {
    return null
  }
  const styleIndex = readOoxmlCellStyleIndex(sheetXml, normalizeTargetStartAddress(target.startAddress))
  if (styleIndex === null) {
    return null
  }
  const fillId = readOoxmlCellXfFillId(stylesXml, styleIndex)
  return fillId === null ? null : readOoxmlFillColor(stylesXml, fillId)
}

function readOoxmlSheetPath(archive: Record<string, Uint8Array>, sheetName: string): string | null {
  const workbookXml = readOoxmlText(archive, 'xl/workbook.xml')
  if (!workbookXml) {
    return null
  }
  const sheet = readOoxmlSheet(workbookXml, sheetName)
  if (!sheet) {
    return null
  }
  const relsXml = readOoxmlText(archive, 'xl/_rels/workbook.xml.rels')
  const relationshipTarget = sheet.relationshipId && relsXml ? readOoxmlRelationshipTarget(relsXml, sheet.relationshipId) : null
  if (relationshipTarget) {
    return normalizeOoxmlWorkbookRelativePath(relationshipTarget)
  }
  return sheet.index >= 0 ? `xl/worksheets/sheet${String(sheet.index + 1)}.xml` : null
}

function readOoxmlSheet(workbookXml: string, sheetName: string): { readonly index: number; readonly relationshipId: string | null } | null {
  const normalizedName = decodeXmlAttribute(sheetName)
  const sheetTags = workbookXml.match(/<sheet\b[^>]*\/?>/gu) ?? []
  for (const [index, tag] of sheetTags.entries()) {
    if (decodeXmlAttribute(readXmlAttribute(tag, 'name') ?? '') !== normalizedName) {
      continue
    }
    return {
      index,
      relationshipId: readXmlAttribute(tag, 'r:id'),
    }
  }
  return null
}

function readOoxmlRelationshipTarget(relsXml: string, relationshipId: string): string | null {
  const relationshipTags = relsXml.match(/<Relationship\b[^>]*\/?>/gu) ?? []
  for (const tag of relationshipTags) {
    if (readXmlAttribute(tag, 'Id') === relationshipId) {
      return readXmlAttribute(tag, 'Target')
    }
  }
  return null
}

function readOoxmlCellStyleIndex(sheetXml: string, address: string): number | null {
  const cellTagMatch = new RegExp(`<c\\b(?=[^>]*\\br="${escapeRegExp(address)}")[^>]*>`, 'u').exec(sheetXml)
  if (!cellTagMatch) {
    return null
  }
  const styleIndex = Number(readXmlAttribute(cellTagMatch[0], 's'))
  return Number.isInteger(styleIndex) && styleIndex >= 0 ? styleIndex : null
}

function readOoxmlCellXfFillId(stylesXml: string, styleIndex: number): number | null {
  const cellXfsXml = readXmlSection(stylesXml, 'cellXfs')
  const xfTag = (cellXfsXml.match(/<xf\b[^>]*(?:\/>|>)/gu) ?? [])[styleIndex]
  if (!xfTag) {
    return null
  }
  const fillId = Number(readXmlAttribute(xfTag, 'fillId'))
  return Number.isInteger(fillId) && fillId >= 0 ? fillId : null
}

function readOoxmlFillColor(stylesXml: string, fillId: number): string | null {
  const fillsXml = readXmlSection(stylesXml, 'fills')
  const fillXml = (fillsXml.match(/<fill\b[^>]*>[\s\S]*?<\/fill>/gu) ?? [])[fillId]
  if (!fillXml) {
    return null
  }
  return normalizeXlsxRgb(readXmlAttribute(fillXml, 'rgb', 'fgColor')) ?? normalizeXlsxRgb(readXmlAttribute(fillXml, 'rgb', 'bgColor'))
}

function readOoxmlText(archive: Record<string, Uint8Array>, path: string): string | null {
  const bytes = archive[path]
  return bytes ? strFromU8(bytes) : null
}

function normalizeOoxmlWorkbookRelativePath(target: string): string {
  const cleanTarget = target.replace(/^\/+|^\.\//gu, '')
  return cleanTarget.startsWith('xl/') ? cleanTarget : `xl/${cleanTarget}`
}

function readXmlSection(xml: string, tagName: string): string {
  return new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)<\\/${tagName}>`, 'u').exec(xml)?.[1] ?? ''
}

function readXmlAttribute(tag: string, attributeName: string, nestedTagName?: string): string | null {
  const source =
    nestedTagName === undefined
      ? tag
      : (new RegExp(`<${nestedTagName}\\b[^>]*>`, 'u').exec(tag)?.[0] ??
        new RegExp(`<${nestedTagName}\\b[^>]*/>`, 'u').exec(tag)?.[0] ??
        '')
  const escapedAttributeName = escapeRegExp(attributeName)
  const match =
    new RegExp(`\\b${escapedAttributeName}="([^"]*)"`, 'u').exec(source) ??
    new RegExp(`\\b${escapedAttributeName}='([^']*)'`, 'u').exec(source)
  return match?.[1] ? decodeXmlAttribute(match[1]) : null
}

function decodeXmlAttribute(value: string): string {
  return value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
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

function writeCommittedStateProofArtifact(
  proof: SameCorpusMutationTargetCommittedStatePhaseProof,
  artifactPath: string,
): SameCorpusMutationTargetCommittedStatePhaseProof {
  const artifactPayload = { ...proof, artifactPath: repoRelativePath(artifactPath) }
  const artifactJson = `${JSON.stringify(stableJsonValue(artifactPayload), null, 2)}\n`
  mkdirSync(dirname(artifactPath), { recursive: true })
  writeFileSync(artifactPath, artifactJson)
  return {
    ...proof,
    artifactPath: artifactPayload.artifactPath,
    artifactSha256: sha256Hex(new TextEncoder().encode(artifactJson)),
  }
}

function repoRelativePath(path: string): string {
  return relative(process.cwd(), path)
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entryValue]) => entryValue !== undefined)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entryValue]) => [key, stableJsonValue(entryValue)]),
    )
  }
  return value
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
