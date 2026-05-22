import { formatAddress } from '@bilig/formula'
import { formatErrorCode, ValueTag, type CellRangeRef } from '@bilig/protocol'
import type {
  WorkbookAgentRenderedCell,
  WorkbookAgentRenderedContext,
  WorkbookAgentRenderedRange,
  WorkbookAgentRenderedSurfaceProof,
} from '@bilig/contracts'
import {
  enumerateWorkbookAgentRangeAddresses,
  normalizeWorkbookAgentRange,
  toWorkbookAgentRangeRef,
  type WorkbookAgentRangeChunk,
} from './workbook-agent-range-chunks.js'

export interface WorkbookVerificationMismatch {
  readonly sheetName: string
  readonly address: string
  readonly field: string
  readonly expected: unknown
  readonly actual: unknown
  readonly source: 'authoritative' | 'rendered'
}

export interface WorkbookRenderedReadbackProof {
  readonly requested: boolean
  readonly requestedRange: CellRangeRef | null
  readonly available: boolean
  readonly matched: boolean | null
  readonly stale: boolean
  readonly capturedRange: CellRangeRef | null
  readonly sourceKind: 'selection' | 'visibleRange' | null
  readonly sourceRange: CellRangeRef | null
  readonly capturedAtUnixMs: number | null
  readonly capturedRevision: number | null
  readonly capturedBatchId: number | null
  readonly surfaceProof: WorkbookAgentRenderedSurfaceProof | null
  readonly surfaceProofMatched: boolean | null
  readonly surfaceProofIncompleteReason: string | null
  readonly truncated: boolean
  readonly sourceTruncated: boolean
  readonly missingCells: readonly string[]
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
  readonly nextChunk: WorkbookAgentRangeChunk | null
  readonly range: WorkbookAgentRenderedRange | null
}

interface AuthoritativeCellLike {
  readonly address?: unknown
  readonly input?: unknown
  readonly value?: unknown
  readonly formula?: unknown
  readonly displayFormat?: unknown
  readonly styleId?: unknown
  readonly numberFormatId?: unknown
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asNonNegativeSafeInteger(value: unknown): number | null {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 ? value : null
}

function isRenderedValueTag(value: unknown): value is ValueTag {
  return (
    value === ValueTag.Empty ||
    value === ValueTag.Number ||
    value === ValueTag.Boolean ||
    value === ValueTag.String ||
    value === ValueTag.Error
  )
}

function normalizeRenderedValue(value: unknown): unknown {
  if (!isRecord(value) || !isRenderedValueTag(value['tag'])) {
    return value
  }
  switch (value['tag']) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value['value'] ?? null
    case ValueTag.Error:
      return typeof value['code'] === 'number' ? formatErrorCode(value['code']) : '#ERROR!'
    default:
      return value
  }
}

function renderedCaptureRevision(context: WorkbookAgentRenderedContext | null | undefined): number | null {
  return asNonNegativeSafeInteger(context?.capturedRevision)
}

function renderedSurfaceProof(context: WorkbookAgentRenderedContext | null | undefined): WorkbookAgentRenderedSurfaceProof | null {
  return context?.surfaceProof ?? null
}

function valuesEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) {
    return true
  }
  return JSON.stringify(left) === JSON.stringify(right)
}

function renderedInputMismatch(input: {
  readonly authoritativeInput: unknown
  readonly authoritativeValue: unknown
  readonly renderedInput: unknown
  readonly renderedValue: unknown
}): boolean {
  if (input.authoritativeInput === null || input.renderedInput === null) {
    return false
  }
  if (valuesEqual(input.authoritativeInput, input.renderedInput)) {
    return false
  }
  return !valuesEqual(input.authoritativeInput, input.renderedValue) && !valuesEqual(input.authoritativeValue, input.renderedInput)
}

function rangeContains(container: CellRangeRef, requested: CellRangeRef): boolean {
  const source = normalizeWorkbookAgentRange(container)
  const target = normalizeWorkbookAgentRange(requested)
  return (
    source.sheetName === target.sheetName &&
    source.startRow <= target.startRow &&
    source.endRow >= target.endRow &&
    source.startCol <= target.startCol &&
    source.endCol >= target.endCol
  )
}

interface RenderedRangeCandidate {
  readonly sourceKind: 'selection' | 'visibleRange'
  readonly range: WorkbookAgentRenderedRange
}

function renderedCandidates(context: WorkbookAgentRenderedContext | null | undefined): RenderedRangeCandidate[] {
  if (!context) {
    return []
  }
  const candidates: RenderedRangeCandidate[] = []
  if (context.selection) {
    candidates.push({ sourceKind: 'selection', range: context.selection })
  }
  if (context.visibleRange) {
    candidates.push({ sourceKind: 'visibleRange', range: context.visibleRange })
  }
  return candidates
}

function pickRenderedRange(
  context: WorkbookAgentRenderedContext | null | undefined,
  requestedRange: CellRangeRef,
): RenderedRangeCandidate | null {
  const candidates = renderedCandidates(context)
  return (
    candidates.find((entry) => {
      const source = toWorkbookAgentRangeRef(entry.range.range)
      const target = toWorkbookAgentRangeRef(requestedRange)
      return source.sheetName === target.sheetName && source.startAddress === target.startAddress && source.endAddress === target.endAddress
    }) ??
    candidates.find((entry) => rangeContains(entry.range.range, requestedRange)) ??
    null
  )
}

function extractRenderedCell(input: {
  readonly renderedRange: WorkbookAgentRenderedRange
  readonly requestedRange: CellRangeRef
  readonly row: number
  readonly col: number
}): WorkbookAgentRenderedCell | null {
  const source = normalizeWorkbookAgentRange(input.renderedRange.range)
  const rowOffset = input.row - source.startRow
  const colOffset = input.col - source.startCol
  if (rowOffset < 0 || colOffset < 0) {
    return null
  }
  return input.renderedRange.rows[rowOffset]?.[colOffset] ?? null
}

function buildExtractedRenderedRange(input: {
  readonly renderedRange: WorkbookAgentRenderedRange
  readonly requestedRange: CellRangeRef
}): {
  readonly range: WorkbookAgentRenderedRange | null
  readonly missingCells: readonly string[]
} {
  const requested = normalizeWorkbookAgentRange(input.requestedRange)
  const rows: WorkbookAgentRenderedCell[][] = []
  const missingCells: string[] = []
  for (let row = requested.startRow; row <= requested.endRow; row += 1) {
    const renderedRow: WorkbookAgentRenderedCell[] = []
    for (let col = requested.startCol; col <= requested.endCol; col += 1) {
      const renderedCell = extractRenderedCell({
        renderedRange: input.renderedRange,
        requestedRange: requested,
        row,
        col,
      })
      if (!renderedCell) {
        missingCells.push(`${requested.sheetName}!${formatAddress(row, col)}`)
        continue
      }
      renderedRow.push(renderedCell)
    }
    rows.push(renderedRow)
  }
  const rowCount = requested.endRow - requested.startRow + 1
  const columnCount = requested.endCol - requested.startCol + 1
  return {
    range:
      missingCells.length > 0
        ? null
        : {
            range: toWorkbookAgentRangeRef(requested),
            rowCount,
            columnCount,
            cellCount: rowCount * columnCount,
            truncated: false,
            rows,
          },
    missingCells,
  }
}

function collectRenderedMismatches(input: {
  readonly sheetName: string
  readonly authoritativeRows: readonly (readonly unknown[])[]
  readonly renderedRange: WorkbookAgentRenderedRange
}): WorkbookVerificationMismatch[] {
  const mismatches: WorkbookVerificationMismatch[] = []
  input.authoritativeRows.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      if (!isRecord(cell)) {
        return
      }
      const authoritative: AuthoritativeCellLike = {
        address: cell['address'],
        input: cell['input'],
        value: cell['value'],
        formula: cell['formula'],
        displayFormat: cell['displayFormat'],
        styleId: cell['styleId'],
        numberFormatId: cell['numberFormatId'],
      }
      const rendered = input.renderedRange.rows[rowIndex]?.[colIndex]
      const address = typeof authoritative.address === 'string' ? authoritative.address : `R${String(rowIndex)}C${String(colIndex)}`
      if (!rendered) {
        mismatches.push({
          sheetName: input.sheetName,
          address,
          field: 'cell',
          expected: 'rendered cell present',
          actual: null,
          source: 'rendered',
        })
        return
      }
      const authoritativeInput = authoritative.input ?? null
      const authoritativeValue = authoritative.value ?? null
      const renderedInput = rendered.input ?? null
      const renderedValue = normalizeRenderedValue(rendered.value)
      const comparisons: Array<readonly [string, unknown, unknown]> = [
        ['value', authoritativeValue, renderedValue],
        ['formula', authoritative.formula ?? null, rendered.formula ?? null],
        ['styleId', authoritative.styleId ?? null, rendered.styleId ?? null],
        ['numberFormatId', authoritative.numberFormatId ?? null, rendered.numberFormatId ?? null],
      ]
      if (authoritative.displayFormat !== null && authoritative.displayFormat !== undefined) {
        comparisons.push(['displayFormat', authoritative.displayFormat, rendered.displayFormat ?? null])
      }
      if (
        renderedInputMismatch({
          authoritativeInput,
          authoritativeValue,
          renderedInput,
          renderedValue,
        })
      ) {
        mismatches.push({
          sheetName: input.sheetName,
          address,
          field: 'input',
          expected: authoritativeInput,
          actual: renderedInput,
          source: 'rendered',
        })
      }
      comparisons.forEach(([field, expected, actual]) => {
        if (!valuesEqual(expected, actual)) {
          mismatches.push({
            sheetName: input.sheetName,
            address,
            field,
            expected,
            actual,
            source: 'rendered',
          })
        }
      })
    })
  })
  return mismatches
}

function buildIncompleteReason(input: {
  readonly hasRenderedContext: boolean
  readonly selectedRange: WorkbookAgentRenderedRange | null
  readonly captureStale: boolean
  readonly surfaceProofIncompleteReason: string | null
  readonly missingCells: readonly string[]
  readonly truncated: boolean
  readonly mismatches: readonly WorkbookVerificationMismatch[]
}): string | null {
  if (!input.hasRenderedContext) {
    return 'No browser-rendered context was attached to this tool call.'
  }
  if (!input.selectedRange) {
    return 'Requested range was not captured in the rendered selection or visible viewport.'
  }
  if (input.captureStale) {
    return 'Rendered capture is older than the requested verification revision.'
  }
  if (input.surfaceProofIncompleteReason) {
    return input.surfaceProofIncompleteReason
  }
  if (input.missingCells.length > 0) {
    return 'Rendered capture was incomplete for the requested range.'
  }
  if (input.truncated) {
    return 'Rendered capture was truncated; verify the next chunk before treating the whole range as proven.'
  }
  if (input.mismatches.length > 0) {
    return 'Rendered capture does not match authoritative workbook state.'
  }
  return null
}

function buildSurfaceProofIncompleteReason(input: {
  readonly minRevision?: number | null | undefined
  readonly proof: WorkbookAgentRenderedSurfaceProof | null
}): string | null {
  if (typeof input.minRevision !== 'number') {
    return null
  }
  const proof = input.proof
  if (!proof) {
    return 'No browser-presented TypeGPU frame proof was attached to this rendered readback.'
  }
  if (proof.mode !== 'typegpu-v3') {
    return 'Rendered readback was not proven by the TypeGPU V3 renderer.'
  }
  if (proof.backendStatus !== 'ready') {
    return 'TypeGPU renderer backend was not ready when the rendered readback was captured.'
  }
  if (proof.frameProofStatus !== 'presented' || !proof.hasPresentedVisibleFrame) {
    return 'Rendered readback has not been presented in a browser-visible TypeGPU frame.'
  }
  if (proof.presentedTilePaneCount <= 0 || proof.presentedHeaderPaneCount <= 0) {
    return 'Presented TypeGPU frame proof did not include visible grid tiles and headers.'
  }
  if (proof.surfaceWidth <= 0 || proof.surfaceHeight <= 0 || proof.surfacePixelWidth <= 0 || proof.surfacePixelHeight <= 0) {
    return 'Presented TypeGPU frame proof did not cover a non-empty visible surface.'
  }
  const authoritativeRevision = asNonNegativeSafeInteger(proof.authoritativeRevision)
  if (authoritativeRevision === null || authoritativeRevision < input.minRevision) {
    return 'Presented TypeGPU frame proof is older than the requested verification revision.'
  }
  const visibleRenderRevision = asNonNegativeSafeInteger(proof.visibleRenderRevision)
  if (visibleRenderRevision === null) {
    return 'Presented TypeGPU frame proof did not include a visible render revision.'
  }
  const tileSceneRevision = asNonNegativeSafeInteger(proof.tileSceneRevision)
  if (tileSceneRevision !== null && tileSceneRevision !== visibleRenderRevision) {
    return 'Presented TypeGPU frame proof revision does not match the current tile scene revision.'
  }
  const projectedRevision = asNonNegativeSafeInteger(proof.projectedRevision)
  if (projectedRevision !== null && projectedRevision !== visibleRenderRevision) {
    return 'Presented TypeGPU frame proof revision does not match the projected viewport revision.'
  }
  return null
}

function isSurfaceProofStale(input: {
  readonly minRevision?: number | null | undefined
  readonly proof: WorkbookAgentRenderedSurfaceProof | null
}): boolean {
  if (typeof input.minRevision !== 'number') {
    return false
  }
  const authoritativeRevision = asNonNegativeSafeInteger(input.proof?.authoritativeRevision)
  return authoritativeRevision === null || authoritativeRevision < input.minRevision
}

export function selectWorkbookRenderedReadback(input: {
  readonly renderedContext: WorkbookAgentRenderedContext | null | undefined
  readonly requestedRange: CellRangeRef
  readonly authoritativeRows?: readonly (readonly unknown[])[]
  readonly minRevision?: number | null
  readonly nextChunk?: WorkbookAgentRangeChunk | null
}): WorkbookRenderedReadbackProof {
  const requestedRange = toWorkbookAgentRangeRef(input.requestedRange)
  const renderedContext = input.renderedContext ?? null
  const selectedRangeCandidate = pickRenderedRange(renderedContext, requestedRange)
  const selectedRange = selectedRangeCandidate?.range ?? null
  const capturedBatchId = asNonNegativeSafeInteger(renderedContext?.batchId)
  const capturedRevision = renderedCaptureRevision(renderedContext)
  const surfaceProof = renderedSurfaceProof(renderedContext)
  const captureStale =
    selectedRange === null || capturedRevision === null || (typeof input.minRevision === 'number' && capturedRevision < input.minRevision)
  const surfaceProofIncompleteReason = buildSurfaceProofIncompleteReason({
    minRevision: input.minRevision,
    proof: surfaceProof,
  })
  const surfaceProofStale = isSurfaceProofStale({
    minRevision: input.minRevision,
    proof: surfaceProof,
  })
  const stale = captureStale || surfaceProofStale
  const extracted = selectedRange
    ? buildExtractedRenderedRange({
        renderedRange: selectedRange,
        requestedRange,
      })
    : {
        range: null,
        missingCells: enumerateWorkbookAgentRangeAddresses(requestedRange, 50).map((address) => `${requestedRange.sheetName}!${address}`),
      }
  const mismatches =
    extracted.range && input.authoritativeRows
      ? collectRenderedMismatches({
          sheetName: requestedRange.sheetName,
          authoritativeRows: input.authoritativeRows,
          renderedRange: extracted.range,
        })
      : []
  const proofTruncated = input.nextChunk != null || extracted.range?.truncated === true
  const sourceTruncated = selectedRange?.truncated === true
  const incompleteReason = buildIncompleteReason({
    hasRenderedContext: renderedContext !== null,
    selectedRange,
    captureStale,
    surfaceProofIncompleteReason,
    missingCells: extracted.missingCells,
    truncated: proofTruncated,
    mismatches,
  })
  const matched = incompleteReason === null ? mismatches.length === 0 : mismatches.length > 0 ? false : null
  return {
    requested: true,
    requestedRange,
    available: selectedRange !== null && extracted.range !== null,
    matched,
    stale,
    capturedRange: extracted.range?.range ?? null,
    sourceKind: selectedRangeCandidate?.sourceKind ?? null,
    sourceRange: selectedRange?.range ?? null,
    capturedAtUnixMs: renderedContext?.capturedAtUnixMs ?? null,
    capturedRevision,
    capturedBatchId,
    surfaceProof,
    surfaceProofMatched:
      typeof input.minRevision === 'number' ? (surfaceProofIncompleteReason === null ? true : surfaceProof ? false : null) : null,
    surfaceProofIncompleteReason,
    truncated: proofTruncated,
    sourceTruncated,
    missingCells: extracted.missingCells,
    mismatches,
    incompleteReason,
    nextChunk: input.nextChunk ?? null,
    range: extracted.range,
  }
}

export function emptyWorkbookRenderedReadbackProof(input: {
  readonly requested: boolean
  readonly requestedRange?: CellRangeRef | null
  readonly reason: string
}): WorkbookRenderedReadbackProof {
  return {
    requested: input.requested,
    requestedRange: input.requestedRange ? toWorkbookAgentRangeRef(input.requestedRange) : null,
    available: false,
    matched: null,
    stale: true,
    capturedRange: null,
    sourceKind: null,
    sourceRange: null,
    capturedAtUnixMs: null,
    capturedRevision: null,
    capturedBatchId: null,
    surfaceProof: null,
    surfaceProofMatched: null,
    surfaceProofIncompleteReason: input.reason,
    truncated: false,
    sourceTruncated: false,
    missingCells: input.requestedRange
      ? enumerateWorkbookAgentRangeAddresses(input.requestedRange, 50).map((address) => `${input.requestedRange!.sheetName}!${address}`)
      : [],
    mismatches: [],
    incompleteReason: input.reason,
    nextChunk: null,
    range: null,
  }
}
