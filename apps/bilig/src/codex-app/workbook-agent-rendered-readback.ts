import { formatAddress } from '@bilig/formula'
import { formatErrorCode, ValueTag, type CellRangeRef } from '@bilig/protocol'
import type {
  WorkbookAgentRenderedCell,
  WorkbookAgentRenderedContext,
  WorkbookAgentRenderedRange,
  WorkbookAgentRenderedVisibleSceneProof,
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
  readonly visibleSceneProof: WorkbookRenderedVisibleSceneProofStatus
  readonly capturedRange: CellRangeRef | null
  readonly sourceRange: CellRangeRef | null
  readonly capturedAtUnixMs: number | null
  readonly capturedRevision: number | null
  readonly capturedBatchId: number | null
  readonly truncated: boolean
  readonly sourceTruncated: boolean
  readonly missingCells: readonly string[]
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
  readonly nextChunk: WorkbookAgentRangeChunk | null
  readonly range: WorkbookAgentRenderedRange | null
}

export interface WorkbookRenderedVisibleSceneProofStatus {
  readonly requested: boolean
  readonly available: boolean
  readonly matched: boolean | null
  readonly rendererMode: string | null
  readonly frameProofStatus: string | null
  readonly frameProofSignature: string | null
  readonly presentedFrameProofSignature: string | null
  readonly currentSceneOwnershipSignature: string | null
  readonly presentedSceneOwnershipSignature: string | null
  readonly gridAuthoritativeRevision: string | null
  readonly typeGpuAuthoritativeRevision: string | null
  readonly visibleAuthoritativeRevision: string | null
  readonly tileSceneRevision: string | null
  readonly visibleRenderRevision: string | null
  readonly hasPresentedFrame: boolean | null
  readonly hasPresentedVisibleFrame: boolean | null
  readonly frameProofMatchesPresentedFrame: boolean | null
  readonly visibleSceneOwnershipMatchesPresentedFrame: boolean | null
  readonly visibleAuthoritativeRevisionMatchesGrid: boolean | null
  readonly visibleRenderRevisionMatchesTileScene: boolean | null
  readonly incompleteReason: string | null
  readonly invalidReasons: readonly string[]
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

function visibleSceneProofInvalidReasons(proof: WorkbookAgentRenderedVisibleSceneProof | null | undefined): string[] {
  const invalidReasons: string[] = []
  if (!proof) {
    return ['No TypeGPU visible-scene proof was attached to the rendered workbook context.']
  }
  if (proof.rendererMode !== 'typegpu-v3') {
    invalidReasons.push(`Renderer mode is ${proof.rendererMode ?? 'missing'}.`)
  }
  if (proof.frameProofStatus !== 'presented') {
    invalidReasons.push(`Frame proof status is ${proof.frameProofStatus ?? 'missing'}.`)
  }
  if (!proof.hasPresentedFrame) {
    invalidReasons.push('Current frame proof has not been presented.')
  }
  if (!proof.hasPresentedVisibleFrame) {
    invalidReasons.push('Current visible scene has not been presented.')
  }
  if (!proof.frameProofMatchesPresentedFrame) {
    invalidReasons.push('Presented frame proof signature does not match the current frame.')
  }
  if (!proof.visibleSceneOwnershipMatchesPresentedFrame) {
    invalidReasons.push('Presented visible-scene ownership does not match the current scene.')
  }
  if (!proof.visibleAuthoritativeRevisionMatchesGrid) {
    invalidReasons.push('Visible authoritative revision does not match the grid authoritative revision.')
  }
  if (!proof.visibleRenderRevisionMatchesTileScene) {
    invalidReasons.push('Visible render revision does not match the tile scene revision.')
  }
  return invalidReasons
}

function buildVisibleSceneProofStatus(
  proof: WorkbookAgentRenderedVisibleSceneProof | null | undefined,
): WorkbookRenderedVisibleSceneProofStatus {
  const invalidReasons = visibleSceneProofInvalidReasons(proof)
  return {
    requested: true,
    available: proof != null,
    matched: proof == null ? null : invalidReasons.length === 0,
    rendererMode: proof?.rendererMode ?? null,
    frameProofStatus: proof?.frameProofStatus ?? null,
    frameProofSignature: proof?.frameProofSignature ?? null,
    presentedFrameProofSignature: proof?.presentedFrameProofSignature ?? null,
    currentSceneOwnershipSignature: proof?.currentSceneOwnershipSignature ?? null,
    presentedSceneOwnershipSignature: proof?.presentedSceneOwnershipSignature ?? null,
    gridAuthoritativeRevision: proof?.gridAuthoritativeRevision ?? null,
    typeGpuAuthoritativeRevision: proof?.typeGpuAuthoritativeRevision ?? null,
    visibleAuthoritativeRevision: proof?.visibleAuthoritativeRevision ?? null,
    tileSceneRevision: proof?.tileSceneRevision ?? null,
    visibleRenderRevision: proof?.visibleRenderRevision ?? null,
    hasPresentedFrame: proof?.hasPresentedFrame ?? null,
    hasPresentedVisibleFrame: proof?.hasPresentedVisibleFrame ?? null,
    frameProofMatchesPresentedFrame: proof?.frameProofMatchesPresentedFrame ?? null,
    visibleSceneOwnershipMatchesPresentedFrame: proof?.visibleSceneOwnershipMatchesPresentedFrame ?? null,
    visibleAuthoritativeRevisionMatchesGrid: proof?.visibleAuthoritativeRevisionMatchesGrid ?? null,
    visibleRenderRevisionMatchesTileScene: proof?.visibleRenderRevisionMatchesTileScene ?? null,
    incompleteReason: invalidReasons.length === 0 ? null : 'Rendered TypeGPU visible-scene proof is incomplete or stale.',
    invalidReasons,
  }
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

function renderedCandidates(context: WorkbookAgentRenderedContext | null | undefined): WorkbookAgentRenderedRange[] {
  if (!context) {
    return []
  }
  return [context.selection, context.visibleRange].filter((entry): entry is WorkbookAgentRenderedRange => entry !== null)
}

function pickRenderedRange(
  context: WorkbookAgentRenderedContext | null | undefined,
  requestedRange: CellRangeRef,
): WorkbookAgentRenderedRange | null {
  const candidates = renderedCandidates(context)
  return (
    candidates.find((entry) => {
      const source = toWorkbookAgentRangeRef(entry.range)
      const target = toWorkbookAgentRangeRef(requestedRange)
      return source.sheetName === target.sheetName && source.startAddress === target.startAddress && source.endAddress === target.endAddress
    }) ??
    candidates.find((entry) => rangeContains(entry.range, requestedRange)) ??
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
  readonly visibleSceneProof: WorkbookRenderedVisibleSceneProofStatus
  readonly selectedRange: WorkbookAgentRenderedRange | null
  readonly revisionStale: boolean
  readonly sceneProofStale: boolean
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
  if (input.revisionStale) {
    return 'Rendered capture is older than the requested verification revision.'
  }
  if (input.sceneProofStale) {
    return input.visibleSceneProof.incompleteReason ?? 'Rendered TypeGPU visible-scene proof is incomplete or stale.'
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

export function selectWorkbookRenderedReadback(input: {
  readonly renderedContext: WorkbookAgentRenderedContext | null | undefined
  readonly requestedRange: CellRangeRef
  readonly authoritativeRows?: readonly (readonly unknown[])[]
  readonly minRevision?: number | null
  readonly nextChunk?: WorkbookAgentRangeChunk | null
}): WorkbookRenderedReadbackProof {
  const requestedRange = toWorkbookAgentRangeRef(input.requestedRange)
  const renderedContext = input.renderedContext ?? null
  const selectedRange = pickRenderedRange(renderedContext, requestedRange)
  const capturedBatchId = asNonNegativeSafeInteger(renderedContext?.batchId)
  const capturedRevision = renderedCaptureRevision(renderedContext)
  const visibleSceneProof = buildVisibleSceneProofStatus(renderedContext?.visibleSceneProof)
  const revisionStale =
    selectedRange === null || capturedRevision === null || (typeof input.minRevision === 'number' && capturedRevision < input.minRevision)
  const sceneProofStale = visibleSceneProof.matched !== true
  const stale = revisionStale || sceneProofStale
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
    visibleSceneProof,
    selectedRange,
    revisionStale,
    sceneProofStale,
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
    visibleSceneProof,
    capturedRange: extracted.range?.range ?? null,
    sourceRange: selectedRange?.range ?? null,
    capturedAtUnixMs: renderedContext?.capturedAtUnixMs ?? null,
    capturedRevision,
    capturedBatchId,
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
    visibleSceneProof: buildVisibleSceneProofStatus(null),
    capturedRange: null,
    sourceRange: null,
    capturedAtUnixMs: null,
    capturedRevision: null,
    capturedBatchId: null,
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
