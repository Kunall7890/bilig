import type { CellSnapshot, CellStyleRecord, RecalcMetrics, Viewport, WorkbookMergeRangeSnapshot } from '@bilig/protocol'
import type { ViewportAxisPatch, ViewportPatch } from './viewport-patch.js'

export const WORKBOOK_VIEW_WINDOW_SCHEMA_VERSION = 1

export type WorkbookViewWindowStatus = 'ready' | 'missing-sheet' | 'identity-mismatch' | 'degraded'

export type WorkbookViewWindowRenderAckStatus = 'not-requested' | 'pending' | 'presented' | 'rejected'

export interface WorkbookViewWindowPresentedRenderAck {
  readonly batchId?: number | undefined
  readonly renderRevision?: number | undefined
  readonly proofSignature?: string | undefined
}

export interface WorkbookViewWindowRenderAckRequest {
  readonly required?: boolean | undefined
  readonly minAuthoritativeRevision?: number | undefined
  readonly presented?: WorkbookViewWindowPresentedRenderAck | null | undefined
}

export interface WorkbookViewWindowSubscription extends Viewport {
  sheetId: number
  sheetName?: string | undefined
  sheetOrdinal?: number | undefined
  renderAckRequest?: WorkbookViewWindowRenderAckRequest | undefined
}

export interface WorkbookViewWindowSheetIdentity {
  sheetId: number
  sheetOrdinal: number
  sheetName: string
}

export interface WorkbookViewWindowCell {
  row: number
  col: number
  address: string
  snapshot: CellSnapshot
  displayText: string
  copyText: string
  editorText: string
  formatId: number
  styleId: string
}

export interface WorkbookViewWindowRenderAck {
  status: WorkbookViewWindowRenderAckStatus
  batchId?: number | undefined
  renderRevision?: number | undefined
  proofSignature?: string | undefined
  reason?: string | undefined
}

export interface WorkbookViewWindow {
  schemaVersion: typeof WORKBOOK_VIEW_WINDOW_SCHEMA_VERSION
  status: WorkbookViewWindowStatus
  reason?: string | undefined
  request: WorkbookViewWindowSubscription
  sheet: WorkbookViewWindowSheetIdentity | null
  authoritativeRevision: number | null
  patchVersion: number | null
  full: boolean
  viewport: Viewport
  freezeRows: number
  freezeCols: number
  metrics: RecalcMetrics | null
  styles: CellStyleRecord[]
  cells: WorkbookViewWindowCell[]
  merges: WorkbookMergeRangeSnapshot[]
  columns: ViewportAxisPatch[]
  rows: ViewportAxisPatch[]
  renderAck: WorkbookViewWindowRenderAck
}

export function resolveWorkbookViewWindowRenderAck(input: {
  readonly request: WorkbookViewWindowSubscription
  readonly authoritativeRevision: number | null | undefined
}): WorkbookViewWindowRenderAck {
  const request = input.request.renderAckRequest
  if (request?.required !== true) {
    return {
      status: 'not-requested',
      reason: 'authoritative-window-built-before-browser-render-ack',
    }
  }

  const presented = request.presented ?? null
  if (!presented) {
    return {
      status: 'pending',
      reason: 'browser-render-ack-required-but-not-supplied',
    }
  }

  const minAuthoritativeRevision = normalizeNonNegativeInteger(request.minAuthoritativeRevision ?? input.authoritativeRevision ?? undefined)
  const renderRevision = normalizeNonNegativeInteger(presented.renderRevision)
  const batchId = normalizeNonNegativeInteger(presented.batchId)
  const proofSignature = normalizeProofSignature(presented.proofSignature)
  const ackBase = {
    ...(batchId === undefined ? {} : { batchId }),
    ...(renderRevision === undefined ? {} : { renderRevision }),
    ...(proofSignature === undefined ? {} : { proofSignature }),
  }

  if (proofSignature === undefined) {
    return {
      status: 'rejected',
      ...ackBase,
      reason: 'browser-render-ack-missing-proof-signature',
    }
  }
  if (minAuthoritativeRevision !== undefined && renderRevision === undefined) {
    return {
      status: 'rejected',
      ...ackBase,
      reason: 'browser-render-ack-missing-render-revision',
    }
  }
  if (minAuthoritativeRevision !== undefined && renderRevision !== undefined && renderRevision < minAuthoritativeRevision) {
    return {
      status: 'rejected',
      ...ackBase,
      reason: 'browser-render-ack-stale',
    }
  }

  return {
    status: 'presented',
    ...ackBase,
  }
}

function normalizeNonNegativeInteger(value: number | undefined): number | undefined {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }
  return Math.max(0, Math.trunc(value))
}

function normalizeProofSignature(value: string | undefined): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : undefined
}

export function buildWorkbookViewWindowFromViewportPatch(input: {
  readonly patch: ViewportPatch
  readonly request: WorkbookViewWindowSubscription
  readonly sheet: WorkbookViewWindowSheetIdentity
  readonly renderAck?: WorkbookViewWindowRenderAck | undefined
}): WorkbookViewWindow {
  const { patch } = input
  return {
    schemaVersion: WORKBOOK_VIEW_WINDOW_SCHEMA_VERSION,
    status: 'ready',
    request: { ...input.request },
    sheet: { ...input.sheet },
    authoritativeRevision: patch.authoritativeRevision ?? null,
    patchVersion: patch.version,
    full: patch.full,
    viewport: {
      rowStart: patch.viewport.rowStart,
      rowEnd: patch.viewport.rowEnd,
      colStart: patch.viewport.colStart,
      colEnd: patch.viewport.colEnd,
    },
    freezeRows: patch.freezeRows ?? 0,
    freezeCols: patch.freezeCols ?? 0,
    metrics: { ...patch.metrics },
    styles: patch.styles.map((style) => ({ ...style })),
    cells: patch.cells.map((cell) => ({
      row: cell.row,
      col: cell.col,
      address: cell.snapshot.address,
      snapshot: cell.snapshot,
      displayText: cell.displayText,
      copyText: cell.copyText,
      editorText: cell.editorText,
      formatId: cell.formatId,
      styleId: cell.styleId,
    })),
    merges: patch.merges?.map((range) => ({ ...range })) ?? [],
    columns: patch.columns.map(cloneViewportAxisPatch),
    rows: patch.rows.map(cloneViewportAxisPatch),
    renderAck:
      input.renderAck ??
      resolveWorkbookViewWindowRenderAck({
        request: input.request,
        authoritativeRevision: patch.authoritativeRevision ?? null,
      }),
  }
}

function cloneViewportAxisPatch(axis: ViewportAxisPatch): ViewportAxisPatch {
  return {
    hidden: axis.hidden,
    index: axis.index,
    size: axis.size,
  }
}

export function buildUnavailableWorkbookViewWindow(input: {
  readonly status: Exclude<WorkbookViewWindowStatus, 'ready'>
  readonly reason: string
  readonly request: WorkbookViewWindowSubscription
  readonly sheet?: WorkbookViewWindowSheetIdentity | null | undefined
  readonly authoritativeRevision?: number | null | undefined
  readonly renderAck?: WorkbookViewWindowRenderAck | undefined
}): WorkbookViewWindow {
  return {
    schemaVersion: WORKBOOK_VIEW_WINDOW_SCHEMA_VERSION,
    status: input.status,
    reason: input.reason,
    request: { ...input.request },
    sheet: input.sheet ? { ...input.sheet } : null,
    authoritativeRevision: input.authoritativeRevision ?? null,
    patchVersion: null,
    full: false,
    viewport: {
      rowStart: input.request.rowStart,
      rowEnd: input.request.rowEnd,
      colStart: input.request.colStart,
      colEnd: input.request.colEnd,
    },
    freezeRows: 0,
    freezeCols: 0,
    metrics: null,
    styles: [],
    cells: [],
    merges: [],
    columns: [],
    rows: [],
    renderAck: input.renderAck ?? { status: 'rejected', reason: input.reason },
  }
}
