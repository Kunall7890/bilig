import type { WorkbookAgentRenderedRange, WorkbookAgentUiContext } from './index.js'

const WORKBOOK_AGENT_STRING_VALUE_TAG = 3

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeRenderedValueForContextKey(value: unknown): unknown {
  if (!isRecord(value) || value['tag'] !== WORKBOOK_AGENT_STRING_VALUE_TAG) {
    return value
  }
  return {
    tag: value['tag'],
    value: value['value'],
    stringId: 0,
  }
}

function normalizeRenderedRangeForContextKey(range: WorkbookAgentRenderedRange | null): WorkbookAgentRenderedRange | null {
  if (range === null) {
    return null
  }
  return {
    ...range,
    rows: range.rows.map((row) =>
      row.map((cell) => ({
        ...cell,
        value: normalizeRenderedValueForContextKey(cell.value),
      })),
    ),
  }
}

export function stringifyWorkbookAgentUiContextSemanticKey(context: WorkbookAgentUiContext | null): string {
  if (context === null) {
    return 'null'
  }
  const rendered = context.rendered
  return JSON.stringify({
    selection: context.selection,
    viewport: context.viewport,
    rendered:
      rendered === undefined
        ? null
        : {
            selection: normalizeRenderedRangeForContextKey(rendered.selection),
            visibleRange: normalizeRenderedRangeForContextKey(rendered.visibleRange),
          },
  })
}

export function stringifyWorkbookAgentUiContextRenderedProofKey(context: WorkbookAgentUiContext | null): string {
  if (context === null) {
    return 'null'
  }
  const rendered = context.rendered
  return JSON.stringify({
    semantic: stringifyWorkbookAgentUiContextSemanticKey(context),
    rendered:
      rendered === undefined
        ? null
        : {
            capturedRevision: rendered.capturedRevision ?? null,
            surfaceProof:
              rendered.surfaceProof === undefined || rendered.surfaceProof === null
                ? null
                : {
                    mode: rendered.surfaceProof.mode,
                    backendStatus: rendered.surfaceProof.backendStatus,
                    frameProofStatus: rendered.surfaceProof.frameProofStatus,
                    hasPresentedFrame: rendered.surfaceProof.hasPresentedFrame,
                    hasPresentedVisibleFrame: rendered.surfaceProof.hasPresentedVisibleFrame,
                    frameProofSignature: rendered.surfaceProof.frameProofSignature,
                    presentedFrameProofSignature: rendered.surfaceProof.presentedFrameProofSignature,
                    authoritativeRevision: rendered.surfaceProof.authoritativeRevision,
                    projectedRevision: rendered.surfaceProof.projectedRevision,
                    visibleRenderRevision: rendered.surfaceProof.visibleRenderRevision,
                    tileSceneRevision: rendered.surfaceProof.tileSceneRevision,
                    currentTilePaneCount: rendered.surfaceProof.currentTilePaneCount,
                    currentHeaderPaneCount: rendered.surfaceProof.currentHeaderPaneCount,
                    presentedTilePaneCount: rendered.surfaceProof.presentedTilePaneCount,
                    presentedHeaderPaneCount: rendered.surfaceProof.presentedHeaderPaneCount,
                    surfaceWidth: rendered.surfaceProof.surfaceWidth,
                    surfaceHeight: rendered.surfaceProof.surfaceHeight,
                    surfacePixelWidth: rendered.surfaceProof.surfacePixelWidth,
                    surfacePixelHeight: rendered.surfaceProof.surfacePixelHeight,
                  },
          },
  })
}
