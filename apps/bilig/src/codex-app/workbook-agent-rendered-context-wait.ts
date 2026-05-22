import { WORKBOOK_AGENT_TOOL_NAMES, normalizeWorkbookAgentToolName } from '@bilig/agent-api'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

const RENDERED_CONTEXT_WAIT_TIMEOUT_MS = 20_000
const RENDERED_CONTEXT_ATTACH_TIMEOUT_MS = 1_000
const RENDERED_CONTEXT_POLL_INTERVAL_MS = 50

type WorkbookAgentUiContext = WorkbookAgentThreadState['durable']['context']

export function hasRenderedContext(context: WorkbookAgentUiContext): boolean {
  return context?.rendered !== undefined
}

function canPollForRenderedContext(context: WorkbookAgentUiContext): boolean {
  return context !== null
}

function renderedRevision(context: WorkbookAgentUiContext): number | null {
  const capturedRevision = context?.rendered?.capturedRevision
  if (typeof capturedRevision === 'number' && Number.isSafeInteger(capturedRevision) && capturedRevision >= 0) {
    return capturedRevision
  }
  return null
}

function surfaceProofRevision(context: WorkbookAgentUiContext): number | null {
  const surfaceProof = context?.rendered?.surfaceProof
  const revision = surfaceProof?.authoritativeRevision
  const visibleRenderRevision = surfaceProof?.visibleRenderRevision
  if (
    surfaceProof?.mode === 'typegpu-v3' &&
    surfaceProof.backendStatus === 'ready' &&
    surfaceProof.frameProofStatus === 'presented' &&
    surfaceProof.hasPresentedFrame &&
    surfaceProof.hasPresentedVisibleFrame &&
    surfaceProof.frameProofSignature.trim().length > 0 &&
    surfaceProof.presentedFrameProofSignature.trim().length > 0 &&
    surfaceProof.frameProofSignature === surfaceProof.presentedFrameProofSignature &&
    surfaceProof.currentTilePaneCount > 0 &&
    surfaceProof.currentHeaderPaneCount > 0 &&
    surfaceProof.presentedTilePaneCount > 0 &&
    surfaceProof.presentedHeaderPaneCount > 0 &&
    surfaceProof.surfaceWidth > 0 &&
    surfaceProof.surfaceHeight > 0 &&
    typeof visibleRenderRevision === 'number' &&
    Number.isSafeInteger(visibleRenderRevision) &&
    visibleRenderRevision >= 0 &&
    visibleRenderRevision === surfaceProof.tileSceneRevision &&
    visibleRenderRevision === surfaceProof.projectedRevision &&
    typeof revision === 'number' &&
    Number.isSafeInteger(revision) &&
    revision >= 0
  ) {
    return Math.min(revision, visibleRenderRevision)
  }
  return null
}

export function hasRenderedContextAtRevision(context: WorkbookAgentUiContext, minRevision: number): boolean {
  const revision = renderedRevision(context)
  const presentedRevision = surfaceProofRevision(context)
  return revision !== null && revision >= minRevision && presentedRevision !== null && presentedRevision >= minRevision
}

export function shouldWaitForRenderedTool(toolName: string): boolean {
  const normalizedTool = normalizeWorkbookAgentToolName(toolName)
  return (
    normalizedTool === WORKBOOK_AGENT_TOOL_NAMES.readRenderedSelection ||
    normalizedTool === WORKBOOK_AGENT_TOOL_NAMES.readRenderedRange ||
    normalizedTool === WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify
  )
}

export async function waitForWorkbookAgentRenderedContext(input: {
  readonly minRevision: number
  readonly refreshContext: () => Promise<WorkbookAgentUiContext>
  readonly isReady?: (context: WorkbookAgentUiContext) => Promise<boolean>
  readonly delay?: (ms: number) => Promise<void>
  readonly now?: () => number
  readonly timeoutMs?: number
  readonly attachTimeoutMs?: number
  readonly pollIntervalMs?: number
}): Promise<WorkbookAgentUiContext> {
  const delay =
    input.delay ??
    ((ms: number) =>
      new Promise<void>((resolve) => {
        setTimeout(resolve, ms)
      }))
  const now = input.now ?? Date.now
  const timeoutMs = input.timeoutMs ?? RENDERED_CONTEXT_WAIT_TIMEOUT_MS
  const attachTimeoutMs = input.attachTimeoutMs ?? RENDERED_CONTEXT_ATTACH_TIMEOUT_MS
  const pollIntervalMs = input.pollIntervalMs ?? RENDERED_CONTEXT_POLL_INTERVAL_MS
  const startedAt = now()
  const deadline = startedAt + timeoutMs
  const attachDeadline = startedAt + attachTimeoutMs

  const pollRenderedContext = async (latestContext: WorkbookAgentUiContext): Promise<WorkbookAgentUiContext> => {
    if (hasRenderedContextAtRevision(latestContext, input.minRevision) && (!input.isReady || (await input.isReady(latestContext)))) {
      return latestContext
    }
    const latestNow = now()
    if (
      !canPollForRenderedContext(latestContext) ||
      latestNow >= deadline ||
      (!hasRenderedContext(latestContext) && latestNow >= attachDeadline)
    ) {
      return latestContext
    }
    await delay(pollIntervalMs)
    return await pollRenderedContext(await input.refreshContext())
  }

  return await pollRenderedContext(await input.refreshContext())
}
