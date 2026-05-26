import {
  appendWorkbookAgentCommandToBundle,
  toWorkbookAgentCommandBundle,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type WorkbookAgentExecutionRecord,
} from '@bilig/agent-api'
import type { WorkbookAgentThreadSnapshot } from '@bilig/contracts'
import type { CellRangeRef } from '@bilig/protocol'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { handleWorkbookAgentToolCall, type WorkbookAgentStartWorkflowRequest } from './workbook-agent-tools.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import { cloneUiContext, type WorkbookAgentThreadState, toContextRef } from './workbook-agent-service-shared.js'
import { inspectWorkbookRange, normalizeWorkbookAgentUiContext } from './workbook-agent-inspection.js'
import { toWorkbookAgentRangeRef } from './workbook-agent-range-chunks.js'
import { selectWorkbookRenderedReadback } from './workbook-agent-rendered-readback.js'
import { shouldWaitForRenderedTool, waitForWorkbookAgentRenderedContext } from './workbook-agent-rendered-context-wait.js'
import { assertWorkbookAgentToolCallOwnsTurn } from './workbook-agent-turn-lifecycle.js'

export function selectWorkbookAgentRenderedVerificationRanges(
  bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
): readonly CellRangeRef[] {
  return bundle.affectedRanges.filter((range) => range.role === 'target').map((range) => toWorkbookAgentRangeRef(range))
}

export async function workbookAgentRenderedVerificationRangesMatch(input: {
  readonly zeroSyncService: ZeroSyncService
  readonly documentId: string
  readonly latestContext: WorkbookAgentThreadState['durable']['context']
  readonly bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>
  readonly minRevision: number
}): Promise<boolean> {
  const targetRanges = selectWorkbookAgentRenderedVerificationRanges(input.bundle)
  if (targetRanges.length === 0) {
    return true
  }
  return await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) => {
    const normalizedContext = normalizeWorkbookAgentUiContext(runtime, input.latestContext)
    return targetRanges.every((targetRange) => {
      const authoritativeReadback = inspectWorkbookRange(runtime, targetRange)
      const authoritativeRows = authoritativeReadback.rows.filter(Array.isArray) as readonly (readonly unknown[])[]
      const renderedReadback = selectWorkbookRenderedReadback({
        renderedContext: normalizedContext?.rendered,
        requestedRange: targetRange,
        authoritativeRows,
        minRevision: input.minRevision,
      })
      return renderedReadback.matched === true
    })
  })
}

export function createWorkbookAgentDynamicToolHandler(input: {
  zeroSyncService: ZeroSyncService
  now: () => number
  getSessionByThreadId: (threadId: string) => WorkbookAgentThreadState
  resolveTurnActorUserId: (sessionState: WorkbookAgentThreadState, turnId: string) => string
  resolveTurnContext: (sessionState: WorkbookAgentThreadState, turnId: string) => WorkbookAgentThreadState['durable']['context']
  stageReviewBundle: (
    sessionState: WorkbookAgentThreadState,
    turnId: string,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => void
  shouldApplyToolBundleImmediately: (
    sessionState: WorkbookAgentThreadState,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => boolean
  applyToolBundleAutomatically: (input: {
    sessionState: WorkbookAgentThreadState
    actorUserId: string
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>
    assertApplyStillAuthorized?: (() => void) | null | undefined
  }) => Promise<WorkbookAgentExecutionRecord | null>
  persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
  emitSnapshot: (threadId: string) => void
  startWorkflow: (input: {
    documentId: string
    threadId: string
    expectedActiveTurnId?: string
    session: SessionIdentity
    body: WorkbookAgentStartWorkflowRequest & {
      context?: WorkbookAgentThreadState['durable']['context']
    }
  }) => Promise<WorkbookAgentThreadSnapshot>
}): (request: CodexDynamicToolCallRequest) => Promise<CodexDynamicToolCallResult> {
  return async (request: CodexDynamicToolCallRequest) => {
    const sessionState = input.getSessionByThreadId(request.threadId)
    const assertRequestTurnOwnsSession = (): void => {
      assertWorkbookAgentToolCallOwnsTurn(sessionState, request.turnId)
    }
    assertRequestTurnOwnsSession()
    const requestActorUserId = input.resolveTurnActorUserId(sessionState, request.turnId)
    let requestContext: WorkbookAgentThreadState['durable']['context'] = null

    const refreshRequestContext = async (): Promise<WorkbookAgentThreadState['durable']['context']> => {
      assertRequestTurnOwnsSession()
      const rawRequestContext = input.resolveTurnContext(sessionState, request.turnId)
      const normalizedContext = await input.zeroSyncService.inspectWorkbook(sessionState.documentId, (runtime) =>
        normalizeWorkbookAgentUiContext(runtime, rawRequestContext),
      )
      assertRequestTurnOwnsSession()
      requestContext = normalizedContext
      if (JSON.stringify(rawRequestContext) !== JSON.stringify(normalizedContext)) {
        sessionState.durable.context = cloneUiContext(normalizedContext)
        sessionState.live.turnContextByTurn.set(request.turnId, cloneUiContext(normalizedContext))
        await input.persistSessionState(sessionState)
        input.emitSnapshot(sessionState.threadId)
      }
      return normalizedContext
    }

    const waitForRenderedContext = async (
      minRevision: number,
      isReady?: (context: WorkbookAgentThreadState['durable']['context']) => Promise<boolean>,
    ): Promise<WorkbookAgentThreadState['durable']['context']> => {
      return await waitForWorkbookAgentRenderedContext({
        minRevision,
        refreshContext: refreshRequestContext,
        ...(isReady ? { isReady } : {}),
      })
    }

    requestContext = await refreshRequestContext()
    if (shouldWaitForRenderedTool(request.tool)) {
      const headRevision = await input.zeroSyncService.getWorkbookHeadRevision(sessionState.documentId)
      requestContext = await waitForRenderedContext(headRevision)
    }

    return handleWorkbookAgentToolCall(
      {
        documentId: sessionState.documentId,
        session: {
          userID: requestActorUserId,
          roles: ['editor'],
        },
        get uiContext() {
          return requestContext
        },
        zeroSyncService: input.zeroSyncService,
        updateUiContext: async (nextContext) => {
          assertRequestTurnOwnsSession()
          sessionState.durable.context = cloneUiContext(nextContext)
          sessionState.live.turnContextByTurn.set(request.turnId, cloneUiContext(nextContext))
          await input.persistSessionState(sessionState)
          input.emitSnapshot(sessionState.threadId)
        },
        awaitRenderedRevision: async (revision) => {
          requestContext = await waitForRenderedContext(revision)
        },
        stageCommand: async (command) => {
          assertRequestTurnOwnsSession()
          const currentReviewBundle = sessionState.durable.reviewQueueItems[0]
            ? toWorkbookAgentCommandBundle(sessionState.durable.reviewQueueItems[0])
            : null
          const previousBundle = sessionState.scope === 'private' ? null : currentReviewBundle
          const baseRevision = await input.zeroSyncService.getWorkbookHeadRevision(sessionState.documentId)
          assertRequestTurnOwnsSession()
          const bundle = appendWorkbookAgentCommandToBundle({
            previousBundle,
            documentId: sessionState.documentId,
            threadId: sessionState.threadId,
            turnId: request.turnId,
            goalText: sessionState.live.promptByTurn.get(request.turnId) ?? 'Update workbook from assistant request',
            baseRevision,
            context: toContextRef(requestContext),
            command,
            now: input.now(),
          })
          if (input.shouldApplyToolBundleImmediately(sessionState, bundle)) {
            assertRequestTurnOwnsSession()
            const executionRecord = await input.applyToolBundleAutomatically({
              sessionState,
              actorUserId: requestActorUserId,
              bundle,
              assertApplyStillAuthorized: assertRequestTurnOwnsSession,
            })
            if (executionRecord && requestContext !== null) {
              requestContext = await waitForRenderedContext(
                executionRecord.appliedRevision,
                async (latestContext) =>
                  await workbookAgentRenderedVerificationRangesMatch({
                    zeroSyncService: input.zeroSyncService,
                    documentId: sessionState.documentId,
                    latestContext,
                    bundle,
                    minRevision: executionRecord.appliedRevision,
                  }),
              )
            }
            return {
              bundle,
              executionRecord,
            }
          }
          if (sessionState.scope === 'private') {
            throw createWorkbookAgentServiceError({
              code: 'WORKBOOK_AGENT_PRIVATE_EXECUTION_BLOCKED',
              message:
                'Private workbook threads execute changes directly and do not queue review items under the current execution policy.',
              statusCode: 409,
              retryable: false,
            })
          }
          assertRequestTurnOwnsSession()
          input.stageReviewBundle(sessionState, request.turnId, bundle)
          await input.persistSessionState(sessionState)
          input.emitSnapshot(sessionState.threadId)
          return {
            bundle,
            executionRecord: null,
            disposition: 'reviewQueued',
          }
        },
        startWorkflow: async (workflowRequest: WorkbookAgentStartWorkflowRequest) => {
          assertRequestTurnOwnsSession()
          const previousRunIds = new Set(sessionState.durable.workflowRuns.map((run) => run.runId))
          const nextSnapshot = await input.startWorkflow({
            documentId: sessionState.documentId,
            threadId: sessionState.threadId,
            expectedActiveTurnId: request.turnId,
            session: {
              userID: requestActorUserId,
              roles: ['editor'],
            },
            body: {
              ...workflowRequest,
              ...(requestContext ? { context: requestContext } : {}),
            },
          })
          const nextRun =
            nextSnapshot.workflowRuns.find((run) => !previousRunIds.has(run.runId)) ??
            nextSnapshot.workflowRuns.find((run) => run.workflowTemplate === workflowRequest.workflowTemplate) ??
            null
          if (!nextRun) {
            throw new Error(`Workflow run not found after starting ${workflowRequest.workflowTemplate}`)
          }
          return nextRun
        },
      },
      request,
    )
  }
}
