import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { areWorkbookAgentUiContextsSemanticallyEqual } from './workbook-agent-ui-context-semantic-key.js'
import { cloneUiContext, type WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import { canUpdateWorkbookAgentActiveTurnContext } from './workbook-agent-service-session-policy.js'

export function updateWorkbookAgentDurableUiContextFromUser(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly context: WorkbookAgentUiContext
  readonly userId: string
}): boolean {
  const nextContext = cloneUiContext(input.context)
  const durableContextChanged = !areWorkbookAgentUiContextsSemanticallyEqual(input.sessionState.durable.context, nextContext)
  const activeTurnId = input.sessionState.live.activeTurnId
  const canUpdateActiveTurnContext = canUpdateWorkbookAgentActiveTurnContext({
    sessionState: input.sessionState,
    userId: input.userId,
  })
  const currentTurnContext =
    activeTurnId !== null && canUpdateActiveTurnContext ? (input.sessionState.live.turnContextByTurn.get(activeTurnId) ?? null) : null
  const turnContextChanged = canUpdateActiveTurnContext && !areWorkbookAgentUiContextsSemanticallyEqual(currentTurnContext, nextContext)
  if (!durableContextChanged && !turnContextChanged) {
    return false
  }
  if (durableContextChanged) {
    input.sessionState.durable.context = nextContext
  }
  if (activeTurnId !== null && turnContextChanged) {
    input.sessionState.live.turnContextByTurn.set(activeTurnId, cloneUiContext(input.context))
  }
  return true
}
