import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { areWorkbookAgentUiContextsSemanticallyEqual } from './workbook-agent-ui-context-semantic-key.js'
import { cloneUiContext, type WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export function updateWorkbookAgentDurableUiContextFromUser(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly context: WorkbookAgentUiContext
  readonly userId: string
}): boolean {
  const nextContext = cloneUiContext(input.context)
  const durableContextChanged = !areWorkbookAgentUiContextsSemanticallyEqual(input.sessionState.durable.context, nextContext)
  const activeTurnId = input.sessionState.live.activeTurnId
  const activeTurnActorUserId = activeTurnId ? input.sessionState.live.turnActorUserIdByTurn.get(activeTurnId) : undefined
  const canUpdateActiveTurnContext =
    activeTurnId !== null && activeTurnId !== undefined && (activeTurnActorUserId === undefined || activeTurnActorUserId === input.userId)
  const currentTurnContext = canUpdateActiveTurnContext ? (input.sessionState.live.turnContextByTurn.get(activeTurnId) ?? null) : null
  const turnContextChanged = canUpdateActiveTurnContext && !areWorkbookAgentUiContextsSemanticallyEqual(currentTurnContext, nextContext)
  if (!durableContextChanged && !turnContextChanged) {
    return false
  }
  if (durableContextChanged) {
    input.sessionState.durable.context = nextContext
  }
  if (turnContextChanged) {
    input.sessionState.live.turnContextByTurn.set(activeTurnId, cloneUiContext(input.context))
  }
  return true
}
