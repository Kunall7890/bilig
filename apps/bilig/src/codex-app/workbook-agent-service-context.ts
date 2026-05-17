import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { ValueTag } from '@bilig/protocol'
import { cloneUiContext, type WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export function stripRenderedWorkbookAgentContext(context: WorkbookAgentUiContext): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: context.selection.sheetName,
      address: context.selection.address,
      ...(context.selection.range
        ? {
            range: {
              startAddress: context.selection.range.startAddress,
              endAddress: context.selection.range.endAddress,
            },
          }
        : {}),
    },
    viewport: { ...context.viewport },
  }
}

export function applyWorkbookAgentStructuralContextHints(
  context: WorkbookAgentUiContext | null,
  commands: readonly WorkbookAgentCommand[],
): WorkbookAgentUiContext | null {
  let nextContext = cloneUiContext(context)
  let structuralContextChanged = false
  for (const command of commands) {
    if (command.kind === 'createSheet') {
      nextContext = {
        selection: {
          sheetName: command.name,
          address: 'A1',
          range: {
            startAddress: 'A1',
            endAddress: 'A1',
          },
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      }
      structuralContextChanged = true
      continue
    }
    if (!nextContext) {
      continue
    }
    if (command.kind === 'renameSheet' && nextContext.selection.sheetName === command.currentName) {
      nextContext = {
        ...stripRenderedWorkbookAgentContext(nextContext),
        selection: {
          ...nextContext.selection,
          sheetName: command.nextName,
        },
      }
      structuralContextChanged = true
      continue
    }
    if (command.kind === 'deleteSheet' && nextContext.selection.sheetName === command.name) {
      nextContext = stripRenderedWorkbookAgentContext(nextContext)
      structuralContextChanged = true
    }
  }
  return structuralContextChanged && nextContext ? stripRenderedWorkbookAgentContext(nextContext) : nextContext
}

type RenderedContext = NonNullable<WorkbookAgentUiContext['rendered']>
type RenderedRange = RenderedContext['visibleRange']

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeRenderedValueForContextKey(value: unknown): unknown {
  if (!isRecord(value) || value['tag'] !== ValueTag.String) {
    return value
  }
  return {
    tag: value['tag'],
    value: value['value'],
    stringId: 0,
  }
}

function normalizeRenderedRangeForContextKey(range: RenderedRange): RenderedRange {
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

export function areWorkbookAgentUiContextsSemanticallyEqual(
  left: WorkbookAgentUiContext | null,
  right: WorkbookAgentUiContext | null,
): boolean {
  return stringifyWorkbookAgentUiContextSemanticKey(left) === stringifyWorkbookAgentUiContextSemanticKey(right)
}

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
