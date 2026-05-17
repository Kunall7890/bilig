import { describe, expect, it } from 'vitest'
import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { ValueTag } from '@bilig/protocol'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'
import {
  areWorkbookAgentUiContextsSemanticallyEqual,
  applyWorkbookAgentStructuralContextHints,
  stripRenderedWorkbookAgentContext,
  updateWorkbookAgentDurableUiContextFromUser,
} from './workbook-agent-service-context.js'

function createContext(overrides: Partial<WorkbookAgentUiContext> = {}): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: 'Revenue',
      address: 'B2',
      range: {
        startAddress: 'B2',
        endAddress: 'C3',
      },
    },
    viewport: {
      rowStart: 0,
      rowEnd: 20,
      colStart: 0,
      colEnd: 10,
    },
    rendered: {
      capturedAtUnixMs: 100,
      capturedRevision: 3,
      batchId: 1,
      selection: null,
      visibleRange: null,
    },
    ...overrides,
  }
}

function createThreadState(context: WorkbookAgentUiContext | null): WorkbookAgentThreadState {
  return {
    documentId: 'doc-1',
    userId: 'alex@example.com',
    storageActorUserId: 'alex@example.com',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    threadId: 'thr-1',
    durable: {
      context,
      entries: [],
      reviewQueueItems: [],
      executionRecords: [],
      workflowRuns: [],
    },
    live: {
      activeTurnId: 'turn-1',
      status: 'idle',
      lastError: null,
      stagedPrivateBundleByTurn: new Map(),
      optimisticUserEntryIdByTurn: new Map(),
      promptByTurn: new Map(),
      turnActorUserIdByTurn: new Map([['turn-1', 'alex@example.com']]),
      turnContextByTurn: new Map(),
      lastAccessedAt: 0,
    },
  }
}

function createRenderedCell(value: string, stringId: number) {
  return {
    address: 'B2',
    input: value,
    value: {
      tag: ValueTag.String,
      value,
      stringId,
    },
    formula: null,
    displayFormat: null,
    styleId: null,
    numberFormatId: null,
    style: null,
  }
}

function createRenderedContext(value: string, overrides: Partial<NonNullable<WorkbookAgentUiContext['rendered']>> = {}) {
  return createContext({
    rendered: {
      capturedAtUnixMs: 100,
      capturedRevision: 3,
      batchId: 1,
      selection: null,
      visibleRange: {
        range: {
          sheetName: 'Revenue',
          startAddress: 'B2',
          endAddress: 'B2',
        },
        rowCount: 1,
        columnCount: 1,
        cellCount: 1,
        truncated: false,
        rows: [[createRenderedCell(value, 1)]],
      },
      ...overrides,
    },
  })
}

describe('workbook agent service context helpers', () => {
  it('strips rendered context while preserving selection and viewport', () => {
    expect(stripRenderedWorkbookAgentContext(createContext())).toEqual({
      selection: {
        sheetName: 'Revenue',
        address: 'B2',
        range: {
          startAddress: 'B2',
          endAddress: 'C3',
        },
      },
      viewport: {
        rowStart: 0,
        rowEnd: 20,
        colStart: 0,
        colEnd: 10,
      },
    })
  })

  it('applies structural command hints to workbook context', () => {
    const renamed = applyWorkbookAgentStructuralContextHints(createContext(), [
      {
        kind: 'renameSheet',
        currentName: 'Revenue',
        nextName: 'Forecast',
      } satisfies Extract<WorkbookAgentCommand, { kind: 'renameSheet' }>,
    ])
    expect(renamed?.selection.sheetName).toBe('Forecast')
    expect(renamed).not.toHaveProperty('rendered')

    const created = applyWorkbookAgentStructuralContextHints(createContext(), [
      {
        kind: 'createSheet',
        name: 'Ops',
      } satisfies Extract<WorkbookAgentCommand, { kind: 'createSheet' }>,
    ])
    expect(created).toEqual({
      selection: {
        sheetName: 'Ops',
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
    })
  })

  it('updates durable and turn context only for the active turn owner', () => {
    const sessionState = createThreadState(null)
    const nextContext = createContext({ rendered: undefined })

    expect(
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: nextContext,
        userId: 'alex@example.com',
      }),
    ).toBe(true)

    expect(sessionState.durable.context).toEqual(nextContext)
    expect(sessionState.live.turnContextByTurn.get('turn-1')).toEqual(nextContext)

    sessionState.live.turnActorUserIdByTurn.set('turn-1', 'casey@example.com')
    const caseyContext = createContext({ selection: { sheetName: 'Sheet2', address: 'A1' } as WorkbookAgentUiContext['selection'] })

    expect(
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: caseyContext,
        userId: 'alex@example.com',
      }),
    ).toBe(true)

    expect(sessionState.durable.context).toEqual(caseyContext)
    expect(sessionState.live.turnContextByTurn.get('turn-1')).toEqual(nextContext)
  })

  it('does not update durable context for rendered proof metadata churn', () => {
    const durableContext = createRenderedContext('stable value')
    const sessionState = createThreadState(durableContext)
    sessionState.live.turnContextByTurn.set('turn-1', durableContext)
    const nextContext = createRenderedContext('stable value', {
      capturedAtUnixMs: 900,
      capturedRevision: 12,
      batchId: 45,
      visibleRange: {
        range: {
          sheetName: 'Revenue',
          startAddress: 'B2',
          endAddress: 'B2',
        },
        rowCount: 1,
        columnCount: 1,
        cellCount: 1,
        truncated: false,
        rows: [[createRenderedCell('stable value', 99)]],
      },
    })

    expect(areWorkbookAgentUiContextsSemanticallyEqual(durableContext, nextContext)).toBe(true)
    expect(
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: nextContext,
        userId: 'alex@example.com',
      }),
    ).toBe(false)
    expect(sessionState.durable.context).toBe(durableContext)
    expect(sessionState.live.turnContextByTurn.get('turn-1')).toBe(durableContext)
  })

  it('updates durable context when rendered visible cell content changes', () => {
    const durableContext = createRenderedContext('before')
    const sessionState = createThreadState(durableContext)
    sessionState.live.turnContextByTurn.set('turn-1', durableContext)
    const nextContext = createRenderedContext('after', {
      capturedAtUnixMs: 900,
      capturedRevision: 12,
      batchId: 45,
    })

    expect(areWorkbookAgentUiContextsSemanticallyEqual(durableContext, nextContext)).toBe(false)
    expect(
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: nextContext,
        userId: 'alex@example.com',
      }),
    ).toBe(true)
    expect(sessionState.durable.context).toEqual(nextContext)
    expect(sessionState.live.turnContextByTurn.get('turn-1')).toEqual(nextContext)
  })

  it('hydrates missing active-turn context even when durable context is already current', () => {
    const durableContext = createContext({ rendered: undefined })
    const sessionState = createThreadState(durableContext)

    expect(
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: durableContext,
        userId: 'alex@example.com',
      }),
    ).toBe(true)
    expect(sessionState.durable.context).toBe(durableContext)
    expect(sessionState.live.turnContextByTurn.get('turn-1')).toEqual(durableContext)

    expect(
      updateWorkbookAgentDurableUiContextFromUser({
        sessionState,
        context: durableContext,
        userId: 'alex@example.com',
      }),
    ).toBe(false)
  })
})
