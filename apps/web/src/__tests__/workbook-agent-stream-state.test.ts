import { describe, expect, it } from 'vitest'
import type { WorkbookAgentStreamEvent, WorkbookAgentThreadSnapshot, WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { updateSnapshotFromTextDelta, updateSnapshotFromToolOutputDelta } from '../workbook-agent-stream-state.js'

function createSnapshot(entries: readonly WorkbookAgentTimelineEntry[] = []): WorkbookAgentThreadSnapshot {
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    },
    entries: [...entries],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
  }
}

function createAssistantEntry(overrides: Partial<WorkbookAgentTimelineEntry> = {}): WorkbookAgentTimelineEntry {
  return {
    id: 'assistant-1',
    kind: 'assistant',
    turnId: 'turn-1',
    text: 'Before ',
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [],
    ...overrides,
  }
}

function createToolEntry(overrides: Partial<WorkbookAgentTimelineEntry> = {}): WorkbookAgentTimelineEntry {
  return {
    id: 'tool-1',
    kind: 'tool',
    turnId: 'turn-1',
    text: null,
    phase: null,
    toolName: 'read_range',
    toolStatus: 'completed',
    argumentsText: null,
    outputText: 'Before\n',
    success: true,
    citations: [],
    ...overrides,
  }
}

describe('workbook agent stream state', () => {
  it('appends text deltas to an existing timeline entry', () => {
    const event = {
      type: 'entryTextDelta',
      itemId: 'assistant-1',
      turnId: 'turn-2',
      entryKind: 'assistant',
      delta: 'after',
    } satisfies WorkbookAgentStreamEvent

    const updated = updateSnapshotFromTextDelta(createSnapshot([createAssistantEntry()]), event)

    expect(updated?.entries).toHaveLength(1)
    expect(updated?.entries[0]).toMatchObject({
      id: 'assistant-1',
      kind: 'assistant',
      turnId: 'turn-2',
      text: 'Before after',
    })
  })

  it('creates a text timeline entry when a delta arrives before the snapshot entry', () => {
    const event = {
      type: 'entryTextDelta',
      itemId: 'assistant-2',
      turnId: 'turn-1',
      entryKind: 'assistant',
      delta: 'Created from stream',
    } satisfies WorkbookAgentStreamEvent

    const updated = updateSnapshotFromTextDelta(createSnapshot(), event)

    expect(updated?.entries).toEqual([
      {
        id: 'assistant-2',
        kind: 'assistant',
        turnId: 'turn-1',
        text: 'Created from stream',
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [],
      },
    ])
  })

  it('appends tool output deltas without downgrading known tool metadata', () => {
    const event = {
      type: 'entryToolOutputDelta',
      itemId: 'tool-1',
      turnId: 'turn-2',
      delta: 'After\n',
    } satisfies WorkbookAgentStreamEvent

    const updated = updateSnapshotFromToolOutputDelta(createSnapshot([createToolEntry()]), event)

    expect(updated?.entries).toHaveLength(1)
    expect(updated?.entries[0]).toMatchObject({
      id: 'tool-1',
      kind: 'tool',
      turnId: 'turn-2',
      toolName: 'read_range',
      toolStatus: 'completed',
      outputText: 'Before\nAfter\n',
      success: true,
    })
  })

  it('creates a command execution tool entry when output arrives before the snapshot entry', () => {
    const event = {
      type: 'entryToolOutputDelta',
      itemId: 'tool-2',
      turnId: 'turn-1',
      delta: 'Running\n',
    } satisfies WorkbookAgentStreamEvent

    const updated = updateSnapshotFromToolOutputDelta(createSnapshot(), event)

    expect(updated?.entries).toEqual([
      {
        id: 'tool-2',
        kind: 'tool',
        turnId: 'turn-1',
        text: null,
        phase: null,
        toolName: 'command_execution',
        toolStatus: 'inProgress',
        argumentsText: null,
        outputText: 'Running\n',
        success: null,
        citations: [],
      },
    ])
  })

  it('leaves null snapshots unchanged', () => {
    const textEvent = {
      type: 'entryTextDelta',
      itemId: 'assistant-1',
      turnId: 'turn-1',
      entryKind: 'assistant',
      delta: 'text',
    } satisfies WorkbookAgentStreamEvent
    const toolEvent = {
      type: 'entryToolOutputDelta',
      itemId: 'tool-1',
      turnId: 'turn-1',
      delta: 'output',
    } satisfies WorkbookAgentStreamEvent

    expect(updateSnapshotFromTextDelta(null, textEvent)).toBeNull()
    expect(updateSnapshotFromToolOutputDelta(null, toolEvent)).toBeNull()
  })
})
