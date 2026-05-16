import type { WorkbookAgentStreamEvent, WorkbookAgentThreadSnapshot, WorkbookAgentTimelineEntry } from '@bilig/contracts'

function createTextEntryFromDelta(event: Extract<WorkbookAgentStreamEvent, { type: 'entryTextDelta' }>) {
  return {
    id: event.itemId,
    kind: event.entryKind,
    turnId: event.turnId,
    text: event.delta,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [],
  } satisfies WorkbookAgentTimelineEntry
}

export function updateSnapshotFromTextDelta(
  snapshot: WorkbookAgentThreadSnapshot | null,
  event: Extract<WorkbookAgentStreamEvent, { type: 'entryTextDelta' }>,
): WorkbookAgentThreadSnapshot | null {
  if (!snapshot) {
    return snapshot
  }
  let matched = false
  return {
    ...snapshot,
    entries: (() => {
      const nextEntries = snapshot.entries.map((entry) => {
        if (entry.id !== event.itemId) {
          return entry
        }
        matched = true
        return {
          ...entry,
          kind: event.entryKind,
          turnId: event.turnId,
          text: `${entry.text ?? ''}${event.delta}`,
        }
      })
      return matched ? nextEntries : [...nextEntries, createTextEntryFromDelta(event)]
    })(),
  }
}

function createToolEntryFromOutputDelta(event: Extract<WorkbookAgentStreamEvent, { type: 'entryToolOutputDelta' }>) {
  return {
    id: event.itemId,
    kind: 'tool',
    turnId: event.turnId,
    text: null,
    phase: null,
    toolName: 'command_execution',
    toolStatus: 'inProgress',
    argumentsText: null,
    outputText: event.delta,
    success: null,
    citations: [],
  } satisfies WorkbookAgentTimelineEntry
}

export function updateSnapshotFromToolOutputDelta(
  snapshot: WorkbookAgentThreadSnapshot | null,
  event: Extract<WorkbookAgentStreamEvent, { type: 'entryToolOutputDelta' }>,
): WorkbookAgentThreadSnapshot | null {
  if (!snapshot) {
    return snapshot
  }
  let matched = false
  return {
    ...snapshot,
    entries: (() => {
      const nextEntries = snapshot.entries.map((entry) => {
        if (entry.id !== event.itemId) {
          return entry
        }
        matched = true
        return {
          ...entry,
          kind: 'tool',
          turnId: event.turnId,
          toolName: entry.toolName ?? 'command_execution',
          toolStatus: entry.toolStatus ?? 'inProgress',
          outputText: `${entry.outputText ?? ''}${event.delta}`,
        } satisfies WorkbookAgentTimelineEntry
      })
      return matched ? nextEntries : [...nextEntries, createToolEntryFromOutputDelta(event)]
    })(),
  }
}
