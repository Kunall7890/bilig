import { describe, expect, it } from 'vitest'
import type { WorkbookAgentSharedReviewRecommendation } from '@bilig/agent-api'
import type { WorkbookAgentThreadSummary, WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { getWorkbookAgentComposerState, shouldSubmitWorkbookAgentComposerKey } from '../workbook-agent-panel-composer.js'
import { getReviewItemCardState, renderPreviewChangeKind } from '../workbook-agent-panel-review-card.js'
import { getVisibleWorkbookAgentThreadSummaries } from '../workbook-agent-panel-thread-list.js'
import {
  getVisibleWorkbookAgentTimelineEntries,
  getWorkbookAgentProgressAnchorIndex,
  isHiddenTimelineEntry,
  summarizeDisclosureText,
  summarizeTimelineCitations,
} from '../workbook-agent-panel-timeline.js'

function timelineEntry(overrides: Partial<WorkbookAgentTimelineEntry>): WorkbookAgentTimelineEntry {
  return {
    id: 'entry-1',
    kind: 'assistant',
    turnId: 'turn-1',
    text: 'message',
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

function threadSummary(overrides: Partial<WorkbookAgentThreadSummary>): WorkbookAgentThreadSummary {
  return {
    threadId: 'thread-1',
    scope: 'private',
    ownerUserId: 'alex@example.com',
    updatedAtUnixMs: 1,
    entryCount: 1,
    reviewQueueItemCount: 0,
    latestEntryText: null,
    ...overrides,
  }
}

function recommendation(decision: 'approved' | 'rejected', userId: string): WorkbookAgentSharedReviewRecommendation {
  return {
    userId,
    decision,
    decidedAtUnixMs: 1,
  }
}

describe('workbook agent panel models', () => {
  it('keeps hidden execution noise out of the visible timeline', () => {
    const hidden = timelineEntry({
      id: 'hidden',
      kind: 'system',
      text: 'Applied automatically selected workbook change set at revision r42.',
    })
    const visible = timelineEntry({ id: 'visible', kind: 'assistant', text: 'Done', turnId: 'turn-2' })

    expect(isHiddenTimelineEntry(hidden)).toBe(true)
    expect(
      getVisibleWorkbookAgentTimelineEntries({
        optimisticEntries: [hidden],
        snapshotEntries: [visible],
      }),
    ).toEqual([visible])
  })

  it('anchors active progress after the latest visible active-turn entry', () => {
    const visibleEntries = [
      timelineEntry({ id: 'old', turnId: 'turn-1' }),
      timelineEntry({ id: 'active-1', turnId: 'turn-2' }),
      timelineEntry({ id: 'active-2', turnId: 'turn-2' }),
    ]

    expect(
      getWorkbookAgentProgressAnchorIndex({
        activeResponseTurnId: 'turn-2',
        showAssistantProgress: true,
        visibleEntries,
      }),
    ).toBe(2)
    expect(
      getWorkbookAgentProgressAnchorIndex({
        activeResponseTurnId: 'turn-2',
        showAssistantProgress: false,
        visibleEntries,
      }),
    ).toBe(-1)
  })

  it('summarizes citations and disclosure text without leaking markdown structure', () => {
    expect(
      summarizeTimelineCitations([
        { kind: 'revision', revision: 7 },
        { kind: 'range', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1', role: 'target' },
        { kind: 'range', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1', role: 'target' },
        { kind: 'range', sheetName: 'Sheet2', startAddress: 'B2', endAddress: 'C3', role: 'source' },
      ]),
    ).toEqual(['Target Sheet1!A1', 'Source Sheet2!B2:C3'])

    expect(summarizeDisclosureText('## Plan\nUse `A1` then [open](https://example.com).')).toBe('Plan Use A1 then open.')
  })

  it('derives review card state from permissions, preview, and shared review status', () => {
    expect(
      getReviewItemCardState({
        commandCount: 3,
        selectedCount: 2,
        preview: {
          ranges: [],
          structuralChanges: [],
          cellDiffs: [],
          effectSummary: {
            displayedCellDiffCount: 0,
            truncatedCellDiffs: false,
            inputChangeCount: 0,
            formulaChangeCount: 0,
            styleChangeCount: 0,
            numberFormatChangeCount: 0,
            structuralChangeCount: 0,
          },
        },
        sharedApprovalOwnerUserId: null,
        sharedReviewOwnerUserId: 'owner@example.com',
        sharedReviewStatus: 'approved',
        sharedReviewDecidedByUserId: 'reviewer@example.com',
        sharedReviewRecommendations: [recommendation('approved', 'a@example.com'), recommendation('rejected', 'b@example.com')],
        isApplyingReviewItem: false,
      }),
    ).toMatchObject({
      selectedCount: 2,
      hasFullSelection: false,
      canApply: true,
      applyLabel: 'Apply',
      sharedReviewOwnerLabel: 'Owner',
      sharedReviewDecisionLabel: 'Reviewer',
      recommendationSummary: '1 approval recommendation · 1 rejection recommendation',
    })

    expect(renderPreviewChangeKind('numberFormat')).toBe('number format')
  })

  it('keeps composer controls aligned with running and loading states', () => {
    expect(
      getWorkbookAgentComposerState({
        canInterruptTurn: false,
        draft: 'stop',
        isLoading: false,
        isRunning: true,
      }),
    ).toEqual({ sendAriaLabel: 'Stop', isSendDisabled: true })
    expect(
      getWorkbookAgentComposerState({
        canInterruptTurn: true,
        draft: '  ',
        isLoading: false,
        isRunning: false,
      }),
    ).toEqual({ sendAriaLabel: 'Send message', isSendDisabled: true })
    expect(
      shouldSubmitWorkbookAgentComposerKey({
        isRunning: false,
        key: 'Enter',
        shiftKey: false,
        isComposing: false,
      }),
    ).toBe(true)
  })

  it('hides the active thread from the thread switcher', () => {
    const active = threadSummary({ threadId: 'thread-active' })
    const sibling = threadSummary({ threadId: 'thread-sibling', scope: 'shared' })

    expect(
      getVisibleWorkbookAgentThreadSummaries({
        activeThreadId: 'thread-active',
        threadSummaries: [active, sibling],
      }),
    ).toEqual([sibling])
  })
})
