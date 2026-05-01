// @vitest-environment jsdom
import type { ComponentProps } from 'react'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { WorkbookAgentThreadSnapshot, WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { WorkbookAgentPanel } from '../WorkbookAgentPanel.js'

beforeEach(() => {
  ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

function createSnapshot(entry: WorkbookAgentTimelineEntry): WorkbookAgentThreadSnapshot {
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
    entries: [entry],
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
  }
}

function createPanelProps(snapshot: WorkbookAgentThreadSnapshot): ComponentProps<typeof WorkbookAgentPanel> {
  return {
    activeThreadId: 'thr-1',
    activeContextLabel: 'Sheet1!A1',
    activeResponseTurnId: null,
    optimisticEntries: [],
    snapshot,
    showAssistantProgress: false,
    activeReviewBundle: null,
    preview: null,
    sharedApprovalOwnerUserId: null,
    sharedReviewOwnerUserId: null,
    sharedReviewStatus: null,
    sharedReviewDecidedByUserId: null,
    sharedReviewRecommendations: [],
    currentUserSharedRecommendation: null,
    canFinalizeSharedBundle: false,
    canRecommendSharedBundle: false,
    selectedCommandIndexes: [],
    workflowRuns: [],
    cancellingWorkflowRunId: null,
    threadSummaries: [],
    draft: '',
    isLoading: false,
    isApplyingReviewItem: false,
    onApplyReviewItem: vi.fn(),
    onDraftChange: vi.fn(),
    onDismissReviewItem: vi.fn(),
    onReviewReviewItem: vi.fn(),
    onInterrupt: vi.fn(),
    onSelectAllReviewCommands: vi.fn(),
    onSelectThread: vi.fn(),
    onToggleReviewCommand: vi.fn(),
    onCancelWorkflowRun: vi.fn(),
    onSubmit: vi.fn(),
  }
}

function createAssistantEntry(text: string): WorkbookAgentTimelineEntry {
  return {
    id: 'assistant-copy-1',
    kind: 'assistant',
    turnId: 'turn-1',
    text,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [],
  }
}

function createUserEntry(text: string): WorkbookAgentTimelineEntry {
  return {
    id: 'user-copy-1',
    kind: 'user',
    turnId: 'turn-1',
    text,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [],
  }
}

describe('WorkbookAgentPanel message copy', () => {
  it('copies assistant message content and shows a checked state after click', async () => {
    const message = '**Done**\n\nUse Sheet1!A1 for the verified edit.'
    const writeText = vi.fn(async (_text: string) => {})
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText,
      } satisfies Pick<Clipboard, 'writeText'>,
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<WorkbookAgentPanel {...createPanelProps(createSnapshot(createAssistantEntry(message)))} />)
    })

    const copyButton = host.querySelector("[data-testid='workbook-agent-message-copy-assistant-copy-1']")
    expect(copyButton instanceof HTMLButtonElement).toBe(true)
    expect(copyButton?.getAttribute('aria-label')).toBe('Copy assistant message')
    expect(copyButton?.getAttribute('data-copy-state')).toBe('idle')

    await act(async () => {
      if (!(copyButton instanceof HTMLButtonElement)) {
        throw new Error('Assistant copy button not found')
      }
      copyButton.click()
      await Promise.resolve()
    })

    expect(writeText).toHaveBeenCalledWith(message)
    expect(copyButton?.getAttribute('aria-label')).toBe('Copied assistant message')
    expect(copyButton?.getAttribute('data-copy-state')).toBe('copied')

    await act(async () => {
      root.unmount()
    })
  })

  it('places the user copy action below the message card instead of inside the message', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<WorkbookAgentPanel {...createPanelProps(createSnapshot(createUserEntry('Copyable user message')))} />)
    })

    const messageCard = host.querySelector("[data-testid='workbook-agent-user-message-card-user-copy-1']")
    const actions = host.querySelector("[data-testid='workbook-agent-user-message-actions-user-copy-1']")
    const copyButton = host.querySelector("[data-testid='workbook-agent-message-copy-user-copy-1']")

    expect(messageCard?.textContent).toContain('Copyable user message')
    expect(messageCard?.contains(copyButton)).toBe(false)
    expect(actions?.contains(copyButton)).toBe(true)
    expect(actions?.previousElementSibling).toBe(messageCard)

    await act(async () => {
      root.unmount()
    })
  })
})
