import { Fragment, type RefObject, useEffect, useRef } from 'react'
import { ScrollArea } from '@base-ui/react/scroll-area'
import type { WorkbookAgentCommandBundle, WorkbookAgentPreviewSummary, WorkbookAgentSharedReviewRecommendation } from '@bilig/agent-api'
import type {
  WorkbookAgentThreadSnapshot,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
  WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
import { cn } from './cn.js'
import { WorkbookAgentComposer } from './workbook-agent-panel-composer.js'
import { AssistantProgressRow, WorkflowRunRow } from './workbook-agent-panel-history.js'
import {
  agentPanelEyebrowTextClass,
  agentPanelFooterClass,
  agentPanelScrollAreaContentClass,
  agentPanelScrollAreaRootClass,
  agentPanelScrollAreaScrollbarClass,
  agentPanelScrollAreaThumbClass,
  agentPanelScrollAreaViewportClass,
  agentPanelTimelineListClass,
} from './workbook-agent-panel-primitives.js'
import { ReviewItemCard } from './workbook-agent-panel-review-card.js'
import { agentPanelThemeStyle } from './workbook-agent-panel-theme.js'
import { ThreadSummaryStrip } from './workbook-agent-panel-thread-list.js'
import {
  getVisibleWorkbookAgentTimelineEntries,
  getWorkbookAgentProgressAnchorIndex,
  WorkbookAgentEntryRow,
} from './workbook-agent-panel-timeline.js'

export function WorkbookAgentPanel(props: {
  readonly activeThreadId: string | null
  readonly activeContextLabel: string | null
  readonly optimisticEntries?: readonly WorkbookAgentTimelineEntry[]
  readonly snapshot: WorkbookAgentThreadSnapshot | null
  readonly activeResponseTurnId: string | null
  readonly showAssistantProgress: boolean
  readonly activeReviewBundle: WorkbookAgentCommandBundle | null
  readonly preview: WorkbookAgentPreviewSummary | null
  readonly sharedApprovalOwnerUserId: string | null
  readonly sharedReviewOwnerUserId: string | null
  readonly sharedReviewStatus: 'pending' | 'approved' | 'rejected' | null
  readonly sharedReviewDecidedByUserId: string | null
  readonly sharedReviewRecommendations: readonly WorkbookAgentSharedReviewRecommendation[]
  readonly currentUserSharedRecommendation: 'approved' | 'rejected' | null
  readonly canFinalizeSharedBundle: boolean
  readonly canRecommendSharedBundle: boolean
  readonly canDismissReviewItem: boolean
  readonly selectedCommandIndexes: readonly number[]
  readonly workflowRuns: readonly WorkbookAgentWorkflowRun[]
  readonly canInterruptTurn?: boolean
  readonly canCancelWorkflowRun: (run: WorkbookAgentWorkflowRun) => boolean
  readonly cancellingWorkflowRunId: string | null
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[]
  readonly draft: string
  readonly isLoading: boolean
  readonly isApplyingReviewItem: boolean
  readonly onApplyReviewItem: () => void
  readonly onDraftChange: (value: string) => void
  readonly onDismissReviewItem: () => void
  readonly onReviewReviewItem: (decision: 'approved' | 'rejected') => void
  readonly onInterrupt: () => void
  readonly onSelectAllReviewCommands: () => void
  readonly onSelectThread: (threadId: string) => void
  readonly onToggleReviewCommand: (commandIndex: number) => void
  readonly onCancelWorkflowRun: (runId: string) => void
  readonly onSubmit: () => void
}) {
  const scrollRef = useRef<HTMLDivElement | null>(null)
  const optimisticEntries = props.optimisticEntries ?? []

  useEffect(() => {
    const node = scrollRef.current
    if (!node) {
      return
    }
    node.scrollTop = node.scrollHeight
  }, [optimisticEntries.length, props.snapshot?.entries.length, props.snapshot?.status])

  const isRunning = props.snapshot?.status === 'inProgress'
  const canInterruptTurn = props.canInterruptTurn ?? true
  const visibleEntries = getVisibleWorkbookAgentTimelineEntries({
    optimisticEntries,
    snapshotEntries: props.snapshot?.entries ?? [],
  })
  const progressAnchorIndex = getWorkbookAgentProgressAnchorIndex({
    activeResponseTurnId: props.activeResponseTurnId,
    showAssistantProgress: props.showAssistantProgress,
    visibleEntries,
  })

  return (
    <div
      className="flex h-full min-h-0 w-full flex-col bg-[var(--wb-app-bg)]"
      data-testid="workbook-agent-panel"
      id="workbook-agent-panel"
      style={agentPanelThemeStyle}
    >
      <WorkbookAgentPanelBody
        activeThreadId={props.activeThreadId}
        cancellingWorkflowRunId={props.cancellingWorkflowRunId}
        isLoading={props.isLoading}
        progressAnchorIndex={progressAnchorIndex}
        scrollRef={scrollRef}
        showAssistantProgress={props.showAssistantProgress}
        threadSummaries={props.threadSummaries}
        visibleEntries={visibleEntries}
        workflowRuns={props.workflowRuns}
        canCancelWorkflowRun={props.canCancelWorkflowRun}
        onCancelWorkflowRun={props.onCancelWorkflowRun}
        onSelectThread={props.onSelectThread}
      />
      <div className={agentPanelFooterClass()}>
        {props.activeReviewBundle ? (
          <div className="mb-3">
            <ReviewItemCard
              activeReviewBundle={props.activeReviewBundle}
              preview={props.preview}
              sharedApprovalOwnerUserId={props.sharedApprovalOwnerUserId}
              sharedReviewOwnerUserId={props.sharedReviewOwnerUserId}
              sharedReviewStatus={props.sharedReviewStatus}
              sharedReviewDecidedByUserId={props.sharedReviewDecidedByUserId}
              sharedReviewRecommendations={props.sharedReviewRecommendations}
              currentUserSharedRecommendation={props.currentUserSharedRecommendation}
              canFinalizeSharedBundle={props.canFinalizeSharedBundle}
              canRecommendSharedBundle={props.canRecommendSharedBundle}
              canDismissReviewItem={props.canDismissReviewItem}
              selectedCommandIndexes={props.selectedCommandIndexes}
              isApplyingReviewItem={props.isApplyingReviewItem}
              onApply={props.onApplyReviewItem}
              onDismiss={props.onDismissReviewItem}
              onReview={props.onReviewReviewItem}
              onSelectAll={props.onSelectAllReviewCommands}
              onToggleCommand={props.onToggleReviewCommand}
            />
          </div>
        ) : null}
        <WorkbookAgentComposer
          canInterruptTurn={canInterruptTurn}
          draft={props.draft}
          isLoading={props.isLoading}
          isRunning={isRunning}
          onDraftChange={props.onDraftChange}
          onInterrupt={props.onInterrupt}
          onSubmit={props.onSubmit}
        />
      </div>
    </div>
  )
}

function WorkbookAgentPanelBody(props: {
  readonly activeThreadId: string | null
  readonly cancellingWorkflowRunId: string | null
  readonly isLoading: boolean
  readonly progressAnchorIndex: number
  readonly scrollRef: RefObject<HTMLDivElement | null>
  readonly showAssistantProgress: boolean
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[]
  readonly visibleEntries: readonly WorkbookAgentTimelineEntry[]
  readonly workflowRuns: readonly WorkbookAgentWorkflowRun[]
  readonly canCancelWorkflowRun: (run: WorkbookAgentWorkflowRun) => boolean
  readonly onCancelWorkflowRun: (runId: string) => void
  readonly onSelectThread: (threadId: string) => void
}) {
  return (
    <ScrollArea.Root className={agentPanelScrollAreaRootClass()}>
      <ScrollArea.Viewport
        ref={props.scrollRef}
        className={agentPanelScrollAreaViewportClass()}
        data-testid="workbook-agent-panel-scroll-viewport"
      >
        <ScrollArea.Content className={agentPanelScrollAreaContentClass()}>
          <div className="box-border flex min-h-full flex-col bg-[var(--wb-app-bg)] px-2.5 py-2.5">
            <ThreadSummaryStrip
              activeThreadId={props.activeThreadId}
              threadSummaries={props.threadSummaries}
              onSelectThread={props.onSelectThread}
            />
            {props.isLoading ? null : props.visibleEntries.length > 0 || props.showAssistantProgress ? (
              <WorkbookAgentTimelineSection
                cancellingWorkflowRunId={props.cancellingWorkflowRunId}
                progressAnchorIndex={props.progressAnchorIndex}
                showAssistantProgress={props.showAssistantProgress}
                visibleEntries={props.visibleEntries}
                workflowRuns={props.workflowRuns}
                canCancelWorkflowRun={props.canCancelWorkflowRun}
                onCancelWorkflowRun={props.onCancelWorkflowRun}
              />
            ) : null}
          </div>
        </ScrollArea.Content>
      </ScrollArea.Viewport>
      <ScrollArea.Scrollbar className={agentPanelScrollAreaScrollbarClass()} orientation="vertical">
        <ScrollArea.Thumb className={agentPanelScrollAreaThumbClass()} />
      </ScrollArea.Scrollbar>
    </ScrollArea.Root>
  )
}

function WorkbookAgentTimelineSection(props: {
  readonly cancellingWorkflowRunId: string | null
  readonly progressAnchorIndex: number
  readonly showAssistantProgress: boolean
  readonly visibleEntries: readonly WorkbookAgentTimelineEntry[]
  readonly workflowRuns: readonly WorkbookAgentWorkflowRun[]
  readonly canCancelWorkflowRun: (run: WorkbookAgentWorkflowRun) => boolean
  readonly onCancelWorkflowRun: (runId: string) => void
}) {
  return (
    <div className="flex flex-col gap-2">
      <div className={agentPanelTimelineListClass()}>
        {props.visibleEntries.map((entry, index) => (
          <Fragment key={entry.id}>
            <div className="min-w-0 w-full max-w-full overflow-hidden">
              <WorkbookAgentEntryRow entry={entry} />
            </div>
            {props.showAssistantProgress && props.progressAnchorIndex === index ? <AssistantProgressRow /> : null}
          </Fragment>
        ))}
        {props.showAssistantProgress && props.progressAnchorIndex < 0 ? <AssistantProgressRow /> : null}
      </div>
      {props.workflowRuns.length > 0 ? (
        <WorkbookAgentWorkflowList
          cancellingWorkflowRunId={props.cancellingWorkflowRunId}
          workflowRuns={props.workflowRuns}
          canCancelWorkflowRun={props.canCancelWorkflowRun}
          onCancelWorkflowRun={props.onCancelWorkflowRun}
        />
      ) : null}
    </div>
  )
}

function WorkbookAgentWorkflowList(props: {
  readonly cancellingWorkflowRunId: string | null
  readonly workflowRuns: readonly WorkbookAgentWorkflowRun[]
  readonly canCancelWorkflowRun: (run: WorkbookAgentWorkflowRun) => boolean
  readonly onCancelWorkflowRun: (runId: string) => void
}) {
  return (
    <div className="pt-1">
      <div className={cn(agentPanelEyebrowTextClass(), 'mb-2')}>Workflows</div>
      <div className="flex flex-col gap-2">
        {props.workflowRuns.slice(0, 5).map((run) => {
          const onCancel = props.canCancelWorkflowRun(run)
            ? () => {
                props.onCancelWorkflowRun(run.runId)
              }
            : null
          return (
            <WorkflowRunRow
              key={run.runId}
              isCancelling={props.cancellingWorkflowRunId === run.runId}
              run={run}
              {...(onCancel ? { onCancel } : {})}
            />
          )
        })}
      </div>
    </div>
  )
}
