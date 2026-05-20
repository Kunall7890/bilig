import { Button } from '@base-ui/react/button'
import { describeWorkbookAgentCommand } from '@bilig/agent-api'
import type { WorkbookAgentCommandBundle, WorkbookAgentPreviewSummary } from '@bilig/agent-api'
import { cn } from './cn.js'
import { agentPanelEyebrowTextClass, agentPanelLabelTextClass, agentPanelMetaTextClass } from './workbook-agent-panel-primitives.js'
import { PreviewRangeList } from './workbook-agent-panel-history.js'
import { renderPreviewChangeKind, type ReviewItemCardState } from './workbook-agent-panel-review-card-model.js'
import { workbookAlertClass, workbookButtonClass, workbookInsetClass, workbookPillClass } from './workbook-shell-chrome.js'

export function ReviewCommandChecklist(props: {
  readonly activeReviewBundle: WorkbookAgentCommandBundle
  readonly selectedCommandIndexes: readonly number[]
  readonly state: Pick<ReviewItemCardState, 'hasFullSelection' | 'selectedCount'>
  readonly onSelectAll: () => void
  readonly onToggleCommand: (commandIndex: number) => void
}) {
  return (
    <div className={cn(workbookInsetClass(), 'mt-3 px-2 py-2')}>
      <div className="flex items-center justify-between gap-3">
        <div className={agentPanelEyebrowTextClass()}>
          {String(props.state.selectedCount)}/{String(props.activeReviewBundle.commands.length)}
        </div>
        {!props.state.hasFullSelection ? (
          <Button
            className="text-[12px] leading-5 font-medium text-[var(--wb-accent)] transition-colors hover:brightness-[0.95] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
            type="button"
            onClick={props.onSelectAll}
          >
            All
          </Button>
        ) : null}
      </div>
      <div className="mt-2 flex flex-col gap-2">
        {props.activeReviewBundle.commands.map((command, index) => {
          const checked = props.selectedCommandIndexes.includes(index)
          const commandLabel = describeWorkbookAgentCommand(command)
          return (
            <div
              key={`${props.activeReviewBundle.id}:${JSON.stringify(command)}`}
              className={cn(
                'flex items-start gap-3 rounded-[var(--wb-radius-control)] border px-3 py-2 transition-colors',
                checked ? 'border-[var(--wb-accent-ring)] bg-[var(--wb-surface)]' : 'border-[var(--wb-border)] bg-[var(--wb-surface)]',
              )}
            >
              <input
                aria-label={`Toggle workbook review item change ${String(index + 1)}: ${commandLabel}`}
                checked={checked}
                className="mt-0.5 h-4 w-4 rounded border-[var(--wb-border)] text-[var(--wb-accent)] focus:ring-[var(--wb-accent-ring)]"
                data-testid={`workbook-agent-review-command-toggle-${String(index)}`}
                type="checkbox"
                onChange={() => {
                  props.onToggleCommand(index)
                }}
              />
              <div className="min-w-0">
                <div className={agentPanelLabelTextClass()}>{commandLabel}</div>
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}

export function ReviewPreviewDetails(props: {
  readonly preview: WorkbookAgentPreviewSummary | null
  readonly fallbackRanges: WorkbookAgentCommandBundle['affectedRanges']
}) {
  return (
    <>
      <PreviewRangeList ranges={props.preview?.ranges ?? props.fallbackRanges} />
      {props.preview?.structuralChanges?.length ? (
        <div className={cn(workbookInsetClass(), agentPanelMetaTextClass(), 'mt-2 border-transparent px-2 py-2')}>
          {props.preview.structuralChanges.join(' · ')}
        </div>
      ) : null}
    </>
  )
}

export function ReviewSharedStatusMessages(props: {
  readonly currentUserSharedRecommendation: 'approved' | 'rejected' | null
  readonly sharedReviewStatus: 'pending' | 'approved' | 'rejected' | null
  readonly state: Pick<
    ReviewItemCardState,
    'recommendationSummary' | 'sharedApprovalOwnerLabel' | 'sharedReviewDecisionLabel' | 'sharedReviewOwnerLabel'
  >
}) {
  return (
    <>
      {props.state.sharedApprovalOwnerLabel ? (
        <div className={cn(workbookAlertClass({ tone: 'warning' }), agentPanelMetaTextClass(), 'mt-2 border-[var(--wb-border-strong)]')}>
          Owner review routes medium/high-risk changes to {props.state.sharedApprovalOwnerLabel} on this shared thread.
        </div>
      ) : null}
      {props.state.recommendationSummary ? (
        <div className={cn(workbookInsetClass(), agentPanelMetaTextClass(), 'mt-2 border-transparent px-2 py-2')}>
          {props.state.recommendationSummary}
          {props.currentUserSharedRecommendation
            ? ` You recommended ${props.currentUserSharedRecommendation === 'approved' ? 'approval' : 'rejection'}.`
            : ''}
        </div>
      ) : null}
      {props.state.sharedReviewOwnerLabel && props.sharedReviewStatus ? (
        <div
          className={cn(
            workbookAlertClass({
              tone: props.sharedReviewStatus === 'rejected' ? 'danger' : 'warning',
            }),
            agentPanelMetaTextClass(),
            'mt-2 border-[var(--wb-border-strong)]',
          )}
        >
          {props.sharedReviewStatus === 'pending'
            ? `Owner review is in progress with ${props.state.sharedReviewOwnerLabel}.`
            : props.sharedReviewStatus === 'approved'
              ? `Approved by ${props.state.sharedReviewDecisionLabel ?? props.state.sharedReviewOwnerLabel}.`
              : `Returned by ${props.state.sharedReviewDecisionLabel ?? props.state.sharedReviewOwnerLabel}.`}
        </div>
      ) : null}
    </>
  )
}

export function ReviewDecisionActions(props: {
  readonly approveLabel: string
  readonly rejectLabel: string
  readonly onReview: (decision: 'approved' | 'rejected') => void
}) {
  return (
    <div className="mt-2 flex items-center justify-end gap-2">
      <Button
        className={workbookButtonClass({ tone: 'neutral' })}
        data-testid="workbook-agent-review-item-reject"
        type="button"
        onClick={() => {
          props.onReview('rejected')
        }}
      >
        {props.rejectLabel}
      </Button>
      <Button
        className={workbookButtonClass({ tone: 'accent' })}
        data-testid="workbook-agent-review-item-approve"
        type="button"
        onClick={() => {
          props.onReview('approved')
        }}
      >
        {props.approveLabel}
      </Button>
    </div>
  )
}

export function ReviewCellDiffList(props: { readonly preview: WorkbookAgentPreviewSummary }) {
  return (
    <div className="mt-2 overflow-hidden rounded-[var(--wb-radius-control)] border border-[var(--wb-border)]">
      <div className="max-h-44 overflow-y-auto">
        {props.preview.cellDiffs.map((diff) => (
          <div
            key={`${diff.sheetName}:${diff.address}`}
            className="grid grid-cols-[minmax(0,1fr)_minmax(0,1fr)] gap-x-2 border-t border-[var(--wb-border)] px-2 py-2 first:border-t-0"
          >
            <div className={cn(agentPanelLabelTextClass(), 'col-span-2')}>
              {diff.sheetName}!{diff.address}
            </div>
            <div className="col-span-2 mt-1 flex flex-wrap gap-1">
              {diff.changeKinds.map((kind) => (
                <span key={kind} className={workbookPillClass({ tone: 'neutral' })}>
                  {renderPreviewChangeKind(kind)}
                </span>
              ))}
            </div>
            <div className={agentPanelMetaTextClass()}>{(diff.beforeFormula ?? String(diff.beforeInput ?? '')) || '(empty)'}</div>
            <div className={agentPanelLabelTextClass()}>{(diff.afterFormula ?? String(diff.afterInput ?? '')) || '(empty)'}</div>
          </div>
        ))}
      </div>
    </div>
  )
}
