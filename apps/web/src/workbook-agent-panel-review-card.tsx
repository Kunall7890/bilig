import { Button } from '@base-ui/react/button'
import type { WorkbookAgentCommandBundle, WorkbookAgentPreviewSummary, WorkbookAgentSharedReviewRecommendation } from '@bilig/agent-api'
import { cn } from './cn.js'
import { agentPanelLabelTextClass } from './workbook-agent-panel-primitives.js'
import { getReviewItemCardState } from './workbook-agent-panel-review-card-model.js'
import {
  ReviewCellDiffList,
  ReviewCommandChecklist,
  ReviewDecisionActions,
  ReviewPreviewDetails,
  ReviewSharedStatusMessages,
} from './workbook-agent-panel-review-card-sections.js'
import { workbookButtonClass, workbookSurfaceClass } from './workbook-shell-chrome.js'

export { getReviewItemCardState, renderPreviewChangeKind, type ReviewItemCardState } from './workbook-agent-panel-review-card-model.js'

export function ReviewItemCard(props: {
  readonly activeReviewBundle: WorkbookAgentCommandBundle
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
  readonly isApplyingReviewItem: boolean
  readonly onApply: () => void
  readonly onDismiss: () => void
  readonly onReview: (decision: 'approved' | 'rejected') => void
  readonly onSelectAll: () => void
  readonly onToggleCommand: (commandIndex: number) => void
}) {
  const state = getReviewItemCardState({
    commandCount: props.activeReviewBundle.commands.length,
    selectedCount: props.selectedCommandIndexes.length,
    preview: props.preview,
    sharedApprovalOwnerUserId: props.sharedApprovalOwnerUserId,
    sharedReviewOwnerUserId: props.sharedReviewOwnerUserId,
    sharedReviewStatus: props.sharedReviewStatus,
    sharedReviewDecidedByUserId: props.sharedReviewDecidedByUserId,
    sharedReviewRecommendations: props.sharedReviewRecommendations,
    isApplyingReviewItem: props.isApplyingReviewItem,
  })

  return (
    <div className={cn(workbookSurfaceClass({ emphasis: 'raised' }), 'border-[var(--wb-border-strong)] px-3 py-3')}>
      <div className={cn(agentPanelLabelTextClass(), 'font-semibold')}>{props.activeReviewBundle.summary}</div>
      <ReviewCommandChecklist
        activeReviewBundle={props.activeReviewBundle}
        selectedCommandIndexes={props.selectedCommandIndexes}
        state={state}
        onSelectAll={props.onSelectAll}
        onToggleCommand={props.onToggleCommand}
      />
      <ReviewPreviewDetails preview={props.preview} fallbackRanges={props.activeReviewBundle.affectedRanges} />
      <ReviewSharedStatusMessages
        currentUserSharedRecommendation={props.currentUserSharedRecommendation}
        sharedReviewStatus={props.sharedReviewStatus}
        state={state}
      />
      {props.canFinalizeSharedBundle && props.sharedReviewStatus !== null ? (
        <ReviewDecisionActions approveLabel="Approve" rejectLabel="Reject" onReview={props.onReview} />
      ) : null}
      {props.canRecommendSharedBundle && props.sharedReviewStatus === 'pending' ? (
        <ReviewDecisionActions approveLabel="Recommend approve" rejectLabel="Recommend reject" onReview={props.onReview} />
      ) : null}
      {props.preview?.cellDiffs?.length ? <ReviewCellDiffList preview={props.preview} /> : null}
      <div className="mt-3 flex items-center justify-end gap-2">
        <Button
          className={workbookButtonClass({ tone: 'neutral' })}
          data-testid="workbook-agent-dismiss-review-item"
          disabled={!props.canDismissReviewItem}
          type="button"
          onClick={props.onDismiss}
        >
          Clear
        </Button>
        <Button
          className={workbookButtonClass({ tone: 'accent', weight: 'strong' })}
          data-testid="workbook-agent-apply-review-item"
          disabled={!state.canApply}
          type="button"
          onClick={props.onApply}
        >
          {state.applyLabel}
        </Button>
      </div>
    </div>
  )
}
