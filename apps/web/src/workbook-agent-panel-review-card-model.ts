import type { WorkbookAgentPreviewChangeKind, WorkbookAgentPreviewSummary, WorkbookAgentSharedReviewRecommendation } from '@bilig/agent-api'
import { formatWorkbookCollaboratorLabel } from './workbook-presence-model.js'

export interface ReviewItemCardState {
  readonly selectedCount: number
  readonly hasFullSelection: boolean
  readonly sharedApprovalOwnerLabel: string | null
  readonly sharedReviewOwnerLabel: string | null
  readonly sharedReviewDecisionLabel: string | null
  readonly recommendationSummary: string | null
  readonly canApply: boolean
  readonly applyLabel: string
}

export function getReviewItemCardState(input: {
  readonly commandCount: number
  readonly selectedCount: number
  readonly preview: WorkbookAgentPreviewSummary | null
  readonly sharedApprovalOwnerUserId: string | null
  readonly sharedReviewOwnerUserId: string | null
  readonly sharedReviewStatus: 'pending' | 'approved' | 'rejected' | null
  readonly sharedReviewDecidedByUserId: string | null
  readonly sharedReviewRecommendations: readonly WorkbookAgentSharedReviewRecommendation[]
  readonly isApplyingReviewItem: boolean
}): ReviewItemCardState {
  const hasFullSelection = input.selectedCount === input.commandCount
  const sharedApprovalOwnerLabel = input.sharedApprovalOwnerUserId ? formatWorkbookCollaboratorLabel(input.sharedApprovalOwnerUserId) : null
  const sharedReviewOwnerLabel = input.sharedReviewOwnerUserId ? formatWorkbookCollaboratorLabel(input.sharedReviewOwnerUserId) : null
  const sharedReviewDecisionLabel = input.sharedReviewDecidedByUserId
    ? formatWorkbookCollaboratorLabel(input.sharedReviewDecidedByUserId)
    : null
  const recommendationSummary = formatRecommendationSummary(input.sharedReviewRecommendations)
  const canApply =
    input.preview !== null &&
    !input.isApplyingReviewItem &&
    input.selectedCount > 0 &&
    sharedApprovalOwnerLabel === null &&
    (input.sharedReviewStatus === null || input.sharedReviewStatus === 'approved')
  const applyLabel =
    input.selectedCount > 0 && !hasFullSelection
      ? input.isApplyingReviewItem
        ? 'Applying…'
        : 'Apply'
      : input.sharedReviewStatus === 'pending'
        ? 'Owner review'
        : input.sharedReviewStatus === 'rejected'
          ? 'Returned'
          : input.isApplyingReviewItem
            ? 'Applying…'
            : 'Apply'

  return {
    selectedCount: input.selectedCount,
    hasFullSelection,
    sharedApprovalOwnerLabel,
    sharedReviewOwnerLabel,
    sharedReviewDecisionLabel,
    recommendationSummary,
    canApply,
    applyLabel,
  }
}

export function formatRecommendationSummary(recommendations: readonly WorkbookAgentSharedReviewRecommendation[]): string | null {
  if (recommendations.length === 0) {
    return null
  }
  const approvalRecommendationCount = recommendations.filter((recommendation) => recommendation.decision === 'approved').length
  const rejectionRecommendationCount = recommendations.filter((recommendation) => recommendation.decision === 'rejected').length
  return `${String(approvalRecommendationCount)} approval ${approvalRecommendationCount === 1 ? 'recommendation' : 'recommendations'} · ${String(rejectionRecommendationCount)} rejection ${rejectionRecommendationCount === 1 ? 'recommendation' : 'recommendations'}`
}

export function renderPreviewChangeKind(kind: WorkbookAgentPreviewChangeKind): string {
  switch (kind) {
    case 'input':
      return 'value'
    case 'formula':
      return 'formula'
    case 'style':
      return 'style'
    case 'numberFormat':
      return 'number format'
  }
}
