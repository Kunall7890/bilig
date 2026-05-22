interface RenderedReadbackLike {
  readonly requested: boolean
  readonly matched: boolean | null
  readonly sourceKind?: 'selection' | 'visibleRange' | null
}

interface FormulaIssueReportLike {
  readonly summary: {
    readonly actionableIssueCount: number
  }
}

interface InvariantReportLike {
  readonly summary: {
    readonly ok: boolean
  }
}

interface RecalculationStatusLike {
  readonly upToDate: boolean
}

export interface WorkbookAgentVerificationStatusInput {
  readonly renderedReadback: readonly RenderedReadbackLike[]
  readonly formulaIssues: FormulaIssueReportLike | null
  readonly invariants: InvariantReportLike | null
  readonly recalculationStatus?: RecalculationStatusLike | null
  readonly requireRenderedSelection?: boolean
  readonly requireTargetRange?: boolean
  readonly targetRangeCount?: number
}

export interface WorkbookAgentVerificationStatus {
  readonly verificationComplete: boolean
  readonly renderedComplete: boolean
  readonly formulaComplete: boolean
  readonly invariantsComplete: boolean
  readonly recalculationComplete: boolean
  readonly renderedSelectionComplete: boolean
  readonly targetRangeComplete: boolean
  readonly missingChecks: readonly string[]
}

export function summarizeWorkbookAgentVerificationStatus(input: WorkbookAgentVerificationStatusInput): WorkbookAgentVerificationStatus {
  const targetRangeComplete = input.requireTargetRange === true ? (input.targetRangeCount ?? 0) > 0 : true
  const requiredRenderedRangeCount = input.requireTargetRange === true && targetRangeComplete ? (input.targetRangeCount ?? 0) : 0
  const renderedComplete =
    requiredRenderedRangeCount > 0
      ? input.renderedReadback.length >= requiredRenderedRangeCount &&
        input.renderedReadback.every((proof) => proof.requested && proof.matched === true)
      : input.renderedReadback.every((proof) => !proof.requested || proof.matched === true)
  const renderedSelectionComplete =
    input.requireRenderedSelection === true && renderedComplete
      ? input.renderedReadback.every((proof) => !proof.requested || proof.sourceKind === 'selection')
      : true
  const formulaComplete = input.formulaIssues !== null && input.formulaIssues.summary.actionableIssueCount === 0
  const invariantsComplete = input.invariants !== null && input.invariants.summary.ok
  const recalculationComplete = input.recalculationStatus?.upToDate === true
  const missingChecks = [
    targetRangeComplete ? null : 'targetRange',
    renderedComplete ? null : 'renderedReadback',
    renderedSelectionComplete ? null : 'renderedSelection',
    input.recalculationStatus === null || input.recalculationStatus === undefined
      ? 'recalculationStatus'
      : recalculationComplete
        ? null
        : 'recalculationStale',
    input.formulaIssues === null ? 'formulaIssues' : formulaComplete ? null : 'formulaIssuesClean',
    input.invariants === null ? 'invariants' : invariantsComplete ? null : 'invariantsClean',
  ].filter((check): check is string => check !== null)
  return {
    verificationComplete:
      targetRangeComplete &&
      renderedComplete &&
      renderedSelectionComplete &&
      recalculationComplete &&
      formulaComplete &&
      invariantsComplete,
    renderedComplete,
    formulaComplete,
    invariantsComplete,
    recalculationComplete,
    renderedSelectionComplete,
    targetRangeComplete,
    missingChecks,
  }
}
