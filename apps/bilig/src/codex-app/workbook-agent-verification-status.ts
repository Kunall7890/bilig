interface RenderedReadbackLike {
  readonly requested: boolean
  readonly matched: boolean | null
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

export interface WorkbookAgentVerificationStatusInput {
  readonly renderedReadback: readonly RenderedReadbackLike[]
  readonly formulaIssues: FormulaIssueReportLike | null
  readonly invariants: InvariantReportLike | null
  readonly requireTargetRange?: boolean
  readonly targetRangeCount?: number
}

export interface WorkbookAgentVerificationStatus {
  readonly verificationComplete: boolean
  readonly renderedComplete: boolean
  readonly formulaComplete: boolean
  readonly invariantsComplete: boolean
  readonly targetRangeComplete: boolean
  readonly missingChecks: readonly string[]
}

export function summarizeWorkbookAgentVerificationStatus(input: WorkbookAgentVerificationStatusInput): WorkbookAgentVerificationStatus {
  const renderedComplete = input.renderedReadback.every((proof) => !proof.requested || proof.matched === true)
  const formulaComplete = input.formulaIssues !== null && input.formulaIssues.summary.actionableIssueCount === 0
  const invariantsComplete = input.invariants !== null && input.invariants.summary.ok
  const targetRangeComplete = input.requireTargetRange === true ? (input.targetRangeCount ?? 0) > 0 : true
  const missingChecks = [
    targetRangeComplete ? null : 'targetRange',
    renderedComplete ? null : 'renderedReadback',
    input.formulaIssues === null ? 'formulaIssues' : formulaComplete ? null : 'formulaIssuesClean',
    input.invariants === null ? 'invariants' : invariantsComplete ? null : 'invariantsClean',
  ].filter((check): check is string => check !== null)
  return {
    verificationComplete: targetRangeComplete && renderedComplete && formulaComplete && invariantsComplete,
    renderedComplete,
    formulaComplete,
    invariantsComplete,
    targetRangeComplete,
    missingChecks,
  }
}
