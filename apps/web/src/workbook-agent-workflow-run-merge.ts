import type { WorkbookAgentWorkflowRun } from '@bilig/contracts'

function hasWorkflowSteps(run: WorkbookAgentWorkflowRun): boolean {
  return run.steps.length > 0
}

function hasWorkflowArtifact(run: WorkbookAgentWorkflowRun): boolean {
  return run.artifact !== null
}

function hasMutationProof(run: WorkbookAgentWorkflowRun): boolean {
  return (
    run.mutationExecuted !== undefined ||
    run.verificationComplete !== undefined ||
    run.mutationStatus !== undefined ||
    run.mutationReceipt !== undefined
  )
}

function mergeOptionalMutationProof<T>(nextValue: T | null | undefined, previousValue: T | null | undefined): T | null | undefined {
  return nextValue === undefined || nextValue === null ? previousValue : nextValue
}

export function mergeWorkbookAgentWorkflowRun(
  previous: WorkbookAgentWorkflowRun,
  next: WorkbookAgentWorkflowRun,
): WorkbookAgentWorkflowRun {
  const parent = next.updatedAtUnixMs >= previous.updatedAtUnixMs ? next : previous
  const childSource = hasWorkflowSteps(next) || hasWorkflowArtifact(next) ? next : previous
  const proofSource = hasMutationProof(next) ? next : previous
  return {
    ...parent,
    steps: hasWorkflowSteps(childSource) ? childSource.steps : [],
    artifact: hasWorkflowArtifact(childSource) ? childSource.artifact : null,
    mutationExecuted: mergeOptionalMutationProof(proofSource.mutationExecuted, previous.mutationExecuted),
    verificationComplete: mergeOptionalMutationProof(proofSource.verificationComplete, previous.verificationComplete),
    mutationStatus: mergeOptionalMutationProof(proofSource.mutationStatus, previous.mutationStatus),
    mutationReceipt: mergeOptionalMutationProof(proofSource.mutationReceipt, previous.mutationReceipt),
  }
}

export function mergeWorkbookAgentWorkflowRuns(input: {
  readonly snapshotRuns: readonly WorkbookAgentWorkflowRun[]
  readonly zeroRuns: readonly WorkbookAgentWorkflowRun[]
}): readonly WorkbookAgentWorkflowRun[] {
  const merged = new Map<string, WorkbookAgentWorkflowRun>()
  for (const run of input.snapshotRuns) {
    merged.set(run.runId, run)
  }
  for (const run of input.zeroRuns) {
    const previous = merged.get(run.runId)
    merged.set(run.runId, previous ? mergeWorkbookAgentWorkflowRun(previous, run) : run)
  }
  return [...merged.values()].toSorted((left, right) => right.updatedAtUnixMs - left.updatedAtUnixMs)
}
