import { describe, expect, it } from 'vitest'
import type { WorkbookAgentWorkflowRun } from '@bilig/contracts'
import { mergeWorkbookAgentWorkflowRun, mergeWorkbookAgentWorkflowRuns } from '../workbook-agent-workflow-run-merge.js'

function createWorkflowRun(overrides: Partial<WorkbookAgentWorkflowRun> = {}): WorkbookAgentWorkflowRun {
  return {
    runId: 'wf-1',
    threadId: 'thr-1',
    startedByUserId: 'alex@example.com',
    workflowTemplate: 'normalizeCurrentSheetWhitespace',
    title: 'Normalize whitespace',
    summary: 'Normalizing whitespace.',
    status: 'running',
    createdAtUnixMs: 100,
    updatedAtUnixMs: 100,
    completedAtUnixMs: null,
    errorMessage: null,
    steps: [],
    artifact: null,
    ...overrides,
  }
}

describe('workbook agent workflow run merge', () => {
  it('keeps snapshot mutation proof and child rows when a newer Zero parent row has only replicated parent fields', () => {
    const snapshotRun = createWorkflowRun({
      summary: 'Applied change set but render proof is incomplete.',
      status: 'completed',
      updatedAtUnixMs: 200,
      completedAtUnixMs: 200,
      steps: [
        {
          stepId: 'apply',
          label: 'Apply workbook mutation',
          status: 'completed',
          summary: 'Applied revision r7.',
          updatedAtUnixMs: 200,
        },
      ],
      artifact: {
        kind: 'markdown',
        title: 'Mutation proof',
        text: 'Visible readback did not match yet.',
      },
      mutationExecuted: true,
      verificationComplete: false,
      mutationStatus: 'verification_incomplete',
      mutationReceipt: {
        status: 'verification_incomplete',
        visibleSceneProof: { valid: false },
      },
    })
    const zeroParentRun = createWorkflowRun({
      summary: 'Replicated parent summary after persistence.',
      status: 'completed',
      updatedAtUnixMs: 250,
      completedAtUnixMs: 250,
    })

    expect(mergeWorkbookAgentWorkflowRun(snapshotRun, zeroParentRun)).toEqual({
      ...zeroParentRun,
      steps: snapshotRun.steps,
      artifact: snapshotRun.artifact,
      mutationExecuted: true,
      verificationComplete: false,
      mutationStatus: 'verification_incomplete',
      mutationReceipt: snapshotRun.mutationReceipt,
    })
  })

  it('sorts merged workflow runs by the freshest parent update while preserving proof fields', () => {
    const snapshotRun = createWorkflowRun({
      runId: 'wf-1',
      updatedAtUnixMs: 200,
      mutationExecuted: true,
      verificationComplete: false,
      mutationStatus: 'verification_incomplete',
    })
    const zeroParentRun = createWorkflowRun({
      runId: 'wf-1',
      updatedAtUnixMs: 400,
      status: 'completed',
    })
    const zeroOnlyRun = createWorkflowRun({
      runId: 'wf-2',
      updatedAtUnixMs: 300,
      title: 'Describe workbook',
      workflowTemplate: 'summarizeWorkbook',
    })

    const merged = mergeWorkbookAgentWorkflowRuns({
      snapshotRuns: [snapshotRun],
      zeroRuns: [zeroOnlyRun, zeroParentRun],
    })

    expect(merged.map((run) => run.runId)).toEqual(['wf-1', 'wf-2'])
    expect(merged[0]).toMatchObject({
      runId: 'wf-1',
      updatedAtUnixMs: 400,
      mutationExecuted: true,
      verificationComplete: false,
      mutationStatus: 'verification_incomplete',
    })
  })
})
