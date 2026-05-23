import { describe, expect, it } from 'vitest'

import {
  assertNoForbiddenZeroWorkflowRunReplicationColumnReferences,
  findForbiddenZeroWorkflowRunReplicationColumnReferences,
} from '../release-check-asset-contracts.ts'

describe('release check asset contracts', () => {
  it('allows workflow mutation proof fields as API payload camelCase without adding Zero parent columns', () => {
    const assets = [
      {
        file: 'apps/web/dist/assets/workbook-vendor.js',
        text: 'const run = { artifactJson: artifact_json, mutationExecuted: true, verificationComplete: false, mutationReceipt: receipt }',
      },
    ]

    expect(findForbiddenZeroWorkflowRunReplicationColumnReferences(assets)).toEqual([])
    expect(() => assertNoForbiddenZeroWorkflowRunReplicationColumnReferences(assets)).not.toThrow()
  })

  it('rejects built web assets that request unreplicated workflow run database columns', () => {
    const assets = [
      {
        file: 'apps/web/dist/assets/index.js',
        text: 'workbook_workflow_run.columns({ mutationExecuted: boolean().from("mutation_executed") })',
      },
    ]

    expect(findForbiddenZeroWorkflowRunReplicationColumnReferences(assets)).toEqual([
      {
        columnName: 'mutation_executed',
        file: 'apps/web/dist/assets/index.js',
      },
    ])
    expect(() => assertNoForbiddenZeroWorkflowRunReplicationColumnReferences(assets)).toThrow(
      'Web bundle includes workbook_workflow_run columns that are intentionally not replicated through Zero',
    )
  })
})
