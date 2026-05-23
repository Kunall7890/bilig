export const forbiddenZeroWorkflowRunReplicationColumns = [
  'mutation_executed',
  'verification_complete',
  'mutation_status',
  'mutation_receipt_json',
] as const

export interface ReleaseCheckAssetText {
  readonly file: string
  readonly text: string
}

export interface ForbiddenZeroWorkflowRunReplicationColumnReference {
  readonly columnName: (typeof forbiddenZeroWorkflowRunReplicationColumns)[number]
  readonly file: string
}

export function findForbiddenZeroWorkflowRunReplicationColumnReferences(
  assets: readonly ReleaseCheckAssetText[],
): readonly ForbiddenZeroWorkflowRunReplicationColumnReference[] {
  return assets.flatMap((asset) =>
    forbiddenZeroWorkflowRunReplicationColumns.flatMap((columnName) =>
      asset.text.includes(columnName)
        ? [
            {
              columnName,
              file: asset.file,
            },
          ]
        : [],
    ),
  )
}

export function assertNoForbiddenZeroWorkflowRunReplicationColumnReferences(assets: readonly ReleaseCheckAssetText[]): void {
  const references = findForbiddenZeroWorkflowRunReplicationColumnReferences(assets)
  if (references.length === 0) {
    return
  }

  const details = references.map((reference) => `${reference.columnName} in ${reference.file}`).join('; ')
  throw new Error(`Web bundle includes workbook_workflow_run columns that are intentionally not replicated through Zero: ${details}`)
}
