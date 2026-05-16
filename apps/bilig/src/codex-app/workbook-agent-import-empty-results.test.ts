import { describe, expect, it } from 'vitest'
import { createEmptyImportWorkflowResult, type ImportWorkflowTemplate } from './workbook-agent-import-empty-results.js'

const CASES: readonly {
  readonly workflowTemplate: ImportWorkflowTemplate
  readonly title: string
  readonly summary: string
  readonly artifactTitle: string
  readonly stepIds: readonly string[]
}[] = [
  {
    workflowTemplate: 'normalizeCurrentSheetHeaders',
    title: 'Normalize Current Sheet Headers',
    summary: 'Empty is empty, so there were no headers to normalize.',
    artifactTitle: 'Header Normalization Preview',
    stepIds: ['inspect-header-row', 'stage-header-normalization', 'draft-header-report'],
  },
  {
    workflowTemplate: 'normalizeCurrentSheetNumberFormats',
    title: 'Normalize Current Sheet Number Formats',
    summary: 'Empty is empty, so there were no numeric cells to format.',
    artifactTitle: 'Number Format Normalization Preview',
    stepIds: ['inspect-number-columns', 'stage-number-formats', 'draft-number-format-report'],
  },
  {
    workflowTemplate: 'normalizeCurrentSheetWhitespace',
    title: 'Normalize Current Sheet Whitespace',
    summary: 'Empty is empty, so there were no text cells to normalize.',
    artifactTitle: 'Whitespace Normalization Preview',
    stepIds: ['inspect-text-cells', 'stage-whitespace-normalization', 'draft-whitespace-report'],
  },
  {
    workflowTemplate: 'fillCurrentSheetFormulasDown',
    title: 'Fill Current Sheet Formulas Down',
    summary: 'Empty is empty, so there were no formula gaps to fill.',
    artifactTitle: 'Formula Fill-Down Preview',
    stepIds: ['inspect-formula-columns', 'stage-formula-fill', 'draft-formula-fill-report'],
  },
]

describe('createEmptyImportWorkflowResult', () => {
  it.each(CASES)('builds a durable no-op result for $workflowTemplate', (testCase) => {
    const result = createEmptyImportWorkflowResult(testCase.workflowTemplate, 'Empty')

    expect(result.title).toBe(testCase.title)
    expect(result.summary).toBe(testCase.summary)
    expect(result.artifact.title).toBe(testCase.artifactTitle)
    expect(result.artifact.text).toContain(`## ${testCase.artifactTitle}`)
    expect(result.artifact.text).toContain('Sheet: Empty')
    expect(result.citations).toEqual([])
    expect(result.commands).toBeUndefined()
    expect(result.goalText).toBeUndefined()
    expect(result.steps.map((step) => step.stepId)).toEqual(testCase.stepIds)
    expect(result.steps).toHaveLength(3)
    expect(result.steps[0]?.summary).toBe('Loaded Empty and found no populated cells.')
  })
})
