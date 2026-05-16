import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type { WorkbookAgentTimelineCitation, WorkbookAgentWorkflowArtifact } from '@bilig/contracts'

export type ImportWorkflowTemplate =
  | 'normalizeCurrentSheetHeaders'
  | 'normalizeCurrentSheetNumberFormats'
  | 'normalizeCurrentSheetWhitespace'
  | 'fillCurrentSheetFormulasDown'

interface ImportWorkflowStepResult {
  readonly stepId: string
  readonly label: string
  readonly summary: string
}

export interface ImportWorkflowExecutionResult {
  readonly title: string
  readonly summary: string
  readonly artifact: WorkbookAgentWorkflowArtifact
  readonly citations: readonly WorkbookAgentTimelineCitation[]
  readonly steps: readonly ImportWorkflowStepResult[]
  readonly commands?: readonly WorkbookAgentCommand[]
  readonly goalText?: string
}

interface EmptyWorkflowStepDefinition {
  readonly stepId: string
  readonly label: string
  readonly summary: (sheetName: string) => string
}

interface EmptyWorkflowDefinition {
  readonly title: string
  readonly emptySummary: string
  readonly artifactTitle: string
  readonly artifactHeading: string
  readonly noChangeSummary: string
  readonly steps: readonly EmptyWorkflowStepDefinition[]
}

const EMPTY_WORKFLOW_DEFINITIONS = {
  normalizeCurrentSheetHeaders: {
    title: 'Normalize Current Sheet Headers',
    emptySummary: 'there were no headers to normalize',
    artifactTitle: 'Header Normalization Preview',
    artifactHeading: '## Header Normalization Preview',
    noChangeSummary: 'No header changes were needed because the sheet is empty.',
    steps: [
      {
        stepId: 'inspect-header-row',
        label: 'Inspect header row',
        summary: (sheetName) => `Loaded ${sheetName} and found no populated cells.`,
      },
      {
        stepId: 'stage-header-normalization',
        label: 'Stage header normalization',
        summary: () => 'Sheet is empty. Header normalization change count: 0.',
      },
      {
        stepId: 'draft-header-report',
        label: 'Draft header report',
        summary: () => 'Prepared the durable empty-sheet header report for the thread.',
      },
    ],
  },
  normalizeCurrentSheetNumberFormats: {
    title: 'Normalize Current Sheet Number Formats',
    emptySummary: 'there were no numeric cells to format',
    artifactTitle: 'Number Format Normalization Preview',
    artifactHeading: '## Number Format Normalization Preview',
    noChangeSummary: 'No number-format changes were needed because the sheet is empty.',
    steps: [
      {
        stepId: 'inspect-number-columns',
        label: 'Inspect numeric columns',
        summary: (sheetName) => `Loaded ${sheetName} and found no populated cells.`,
      },
      {
        stepId: 'stage-number-formats',
        label: 'Stage number formats',
        summary: () => 'Sheet is empty. Number-format change count: 0.',
      },
      {
        stepId: 'draft-number-format-report',
        label: 'Draft number-format report',
        summary: () => 'Prepared the durable empty-sheet number-format report for the thread.',
      },
    ],
  },
  normalizeCurrentSheetWhitespace: {
    title: 'Normalize Current Sheet Whitespace',
    emptySummary: 'there were no text cells to normalize',
    artifactTitle: 'Whitespace Normalization Preview',
    artifactHeading: '## Whitespace Normalization Preview',
    noChangeSummary: 'No whitespace changes were needed because the sheet is empty.',
    steps: [
      {
        stepId: 'inspect-text-cells',
        label: 'Inspect text cells',
        summary: (sheetName) => `Loaded ${sheetName} and found no populated cells.`,
      },
      {
        stepId: 'stage-whitespace-normalization',
        label: 'Stage whitespace normalization',
        summary: () => 'Sheet is empty. Whitespace normalization change count: 0.',
      },
      {
        stepId: 'draft-whitespace-report',
        label: 'Draft whitespace report',
        summary: () => 'Prepared the durable empty-sheet whitespace report for the thread.',
      },
    ],
  },
  fillCurrentSheetFormulasDown: {
    title: 'Fill Current Sheet Formulas Down',
    emptySummary: 'there were no formula gaps to fill',
    artifactTitle: 'Formula Fill-Down Preview',
    artifactHeading: '## Formula Fill-Down Preview',
    noChangeSummary: 'No fill-down changes were needed because the sheet is empty.',
    steps: [
      {
        stepId: 'inspect-formula-columns',
        label: 'Inspect formula columns',
        summary: (sheetName) => `Loaded ${sheetName} and found no populated cells.`,
      },
      {
        stepId: 'stage-formula-fill',
        label: 'Stage formula fill-down',
        summary: () => 'Sheet is empty. Formula fill-down change count: 0.',
      },
      {
        stepId: 'draft-formula-fill-report',
        label: 'Draft formula fill-down report',
        summary: () => 'Prepared the durable empty-sheet formula fill-down report for the thread.',
      },
    ],
  },
} satisfies Record<ImportWorkflowTemplate, EmptyWorkflowDefinition>

export function createEmptyImportWorkflowResult(
  workflowTemplate: ImportWorkflowTemplate,
  sheetName: string,
): ImportWorkflowExecutionResult {
  const definition = EMPTY_WORKFLOW_DEFINITIONS[workflowTemplate]
  return {
    title: definition.title,
    summary: `${sheetName} is empty, so ${definition.emptySummary}.`,
    artifact: {
      kind: 'markdown',
      title: definition.artifactTitle,
      text: [definition.artifactHeading, '', `Sheet: ${sheetName}`, '', definition.noChangeSummary].join('\n'),
    },
    citations: [],
    steps: definition.steps.map((step) => ({
      stepId: step.stepId,
      label: step.label,
      summary: step.summary(sheetName),
    })),
  }
}
