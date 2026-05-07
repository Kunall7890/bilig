import { setRangeNumberFormatArgsSchema } from '@bilig/zero-sync'
import { z } from 'zod'
import { cellRangeRefSchema, rangeOrSelectorSchema } from './workbook-agent-selector-tooling.js'
import { workbookAgentStylePatchSchema } from './workbook-agent-style-patches.js'
import { workbookSemanticSelectorSchema } from './workbook-selector-resolver.js'

export const MAX_MUTATION_RANGE_CELLS = 400
export const MAX_READ_RANGE_CELLS = 4000

export const inspectCellToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  address: z.string().min(1).optional(),
})
export const formulaIssueToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  limit: z.number().int().positive().max(200).optional(),
})
export const readRecentChangesToolArgsSchema = z.object({
  limit: z.number().int().positive().max(50).optional(),
})
export const startWorkflowToolArgsSchema = z.discriminatedUnion('workflowTemplate', [
  z.object({
    workflowTemplate: z.literal('summarizeWorkbook'),
  }),
  z.object({
    workflowTemplate: z.literal('summarizeCurrentSheet'),
  }),
  z.object({
    workflowTemplate: z.literal('describeRecentChanges'),
  }),
  z.object({
    workflowTemplate: z.literal('findFormulaIssues'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('highlightFormulaIssues'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('repairFormulaIssues'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('highlightCurrentSheetOutliers'),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(100).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('styleCurrentSheetHeaders'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('normalizeCurrentSheetHeaders'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('normalizeCurrentSheetNumberFormats'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('normalizeCurrentSheetWhitespace'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('fillCurrentSheetFormulasDown'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('traceSelectionDependencies'),
  }),
  z.object({
    workflowTemplate: z.literal('explainSelectionCell'),
  }),
  z.object({
    workflowTemplate: z.literal('searchWorkbookQuery'),
    query: z.string().trim().min(1),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(50).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('createCurrentSheetRollup'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('createCurrentSheetReviewTab'),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal('createSheet'),
    name: z.string().trim().min(1),
  }),
  z.object({
    workflowTemplate: z.literal('renameCurrentSheet'),
    name: z.string().trim().min(1),
  }),
  z.object({
    workflowTemplate: z.literal('hideCurrentRow'),
  }),
  z.object({
    workflowTemplate: z.literal('hideCurrentColumn'),
  }),
  z.object({
    workflowTemplate: z.literal('unhideCurrentRow'),
  }),
  z.object({
    workflowTemplate: z.literal('unhideCurrentColumn'),
  }),
])
export type WorkbookAgentStartWorkflowRequest = z.infer<typeof startWorkflowToolArgsSchema>

export const searchWorkbookToolArgsSchema = z.object({
  query: z.string().trim().min(1),
  sheetName: z.string().min(1).optional(),
  limit: z.number().int().positive().max(50).optional(),
})
export const traceDependenciesToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  address: z.string().min(1).optional(),
  direction: z.enum(['precedents', 'dependents', 'both']).optional(),
  depth: z.number().int().positive().max(4).optional(),
})
export const setActiveSheetToolArgsSchema = z.object({
  sheetName: z.string().trim().min(1),
  address: z.string().trim().min(1).optional(),
})
export const setSelectionToolArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
  address: z.string().trim().min(1),
  endAddress: z.string().trim().min(1).optional(),
})
export const readRenderedRangeToolArgsSchema = z.object({
  sheetName: z.string().trim().min(1),
  startAddress: z.string().trim().min(1),
  endAddress: z.string().trim().min(1),
})
export const applyAndVerifyToolArgsSchema = z.object({
  range: cellRangeRefSchema.optional(),
  includeFormulaIssues: z.boolean().optional(),
  includeInvariants: z.boolean().optional(),
})
export const undoWorkbookMutationToolArgsSchema = z.object({
  revision: z.number().int().positive().optional(),
})

export const clearRangeToolArgsSchema = rangeOrSelectorSchema
export const formatRangeToolArgsSchema = z
  .object({
    range: cellRangeRefSchema.optional(),
    selector: workbookSemanticSelectorSchema.optional(),
    patch: workbookAgentStylePatchSchema.optional(),
    numberFormat: setRangeNumberFormatArgsSchema.shape.format.optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: 'Provide exactly one of range or selector',
  })
  .refine((value) => value.patch !== undefined || value.numberFormat !== undefined, {
    message: 'patch or numberFormat is required',
  })
