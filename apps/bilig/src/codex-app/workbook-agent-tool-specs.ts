import { WORKBOOK_AGENT_TOOL_NAMES, type CodexDynamicToolSpec } from '@bilig/agent-api'
import { workbookAgentAnnotationToolSpecs } from './workbook-agent-annotation-tools.js'
import { workbookAgentAuditToolSpecs } from './workbook-agent-audit-tools.js'
import { workbookAgentConditionalFormatToolSpecs } from './workbook-agent-conditional-format-tools.js'
import { workbookAgentMediaToolSpecs } from './workbook-agent-media-tools.js'
import { workbookAgentObjectToolSpecs } from './workbook-agent-object-tools.js'
import { workbookAgentProtectionToolSpecs } from './workbook-agent-protection-tools.js'
import { workbookAgentSheetReadToolSpecs } from './workbook-agent-sheet-read-tools.js'
import { workbookAgentStructuralToolSpecs } from './workbook-agent-structural-tools.js'
import { workbookAgentValidationToolSpecs } from './workbook-agent-validation-tools.js'
import { workbookAgentStylePatchJsonSchema } from './workbook-agent-style-patches.js'
import { cellRangeRefJsonSchema, rangeOrSelectorJsonSchema, workbookSemanticSelectorJsonSchema } from './workbook-agent-selector-tooling.js'

const rangeTargetJsonSchema = {
  oneOf: [cellRangeRefJsonSchema, workbookSemanticSelectorJsonSchema],
}

function createDynamicToolSpecs(): readonly CodexDynamicToolSpec[] {
  return [
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.getContext,
      description:
        'Read the current browser workbook context, including selection geometry, the visible viewport, freeze panes, and hidden or resized axes in view.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
      description: 'Read a workbook summary with sheet names, populated cell counts, and used ranges.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.setActiveSheet,
      description: 'Make a sheet the active browser sheet for the attached workbook view, optionally moving selection to one cell.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['sheetName'],
        properties: {
          sheetName: { type: 'string' },
          address: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.setSelection,
      description: 'Move the attached browser workbook selection to a cell or rectangular range.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['address'],
        properties: {
          sheetName: { type: 'string' },
          address: { type: 'string' },
          endAddress: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRenderedSelection,
      description: 'Read the latest browser-rendered snapshot for the selected cells and compare it with authoritative workbook state.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRenderedRange,
      description:
        'Read the latest browser-rendered snapshot for a range when it is cached by the attached view and compare it with authoritative workbook state.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['sheetName', 'startAddress', 'endAddress'],
        properties: {
          sheetName: { type: 'string' },
          startAddress: { type: 'string' },
          endAddress: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify,
      description:
        'Verify the latest applied workbook state by recalculating, reading authoritative cells, reading rendered browser cells when cached, scanning formula issues, and checking invariants.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          range: cellRangeRefJsonSchema,
          includeFormulaIssues: { type: 'boolean' },
          includeInvariants: { type: 'boolean' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.undoWorkbookMutation,
      description:
        'Undo the latest undoable workbook mutation for this assistant user, or revert a specific revision, then return the new head revision and verification context.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          revision: { type: 'number' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.listNamedRanges,
      description: 'List workbook named ranges and named references, including resolved cell ranges and structured references.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.listTables,
      description: 'List workbook tables with sheet location, range, header/totals settings, and column names.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    ...workbookAgentSheetReadToolSpecs,
    ...workbookAgentAnnotationToolSpecs,
    ...workbookAgentConditionalFormatToolSpecs,
    ...workbookAgentObjectToolSpecs,
    ...workbookAgentMediaToolSpecs,
    ...workbookAgentProtectionToolSpecs,
    ...workbookAgentValidationToolSpecs,
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRange,
      description:
        'Read a rectangular cell range or selector target, including inputs, formulas, style ids, number-format ids, referenced formatting records, and sheet-state metadata for that window.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          sheetName: { type: 'string' },
          startAddress: { type: 'string' },
          endAddress: { type: 'string' },
          selector: workbookSemanticSelectorJsonSchema,
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readSelection,
      description:
        'Read the currently selected range from the attached browser workbook context with values, formulas, formatting catalogs, and local sheet-state metadata.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange,
      description:
        'Read the currently visible viewport range from the attached browser workbook context with values, formulas, formatting catalogs, and local sheet-state metadata.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges,
      description: 'Read the most recent durable workbook changes, including revisions, summaries, and affected ranges.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          limit: { type: 'number' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
      description:
        'Run a built-in workbook workflow for durable summaries, analysis, cleanup, search, rollups, review tabs, and safe structural tasks.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['workflowTemplate'],
        properties: {
          workflowTemplate: {
            type: 'string',
            enum: [
              'summarizeWorkbook',
              'summarizeCurrentSheet',
              'describeRecentChanges',
              'findFormulaIssues',
              'highlightFormulaIssues',
              'repairFormulaIssues',
              'highlightCurrentSheetOutliers',
              'styleCurrentSheetHeaders',
              'normalizeCurrentSheetHeaders',
              'normalizeCurrentSheetNumberFormats',
              'normalizeCurrentSheetWhitespace',
              'fillCurrentSheetFormulasDown',
              'traceSelectionDependencies',
              'explainSelectionCell',
              'searchWorkbookQuery',
              'createCurrentSheetRollup',
              'createCurrentSheetReviewTab',
              'createSheet',
              'renameCurrentSheet',
              'hideCurrentRow',
              'hideCurrentColumn',
              'unhideCurrentRow',
              'unhideCurrentColumn',
            ],
          },
          query: { type: 'string' },
          sheetName: { type: 'string' },
          limit: { type: 'number' },
          name: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
      description:
        'Explain one cell, including input, current value, formula, display format, style record, number-format record, version, cycle status, and direct precedents/dependents. Defaults to the current selection when no address is provided.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          sheetName: { type: 'string' },
          address: { type: 'string' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues,
      description: 'Scan the workbook for broken formulas, error cells, cycles, and formulas still running through the JS fallback path.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          sheetName: { type: 'string' },
          limit: { type: 'number' },
        },
      },
    },
    ...workbookAgentAuditToolSpecs,
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook,
      description: 'Search workbook sheet names, addresses, formulas, inputs, and displayed values through the warm local runtime.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['query'],
        properties: {
          query: { type: 'string' },
          sheetName: { type: 'string' },
          limit: { type: 'number' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.traceDependencies,
      description:
        'Trace workbook precedents and dependents from one cell for multiple hops. Defaults to the current selection when no address is provided.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          sheetName: { type: 'string' },
          address: { type: 'string' },
          direction: {
            type: 'string',
            enum: ['precedents', 'dependents', 'both'],
          },
          depth: { type: 'number' },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.writeRange,
      description:
        'Write a rectangular matrix of spreadsheet inputs starting at a top-left address or selector target. Use primitives for literals, {formula} for formulas, and null to clear a cell.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['values'],
        properties: {
          sheetName: { type: 'string' },
          startAddress: { type: 'string' },
          selector: workbookSemanticSelectorJsonSchema,
          values: {
            type: 'array',
            items: {
              type: 'array',
              items: {
                oneOf: [
                  { type: 'string' },
                  { type: 'number' },
                  { type: 'boolean' },
                  { type: 'null' },
                  {
                    oneOf: [
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'text' },
                          value: { type: 'string' },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'number' },
                          value: { oneOf: [{ type: 'string' }, { type: 'number' }] },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'date' },
                          value: { oneOf: [{ type: 'string' }, { type: 'number' }] },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'value'],
                        properties: {
                          type: { type: 'string', const: 'boolean' },
                          value: { oneOf: [{ type: 'string' }, { type: 'boolean' }] },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type'],
                        properties: {
                          type: { type: 'string', const: 'blank' },
                        },
                      },
                      {
                        type: 'object',
                        additionalProperties: false,
                        required: ['type', 'formula'],
                        properties: {
                          type: { type: 'string', const: 'formula' },
                          formula: { type: 'string' },
                        },
                      },
                    ],
                  },
                  {
                    type: 'object',
                    additionalProperties: false,
                    required: ['value'],
                    properties: {
                      value: {
                        oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }],
                      },
                    },
                  },
                  {
                    type: 'object',
                    additionalProperties: false,
                    required: ['formula'],
                    properties: {
                      formula: { type: 'string' },
                    },
                  },
                ],
              },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.setFormula,
      description:
        'Write one or more formulas into a target range or selector-resolved anchor. Use this when the request is explicitly about formulas rather than generic cell input.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['formulas'],
        properties: {
          range: cellRangeRefJsonSchema,
          selector: workbookSemanticSelectorJsonSchema,
          formulas: {
            type: 'array',
            items: {
              type: 'array',
              items: { type: 'string' },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.clearRange,
      description: 'Clear a rectangular range of cells or a selector-resolved workbook region.',
      inputSchema: rangeOrSelectorJsonSchema,
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.formatRange,
      description:
        'Apply style and/or number-format changes to a range or selector target. Use patch for style properties and numberFormat for number formatting.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          range: cellRangeRefJsonSchema,
          selector: workbookSemanticSelectorJsonSchema,
          patch: workbookAgentStylePatchJsonSchema,
          numberFormat: {
            oneOf: [{ type: 'string' }, { type: 'object' }],
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.fillRange,
      description:
        'Fill a target range from a source range using spreadsheet fill semantics. Source and target may be concrete ranges or semantic selectors.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['source', 'target'],
        properties: {
          source: rangeTargetJsonSchema,
          target: rangeTargetJsonSchema,
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.copyRange,
      description: 'Copy a source range into a target range. Source and target may be concrete ranges or semantic selectors.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['source', 'target'],
        properties: {
          source: rangeTargetJsonSchema,
          target: rangeTargetJsonSchema,
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.moveRange,
      description: 'Move a source range into a target range. Source and target may be concrete ranges or semantic selectors.',
      inputSchema: {
        type: 'object',
        additionalProperties: false,
        required: ['source', 'target'],
        properties: {
          source: rangeTargetJsonSchema,
          target: rangeTargetJsonSchema,
        },
      },
    },
    ...workbookAgentStructuralToolSpecs,
  ] satisfies readonly CodexDynamicToolSpec[]
}

export const workbookAgentDynamicToolSpecs = createDynamicToolSpecs()
