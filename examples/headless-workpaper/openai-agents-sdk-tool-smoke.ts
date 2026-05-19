import { Agent, RunContext, invokeFunctionTool, tool } from '@openai/agents'
import { z } from 'zod'

import {
  buildWorkbook,
  createWorkPaperToolHandlers,
  requireWorkPaperReadResult,
  requireWorkPaperWriteResult,
  type WorkPaperReadResult,
} from './ai-sdk-workpaper-tool-smoke-shared.ts'

type OpenAiAgentsSdkToolContext = {
  surface: 'provider-free-smoke'
}

const readSummaryInputSchema = z.object({
  range: z.string().default('Summary!A1:B5'),
})

const setInputCellInputSchema = z.object({
  sheetName: z.literal('Inputs'),
  address: z.string().regex(/^[A-Z]+[1-9][0-9]*$/),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
})

const workpaper = buildWorkbook()
const handlers = createWorkPaperToolHandlers(workpaper)
const tools = {
  readWorkPaperSummary: tool({
    name: 'read_workpaper_summary',
    description: 'Read computed WorkPaper summary values for a small A1 range.',
    parameters: readSummaryInputSchema,
    execute: async ({ range = 'Summary!A1:B5' }) => handlers.readWorkPaperSummary(range),
  }),

  setWorkPaperInputCell: tool({
    name: 'set_workpaper_input_cell',
    description: 'Set one validated WorkPaper input cell and return before/after formula readback.',
    parameters: setInputCellInputSchema,
    execute: async (args) => handlers.setWorkPaperInputCell(args),
  }),
}

const agent = new Agent<OpenAiAgentsSdkToolContext>({
  name: 'WorkPaper verification agent',
  instructions: 'Use the WorkPaper tools for workbook reads and edits. Answer only from computed readback returned by the tools.',
  tools: [tools.readWorkPaperSummary, tools.setWorkPaperInputCell],
})

const runContext = new RunContext<OpenAiAgentsSdkToolContext>({
  surface: 'provider-free-smoke',
})
const readResult = requireWorkPaperReadResult(
  await invokeFunctionTool({
    tool: tools.readWorkPaperSummary,
    runContext,
    input: JSON.stringify({
      range: 'Summary!A1:B5',
    }),
  }),
)
const writeResult = requireWorkPaperWriteResult(
  await invokeFunctionTool({
    tool: tools.setWorkPaperInputCell,
    runContext,
    input: JSON.stringify({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    }),
  }),
)

const output = {
  apiShape: 'OpenAI Agents SDK Agent -> tool() -> invokeFunctionTool()',
  package: '@openai/agents',
  agentName: agent.name,
  toolNames: [tools.readWorkPaperSummary.name, tools.setWorkPaperInputCell.name],
  toolSchemas: [tools.readWorkPaperSummary, tools.setWorkPaperInputCell].map((sdkTool) => ({
    name: sdkTool.name,
    strict: sdkTool.strict,
    parameterKeys: readSchemaPropertyKeys(sdkTool.parameters),
  })),
  readResult,
  writeResult,
  finalText: `Edited ${writeResult.editedCell}; expected ARR changed from ${writeResult.before.expectedArr} to ${writeResult.after.expectedArr}.`,
}

assertOpenAiAgentsSdkProof(output)

console.log(JSON.stringify(output, null, 2))

function assertOpenAiAgentsSdkProof(proof: typeof output): void {
  if (proof.apiShape !== 'OpenAI Agents SDK Agent -> tool() -> invokeFunctionTool()') {
    throw new Error(`Unexpected OpenAI Agents SDK API shape: ${proof.apiShape}`)
  }

  if (!sameJson(proof.toolNames, ['read_workpaper_summary', 'set_workpaper_input_cell'])) {
    throw new Error(`Unexpected OpenAI Agents SDK tool names: ${JSON.stringify(proof.toolNames)}`)
  }

  if (readGridNumber(proof.readResult, 2, 1, 'read summary expected ARR') !== 60000) {
    throw new Error(`Unexpected summary before edit: ${JSON.stringify(proof.readResult.values)}`)
  }

  if (proof.writeResult.editedCell !== 'Inputs!B3') {
    throw new Error(`Unexpected edited cell: ${proof.writeResult.editedCell}`)
  }

  if (proof.writeResult.before.expectedArr !== 60000 || proof.writeResult.after.expectedArr !== 96000) {
    throw new Error(`Unexpected ARR readback: ${JSON.stringify(proof.writeResult)}`)
  }

  if (
    proof.writeResult.checks.previousValue !== 0.25 ||
    proof.writeResult.checks.newValue !== 0.4 ||
    !proof.writeResult.checks.formulasPersisted ||
    !proof.writeResult.checks.restoredMatchesAfter ||
    !proof.writeResult.checks.expectedArrChanged
  ) {
    throw new Error(`OpenAI Agents SDK WorkPaper checks failed: ${JSON.stringify(proof.writeResult.checks)}`)
  }

  if (!proof.finalText.includes('Edited Inputs!B3')) {
    throw new Error(`Unexpected final OpenAI Agents SDK text: ${proof.finalText}`)
  }
}

function readGridNumber(result: WorkPaperReadResult, row: number, col: number, label: string): number {
  const cell = result.values[row]?.[col]
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readSchemaPropertyKeys(schema: unknown): string[] {
  if (typeof schema !== 'object' || schema === null || Array.isArray(schema)) {
    return []
  }
  const properties = Reflect.get(schema, 'properties')
  if (typeof properties !== 'object' || properties === null || Array.isArray(properties)) {
    return []
  }
  return Object.keys(properties)
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
