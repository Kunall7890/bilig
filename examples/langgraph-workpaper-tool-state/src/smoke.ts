import { AIMessage, ToolMessage } from '@langchain/core/messages'
import { tool } from '@langchain/core/tools'
import { END, MessagesAnnotation, START, StateGraph } from '@langchain/langgraph'
import { ToolNode } from '@langchain/langgraph/prebuilt'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'
import { z } from 'zod'

interface QuoteSummary {
  readonly subtotal: number
  readonly discountAmount: number
  readonly taxableAmount: number
  readonly taxAmount: number
  readonly total: number
  readonly marginAmount: number
}

interface WorkPaperProof {
  readonly editedCell: 'Inputs!B2'
  readonly before: QuoteSummary
  readonly after: QuoteSummary
  readonly afterRestore: QuoteSummary
  readonly persistedDocumentBytes: number
  readonly verified: boolean
}

interface LangGraphWorkPaperSmokeOutput {
  readonly framework: 'langgraphjs-toolnode'
  readonly graphNodes: readonly ['agent_requests_workpaper_tools', 'tools']
  readonly toolMessageNames: readonly ['read_workpaper_quote', 'set_workpaper_quantity']
  readonly readBefore: QuoteSummary
  readonly proof: WorkPaperProof
  readonly stateProof: {
    readonly messageCount: number
    readonly toolMessageCount: number
    readonly quantityToolCallId: 'call_set_quantity'
  }
}

const quoteWorkbook = buildWorkbook()

try {
  const tools = createWorkPaperTools(quoteWorkbook)
  const toolNode = new ToolNode(tools)

  const graph = new StateGraph(MessagesAnnotation)
    .addNode('agent_requests_workpaper_tools', () => ({
      messages: [
        new AIMessage({
          content: '',
          tool_calls: [
            {
              name: 'read_workpaper_quote',
              args: {},
              id: 'call_read_quote',
              type: 'tool_call',
            },
            {
              name: 'set_workpaper_quantity',
              args: { quantity: 18 },
              id: 'call_set_quantity',
              type: 'tool_call',
            },
          ],
        }),
      ],
    }))
    .addNode('tools', toolNode)
    .addEdge(START, 'agent_requests_workpaper_tools')
    .addEdge('agent_requests_workpaper_tools', 'tools')
    .addEdge('tools', END)
    .compile()

  const state = await graph.invoke({ messages: [] })
  const toolMessages = state.messages.filter((message): message is ToolMessage => message instanceof ToolMessage)
  const readMessage = requireToolMessage(toolMessages, 'call_read_quote')
  const writeMessage = requireToolMessage(toolMessages, 'call_set_quantity')
  const readBefore = requireQuoteSummary(parseToolMessageJson(readMessage), 'read_workpaper_quote')
  const proof = requireWorkPaperProof(parseToolMessageJson(writeMessage))

  const output: LangGraphWorkPaperSmokeOutput = {
    framework: 'langgraphjs-toolnode',
    graphNodes: ['agent_requests_workpaper_tools', 'tools'],
    toolMessageNames: ['read_workpaper_quote', 'set_workpaper_quantity'],
    readBefore,
    proof,
    stateProof: {
      messageCount: state.messages.length,
      toolMessageCount: toolMessages.length,
      quantityToolCallId: 'call_set_quantity',
    },
  }

  assertSmokeOutput(output)
  console.log(JSON.stringify(output, null, 2))
} finally {
  quoteWorkbook.dispose()
}

function buildWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Field', 'Value'],
      ['Quantity', 12],
      ['Unit price', 125],
      ['Discount rate', 0.1],
      ['Tax rate', 0.08],
      ['Unit cost', 52],
    ],
    Summary: [
      ['Field', 'Value'],
      ['Subtotal', '=Inputs!B2*Inputs!B3'],
      ['Discount amount', '=B2*Inputs!B4'],
      ['Taxable amount', '=B2-B3'],
      ['Tax amount', '=B4*Inputs!B5'],
      ['Total', '=B4+B5'],
      ['Margin amount', '=B4-(Inputs!B2*Inputs!B6)'],
    ],
  })
}

function createWorkPaperTools(workbook: WorkPaper) {
  return [
    tool(() => readQuoteSummary(workbook), {
      name: 'read_workpaper_quote',
      description: 'Read the formula-backed quote summary from the current Bilig WorkPaper.',
      schema: z.object({}),
    }),
    tool(({ quantity }) => setQuantityAndProve(workbook, quantity), {
      name: 'set_workpaper_quantity',
      description: 'Set Inputs!B2, recalculate formulas, persist WorkPaper JSON, restore it, and return readback proof.',
      schema: z.object({
        quantity: z.number().finite().positive(),
      }),
    }),
  ]
}

function setQuantityAndProve(workbook: WorkPaper, quantity: number): WorkPaperProof {
  const inputsSheet = requireSheet(workbook, 'Inputs')
  const before = readQuoteSummary(workbook)

  workbook.setCellContents({ sheet: inputsSheet, row: 1, col: 1 }, quantity)
  const after = readQuoteSummary(workbook)

  const document = exportWorkPaperDocument(workbook, { includeConfig: true })
  const serialized = serializeWorkPaperDocument(document)
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))

  try {
    const afterRestore = readQuoteSummary(restored)

    return {
      editedCell: 'Inputs!B2',
      before,
      after,
      afterRestore,
      persistedDocumentBytes: new TextEncoder().encode(serialized).byteLength,
      verified: sameSummary(after, afterRestore) && before.total !== after.total,
    }
  } finally {
    restored.dispose()
  }
}

function readQuoteSummary(workbook: WorkPaper): QuoteSummary {
  const sheet = requireSheet(workbook, 'Summary')

  return {
    subtotal: readNumberCell(workbook, sheet, 1, 'subtotal'),
    discountAmount: readNumberCell(workbook, sheet, 2, 'discount amount'),
    taxableAmount: readNumberCell(workbook, sheet, 3, 'taxable amount'),
    taxAmount: readNumberCell(workbook, sheet, 4, 'tax amount'),
    total: readNumberCell(workbook, sheet, 5, 'total'),
    marginAmount: readNumberCell(workbook, sheet, 6, 'margin amount'),
  }
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)

  if (sheetId === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }

  return sheetId
}

function readNumberCell(workbook: WorkPaper, sheet: number, row: number, label: string): number {
  const cell = workbook.getCellValue({ sheet, row, col: 1 })
  const value = typeof cell === 'object' && cell !== null ? Reflect.get(cell, 'value') : cell

  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`Expected numeric ${label}, got ${JSON.stringify(cell)}`)
  }

  return value
}

function requireToolMessage(messages: readonly ToolMessage[], toolCallId: string): ToolMessage {
  const message = messages.find((item) => item.tool_call_id === toolCallId)

  if (message === undefined) {
    throw new Error(`Missing ToolMessage ${toolCallId}; received ${messages.map((item) => item.tool_call_id).join(', ')}`)
  }

  return message
}

function parseToolMessageJson(message: ToolMessage): unknown {
  if (typeof message.content !== 'string') {
    throw new Error(`Expected string ToolMessage content, got ${JSON.stringify(message.content)}`)
  }

  return JSON.parse(message.content)
}

function requireWorkPaperProof(value: unknown): WorkPaperProof {
  if (!isRecord(value)) {
    throw new Error(`Expected WorkPaper proof object, got ${JSON.stringify(value)}`)
  }

  const editedCell = value.editedCell
  const before = requireQuoteSummary(value.before, 'before')
  const after = requireQuoteSummary(value.after, 'after')
  const afterRestore = requireQuoteSummary(value.afterRestore, 'afterRestore')
  const persistedDocumentBytes = value.persistedDocumentBytes
  const verified = value.verified

  if (
    editedCell !== 'Inputs!B2' ||
    typeof persistedDocumentBytes !== 'number' ||
    !Number.isFinite(persistedDocumentBytes) ||
    typeof verified !== 'boolean'
  ) {
    throw new Error(`Malformed WorkPaper proof: ${JSON.stringify(value)}`)
  }

  return {
    editedCell,
    before,
    after,
    afterRestore,
    persistedDocumentBytes,
    verified,
  }
}

function requireQuoteSummary(value: unknown, label: string): QuoteSummary {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} quote summary, got ${JSON.stringify(value)}`)
  }

  return {
    subtotal: readFiniteNumberProperty(value, 'subtotal', label),
    discountAmount: readFiniteNumberProperty(value, 'discountAmount', label),
    taxableAmount: readFiniteNumberProperty(value, 'taxableAmount', label),
    taxAmount: readFiniteNumberProperty(value, 'taxAmount', label),
    total: readFiniteNumberProperty(value, 'total', label),
    marginAmount: readFiniteNumberProperty(value, 'marginAmount', label),
  }
}

function assertSmokeOutput(output: LangGraphWorkPaperSmokeOutput): void {
  if (output.readBefore.total !== 1458) {
    throw new Error(`Expected pre-edit total 1458, got ${output.readBefore.total}`)
  }

  if (output.proof.before.total !== 1458 || output.proof.after.total !== 2187 || output.proof.afterRestore.total !== 2187) {
    throw new Error(`Unexpected total proof: ${JSON.stringify(output.proof)}`)
  }

  if (!output.proof.verified || output.proof.persistedDocumentBytes <= 0) {
    throw new Error(`WorkPaper proof failed: ${JSON.stringify(output.proof)}`)
  }

  if (output.stateProof.toolMessageCount !== 2) {
    throw new Error(`Expected two ToolMessages, got ${output.stateProof.toolMessageCount}`)
  }
}

function sameSummary(left: QuoteSummary, right: QuoteSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readFiniteNumberProperty(value: Record<string, unknown>, property: keyof QuoteSummary, label: string): number {
  const nestedValue = value[property]

  if (typeof nestedValue !== 'number' || !Number.isFinite(nestedValue)) {
    throw new Error(`Expected numeric ${label}.${property}, got ${JSON.stringify(nestedValue)}`)
  }

  return nestedValue
}
