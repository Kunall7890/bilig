import { mkdirSync, rmSync } from 'node:fs'
import { join } from 'node:path'

import { AIMessage, ToolMessage } from '@langchain/core/messages'
import { ToolNode } from '@langchain/langgraph/prebuilt'
import { MultiServerMCPClient } from '@langchain/mcp-adapters'

type JsonObject = Record<string, unknown>

interface SmokeOutput {
  readonly framework: 'langchainjs-mcp-adapters-toolnode'
  readonly mcpTransport: 'stdio'
  readonly workpaperPackage: '@bilig/workpaper@latest'
  readonly discoveredTools: readonly string[]
  readonly toolMessageNames: readonly string[]
  readonly editedCell: 'Inputs!B3'
  readonly dependentCell: 'Summary!B3'
  readonly before: number
  readonly after: number
  readonly afterRestart: number
  readonly displayValue: string
  readonly persistence: {
    readonly persisted: boolean
    readonly serializedBytes: number
  }
  readonly exportedDocumentBytes: number
  readonly checks: {
    readonly discoveredFileBackedTools: boolean
    readonly toolNodeReturnedToolMessages: boolean
    readonly dependentCellChanged: boolean
    readonly persistedToDisk: boolean
    readonly restartReadbackMatchesAfter: boolean
    readonly displayValueRead: boolean
    readonly exportedWorkPaperDocument: boolean
  }
  readonly verified: boolean
}

const workspaceDir = join(process.cwd(), '.tmp')
const workpaperPath = join(workspaceDir, 'pricing.workpaper.json')

mkdirSync(workspaceDir, { recursive: true })

let client: MultiServerMCPClient | undefined
let restartClient: MultiServerMCPClient | undefined
try {
  client = createBiligMcpClient({ initDemoWorkPaper: true, writable: true })
  const { discoveredTools, toolNode } = await loadToolNode(client)
  const toolNames = {
    readCell: requireToolName(discoveredTools, 'read_cell'),
    setCellContents: requireToolName(discoveredTools, 'set_cell_contents'),
    getCellDisplayValue: requireToolName(discoveredTools, 'get_cell_display_value'),
    exportWorkPaperDocument: requireToolName(discoveredTools, 'export_workpaper_document'),
  }

  const beforeMessage = await invokeOneTool(toolNode, {
    id: 'call_read_before',
    name: toolNames.readCell,
    args: { sheetName: 'Summary', address: 'B3' },
  })
  const writeMessage = await invokeOneTool(toolNode, {
    id: 'call_write_input',
    name: toolNames.setCellContents,
    args: {
      sheetName: 'Inputs',
      address: 'B3',
      value: '=0.4',
    },
  })
  const afterMessage = await invokeOneTool(toolNode, {
    id: 'call_read_after',
    name: toolNames.readCell,
    args: { sheetName: 'Summary', address: 'B3' },
  })
  const displayMessage = await invokeOneTool(toolNode, {
    id: 'call_display_value',
    name: toolNames.getCellDisplayValue,
    args: { sheetName: 'Summary', address: 'B3' },
  })
  const exportMessage = await invokeOneTool(toolNode, {
    id: 'call_export_document',
    name: toolNames.exportWorkPaperDocument,
    args: { includeConfig: true },
  })

  const beforeResult = parseToolMessageJson(beforeMessage)
  const writeResult = parseToolMessageJson(writeMessage)
  const afterResult = parseToolMessageJson(afterMessage)
  const displayResult = parseToolMessageJson(displayMessage)
  const exportResult = parseToolMessageJson(exportMessage)
  const before = readCellNumber(beforeResult, 'read_cell.value')
  const after = readCellNumber(afterResult, 'read_cell after.value')
  const displayValue = requireString(displayResult['displayValue'], 'displayValue')
  const persistence = requireRecord(writeResult['persistence'], 'persistence')
  const writeChecks = requireRecord(writeResult['checks'], 'set_cell_contents checks')
  const exportedDocumentBytes = requireNumber(exportResult['serializedBytes'], 'export_workpaper_document.serializedBytes')

  await client.close()
  client = undefined

  restartClient = createBiligMcpClient({ initDemoWorkPaper: false, writable: false })
  const restartToolNode = (await loadToolNode(restartClient)).toolNode
  const restartMessage = await invokeOneTool(restartToolNode, {
    id: 'call_read_after_restart',
    name: toolNames.readCell,
    args: { sheetName: 'Summary', address: 'B3' },
  })
  const afterRestart = readCellNumber(parseToolMessageJson(restartMessage), 'read_cell after restart.value')

  const checks = {
    discoveredFileBackedTools: [
      'list_sheets',
      'read_range',
      'read_cell',
      'set_cell_contents',
      'get_cell_display_value',
      'export_workpaper_document',
      'validate_formula',
    ].every((name) => discoveredTools.includes(name)),
    toolNodeReturnedToolMessages: [beforeMessage, writeMessage, afterMessage, displayMessage, exportMessage, restartMessage].every(
      (message) => message instanceof ToolMessage,
    ),
    dependentCellChanged: before === 60_000 && after === 96_000,
    persistedToDisk: persistence['persisted'] === true && requireNumber(persistence['serializedBytes'], 'persistence.serializedBytes') > 0,
    restartReadbackMatchesAfter: writeChecks['persisted'] === true && afterRestart === after,
    displayValueRead: displayValue === '96000',
    exportedWorkPaperDocument: isRecord(exportResult['document']) && exportedDocumentBytes > 0,
  }
  const output: SmokeOutput = {
    framework: 'langchainjs-mcp-adapters-toolnode',
    mcpTransport: 'stdio',
    workpaperPackage: '@bilig/workpaper@latest',
    discoveredTools,
    toolMessageNames: [beforeMessage, writeMessage, afterMessage, displayMessage, exportMessage, restartMessage].map(messageName),
    editedCell: 'Inputs!B3',
    dependentCell: 'Summary!B3',
    before,
    after,
    afterRestart,
    displayValue,
    persistence: {
      persisted: persistence['persisted'] === true,
      serializedBytes: requireNumber(persistence['serializedBytes'], 'persistence.serializedBytes'),
    },
    exportedDocumentBytes,
    checks,
    verified: Object.values(checks).every(Boolean),
  }

  if (!output.verified) {
    throw new Error(`LangChain MCP WorkPaper smoke failed: ${JSON.stringify(output.checks)}`)
  }

  console.log(JSON.stringify(output, null, 2))
} finally {
  await client?.close()
  await restartClient?.close()
  rmSync(workspaceDir, { force: true, recursive: true })
}

function createBiligMcpClient(options: { readonly initDemoWorkPaper: boolean; readonly writable: boolean }): MultiServerMCPClient {
  const args = ['exec', '--yes', '--package', '@bilig/workpaper@latest', '--', 'bilig-workpaper-mcp', '--workpaper', workpaperPath]
  if (options.initDemoWorkPaper) {
    args.push('--init-demo-workpaper')
  }
  if (options.writable) {
    args.push('--writable')
  }

  return new MultiServerMCPClient({
    throwOnLoadError: true,
    prefixToolNameWithServerName: false,
    useStandardContentBlocks: true,
    mcpServers: {
      bilig_workpaper: {
        transport: 'stdio',
        command: 'npm',
        args,
      },
    },
  })
}

async function loadToolNode(mcpClient: MultiServerMCPClient): Promise<{
  readonly discoveredTools: readonly string[]
  readonly toolNode: ToolNode
}> {
  const tools = await mcpClient.getTools()
  return {
    discoveredTools: tools.map((item) => item.name).toSorted(),
    toolNode: new ToolNode(tools),
  }
}

async function invokeOneTool(
  toolNode: ToolNode,
  call: { readonly id: string; readonly name: string; readonly args: JsonObject },
): Promise<ToolMessage> {
  const result: unknown = await toolNode.invoke({
    messages: [
      new AIMessage({
        content: '',
        tool_calls: [
          {
            id: call.id,
            name: call.name,
            args: call.args,
            type: 'tool_call',
          },
        ],
      }),
    ],
  })
  const messages = readToolNodeMessages(result, call.name)
  const toolMessages = messages.filter((item): item is ToolMessage => item instanceof ToolMessage)

  if (toolMessages.length !== 1) {
    throw new Error(`Expected one ToolMessage for ${call.name}, received ${toolMessages.length}`)
  }

  return toolMessages[0]
}

function readToolNodeMessages(value: unknown, toolName: string): readonly unknown[] {
  const record = requireRecord(value, `${toolName} ToolNode result`)
  const messages = record['messages']
  if (!Array.isArray(messages)) {
    throw new Error(`Expected ${toolName} ToolNode result to include messages`)
  }
  return messages
}

function requireToolName(discoveredTools: readonly string[], name: string): string {
  if (discoveredTools.includes(name)) {
    return name
  }
  const suffix = `__${name}`
  const match = discoveredTools.filter((item) => item.endsWith(suffix))
  if (match.length === 1) {
    return match[0]
  }
  throw new Error(`Missing MCP tool ${name}; discovered ${discoveredTools.join(', ')}`)
}

function parseToolMessageJson(message: ToolMessage): JsonObject {
  const content = message.content
  if (typeof content === 'string') {
    return unpackMcpToolContent(JSON.parse(content), messageName(message))
  }
  if (Array.isArray(content)) {
    for (const block of content) {
      if (isRecord(block) && block['type'] === 'text' && typeof block['text'] === 'string') {
        return unpackMcpToolContent(JSON.parse(block['text']), messageName(message))
      }
    }
  }
  throw new Error(`Expected JSON text ToolMessage content for ${messageName(message)}`)
}

function unpackMcpToolContent(value: unknown, label: string): JsonObject {
  const record = requireRecord(value, label)
  const structuredContent = record['structuredContent']
  if (isRecord(structuredContent)) {
    return structuredContent
  }
  if (typeof record['text'] === 'string') {
    return unpackMcpToolContent(JSON.parse(record['text']), `${label}.text`)
  }
  return record
}

function readCellNumber(cell: JsonObject, label: string): number {
  const value = cell['value']
  if (typeof value === 'number') {
    return requireNumber(value, label)
  }
  if (isRecord(value)) {
    return requireNumber(value['value'], label)
  }
  return requireNumber(value, label)
}

function messageName(message: ToolMessage): string {
  const value = Reflect.get(message, 'name')
  return typeof value === 'string' ? value : 'unknown_tool'
}

function requireRecord(value: unknown, label: string): JsonObject {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object`)
  }
  return value
}

function requireString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string`)
  }
  return value
}

function requireNumber(value: unknown, label: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`Expected ${label} to be a finite number, got ${JSON.stringify(value)}`)
  }
  return value
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
