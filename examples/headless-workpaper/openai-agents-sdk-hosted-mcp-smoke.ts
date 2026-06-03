import { Agent, MCPServerStreamableHttp, RunContext, getAllMcpTools, invokeFunctionTool } from '@openai/agents'

type JsonObject = { readonly [key: string]: unknown }
type OpenAiMcpFunctionTool = Extract<Awaited<ReturnType<typeof getAllMcpTools>>[number], { readonly type: 'function' }>

type MappedMcpTextContent = {
  readonly type: 'text'
  readonly text: string
}

type WorkPaperHostedMcpWriteResult = {
  readonly editedCell: 'Inputs!B3'
  readonly readbackRange: 'Summary!A1:B4'
  readonly beforeExpectedArr: number
  readonly afterExpectedArr: number
  readonly restoredExpectedArr: number
  readonly persistence: {
    readonly persisted: boolean
    readonly serializedBytes: number
  }
  readonly checks: {
    readonly persisted: boolean
    readonly readbackChanged: boolean
    readonly restoredReadbackMatchesAfter: boolean
    readonly previousSerialized: unknown
    readonly newSerialized: unknown
  }
}

type OpenAiAgentsSdkHostedMcpProof = {
  readonly apiShape: 'OpenAI Agents SDK Agent -> MCPServerStreamableHttp -> getAllMcpTools() -> invokeFunctionTool()'
  readonly package: '@openai/agents'
  readonly agentName: string
  readonly mcpServerName: string
  readonly remoteEndpoint: 'https://bilig.proompteng.ai/mcp'
  readonly transport: 'streamable-http'
  readonly stateless: true
  readonly rawMcpToolNames: readonly string[]
  readonly functionToolNames: readonly string[]
  readonly writeResult: WorkPaperHostedMcpWriteResult
  readonly finalText: string
}

const expectedHostedToolNames = [
  'list_sheets',
  'read_range',
  'read_cell',
  'set_cell_contents',
  'set_cell_contents_and_readback',
  'get_cell_display_value',
  'export_workpaper_document',
  'validate_formula',
] as const

const server = new MCPServerStreamableHttp({
  url: 'https://bilig.proompteng.ai/mcp',
  name: 'bilig-workpaper-hosted',
  cacheToolsList: false,
  timeout: 15_000,
})

try {
  await server.connect()

  const agent = new Agent({
    name: 'WorkPaper hosted MCP verification agent',
    instructions:
      'Use hosted Bilig WorkPaper MCP tools for stateless workbook reads and edits. Answer only from computed readback returned by the MCP tool.',
    mcpServers: [server],
  })

  const runContext = new RunContext()
  const rawMcpTools = await server.listTools()
  const functionTools = await getAllMcpTools({
    mcpServers: [server],
    runContext,
    agent,
    convertSchemasToStrict: true,
  })
  const writeTool = requireFunctionTool(functionTools, 'set_cell_contents_and_readback')
  const writeResult = readWorkPaperHostedMcpWriteResult(
    await invokeFunctionTool({
      tool: writeTool,
      runContext,
      input: JSON.stringify({
        sheetName: 'Inputs',
        address: 'B3',
        value: '=0.4',
        readbackRange: 'Summary!A1:B4',
      }),
    }),
  )

  const proof: OpenAiAgentsSdkHostedMcpProof = {
    apiShape: 'OpenAI Agents SDK Agent -> MCPServerStreamableHttp -> getAllMcpTools() -> invokeFunctionTool()',
    package: '@openai/agents',
    agentName: agent.name,
    mcpServerName: server.name,
    remoteEndpoint: 'https://bilig.proompteng.ai/mcp',
    transport: 'streamable-http',
    stateless: true,
    rawMcpToolNames: rawMcpTools.map((tool) => tool.name),
    functionToolNames: functionTools.map((tool) => tool.name),
    writeResult,
    finalText: `Hosted MCP edited ${writeResult.editedCell}; expected ARR changed from ${writeResult.beforeExpectedArr} to ${writeResult.afterExpectedArr}.`,
  }

  assertOpenAiAgentsSdkHostedMcpProof(proof)
  console.log(JSON.stringify(proof, null, 2))
} finally {
  await server.close()
}

function requireFunctionTool(functionTools: Awaited<ReturnType<typeof getAllMcpTools>>, name: string): OpenAiMcpFunctionTool {
  const found = functionTools.find((tool) => tool.name === name)
  if (found === undefined || found.type !== 'function') {
    throw new Error(`Missing OpenAI Agents SDK hosted MCP function tool ${name}: ${JSON.stringify(functionTools.map((tool) => tool.name))}`)
  }
  return found
}

function readWorkPaperHostedMcpWriteResult(output: unknown): WorkPaperHostedMcpWriteResult {
  const content = readMcpTextContent(output)
  const parsed = readRecord(JSON.parse(content.text), 'MCP text payload')
  const persistence = readRecord(parsed.persistence, 'persistence')
  const checks = readRecord(parsed.checks, 'checks')

  return {
    editedCell: readLiteral(parsed.editedCell, 'Inputs!B3', 'editedCell'),
    readbackRange: readLiteral(parsed.readbackRange, 'Summary!A1:B4', 'readbackRange'),
    beforeExpectedArr: readReadbackNumber(parsed.beforeReadback, 2, 1, 'beforeReadback Summary!B3'),
    afterExpectedArr: readReadbackNumber(parsed.afterReadback, 2, 1, 'afterReadback Summary!B3'),
    restoredExpectedArr: readReadbackNumber(parsed.restoredReadback, 2, 1, 'restoredReadback Summary!B3'),
    persistence: {
      persisted: readBoolean(persistence.persisted, 'persistence.persisted'),
      serializedBytes: readNumber(persistence.serializedBytes, 'persistence.serializedBytes'),
    },
    checks: {
      persisted: readBoolean(checks.persisted, 'checks.persisted'),
      readbackChanged: readBoolean(checks.readbackChanged, 'checks.readbackChanged'),
      restoredReadbackMatchesAfter: readBoolean(checks.restoredReadbackMatchesAfter, 'checks.restoredReadbackMatchesAfter'),
      previousSerialized: checks.previousSerialized,
      newSerialized: checks.newSerialized,
    },
  }
}

function readMcpTextContent(output: unknown): MappedMcpTextContent {
  if (Array.isArray(output)) {
    if (output.length !== 1) {
      throw new Error(`Expected one MCP content item, received ${JSON.stringify(output)}`)
    }
    return readMcpTextContent(output[0])
  }

  if (
    typeof output === 'object' &&
    output !== null &&
    'type' in output &&
    'text' in output &&
    output.type === 'text' &&
    typeof output.text === 'string'
  ) {
    return {
      type: output.type,
      text: output.text,
    }
  }

  throw new Error(`Expected MCP text content, received ${JSON.stringify(output)}`)
}

function assertOpenAiAgentsSdkHostedMcpProof(proof: OpenAiAgentsSdkHostedMcpProof): void {
  if (proof.apiShape !== 'OpenAI Agents SDK Agent -> MCPServerStreamableHttp -> getAllMcpTools() -> invokeFunctionTool()') {
    throw new Error('Unexpected OpenAI Agents SDK hosted MCP API shape')
  }

  if (!sameJson(proof.rawMcpToolNames, expectedHostedToolNames)) {
    throw new Error(`Unexpected hosted raw MCP tool names: ${JSON.stringify(proof.rawMcpToolNames)}`)
  }

  if (!sameJson(proof.functionToolNames, expectedHostedToolNames)) {
    throw new Error(`Unexpected hosted function tool names: ${JSON.stringify(proof.functionToolNames)}`)
  }

  if (
    proof.writeResult.editedCell !== 'Inputs!B3' ||
    proof.writeResult.readbackRange !== 'Summary!A1:B4' ||
    proof.writeResult.beforeExpectedArr !== 60000 ||
    proof.writeResult.afterExpectedArr !== 96000 ||
    proof.writeResult.restoredExpectedArr !== 96000 ||
    proof.writeResult.persistence.persisted ||
    proof.writeResult.checks.persisted ||
    !proof.writeResult.checks.readbackChanged ||
    !proof.writeResult.checks.restoredReadbackMatchesAfter ||
    proof.writeResult.checks.previousSerialized !== 0.25 ||
    proof.writeResult.checks.newSerialized !== '=0.4'
  ) {
    throw new Error(`OpenAI Agents SDK hosted MCP WorkPaper proof failed: ${JSON.stringify(proof)}`)
  }

  if (proof.writeResult.persistence.serializedBytes <= 0) {
    throw new Error(`Hosted MCP proof did not return serialized bytes: ${JSON.stringify(proof.writeResult.persistence)}`)
  }

  if (!proof.finalText.includes('Hosted MCP edited Inputs!B3')) {
    throw new Error(`Unexpected OpenAI Agents SDK hosted MCP final text: ${proof.finalText}`)
  }
}

function readReadbackNumber(value: unknown, row: number, col: number, label: string): number {
  const readback = readRecord(value, label)
  const values = Reflect.get(readback, 'values')
  if (!Array.isArray(values)) {
    throw new Error(`Expected ${label}.values to be an array, received ${JSON.stringify(values)}`)
  }
  const readbackRow = values[row]
  if (!Array.isArray(readbackRow)) {
    throw new Error(`Expected ${label} row ${row} to be an array, received ${JSON.stringify(readbackRow)}`)
  }
  const cell = readRecord(readbackRow[col], label)
  return readNumber(cell.value, `${label}.value`)
}

function readLiteral<const T extends string>(value: unknown, expected: T, label: string): T {
  if (value !== expected) {
    throw new Error(`Expected ${label} ${expected}, received ${JSON.stringify(value)}`)
  }
  return expected
}

function readRecord(value: unknown, label: string): JsonObject {
  if (!isJsonObject(value)) {
    throw new Error(`Expected ${label} to be an object, received ${JSON.stringify(value)}`)
  }
  return value
}

function isJsonObject(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(value)}`)
  }
  return Math.round(value * 100) / 100
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${label} to be a boolean, received ${JSON.stringify(value)}`)
  }
  return value
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
