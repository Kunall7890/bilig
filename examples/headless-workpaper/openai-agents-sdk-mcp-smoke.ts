import { Agent, MCPServerStdio, RunContext, getAllMcpTools, invokeFunctionTool } from '@openai/agents'

type JsonObject = { readonly [key: string]: unknown }
type OpenAiMcpFunctionTool = Extract<Awaited<ReturnType<typeof getAllMcpTools>>[number], { readonly type: 'function' }>

type MappedMcpTextContent = {
  readonly type: 'text'
  readonly text: string
}

type WorkPaperMcpWriteResult = {
  readonly editedCell: 'Inputs!B3'
  readonly before: {
    readonly expectedArr: number
    readonly targetGap: number
  }
  readonly after: {
    readonly expectedArr: number
    readonly targetGap: number
  }
  readonly restored: {
    readonly expectedArr: number
    readonly targetGap: number
  }
  readonly checks: {
    readonly previousValue: unknown
    readonly newValue: unknown
    readonly formulasPersisted: boolean
    readonly restoredMatchesAfter: boolean
    readonly expectedArrChanged: boolean
  }
}

type OpenAiAgentsSdkMcpProof = {
  readonly apiShape: 'OpenAI Agents SDK Agent -> MCPServerStdio -> getAllMcpTools() -> invokeFunctionTool()'
  readonly package: '@openai/agents'
  readonly agentName: string
  readonly mcpServerName: string
  readonly rawMcpToolNames: readonly string[]
  readonly functionToolNames: readonly string[]
  readonly writeResult: WorkPaperMcpWriteResult
  readonly finalText: string
}

const server = new MCPServerStdio({
  name: 'bilig-workpaper-stdio',
  fullCommand: 'npm run --silent agent:mcp-stdio',
  cwd: process.cwd(),
  cacheToolsList: false,
  timeout: 10_000,
})

try {
  await server.connect()

  const agent = new Agent({
    name: 'WorkPaper MCP verification agent',
    instructions:
      'Use the Bilig WorkPaper MCP tools for workbook reads and edits. Answer only from computed readback returned by the MCP tool.',
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
  const writeTool = requireFunctionTool(functionTools, 'set_workpaper_input_cell')
  const writeResult = readWorkPaperMcpWriteResult(
    await invokeFunctionTool({
      tool: writeTool,
      runContext,
      input: JSON.stringify({
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      }),
    }),
  )

  const proof: OpenAiAgentsSdkMcpProof = {
    apiShape: 'OpenAI Agents SDK Agent -> MCPServerStdio -> getAllMcpTools() -> invokeFunctionTool()',
    package: '@openai/agents',
    agentName: agent.name,
    mcpServerName: server.name,
    rawMcpToolNames: rawMcpTools.map((tool) => tool.name),
    functionToolNames: functionTools.map((tool) => tool.name),
    writeResult,
    finalText: `MCP edited ${writeResult.editedCell}; expected ARR changed from ${writeResult.before.expectedArr} to ${writeResult.after.expectedArr}.`,
  }

  assertOpenAiAgentsSdkMcpProof(proof)
  console.log(JSON.stringify(proof, null, 2))
} finally {
  await server.close()
}

function requireFunctionTool(functionTools: Awaited<ReturnType<typeof getAllMcpTools>>, name: string): OpenAiMcpFunctionTool {
  const found = functionTools.find((tool) => tool.name === name)
  if (found === undefined || found.type !== 'function') {
    throw new Error(`Missing OpenAI Agents SDK MCP function tool ${name}: ${JSON.stringify(functionTools.map((tool) => tool.name))}`)
  }
  return found
}

function readWorkPaperMcpWriteResult(output: unknown): WorkPaperMcpWriteResult {
  const content = readMcpTextContent(output)
  const parsed = readRecord(JSON.parse(content.text), 'MCP text payload')
  const before = readSummary(parsed.before, 'before')
  const after = readSummary(parsed.after, 'after')
  const restored = readSummary(parsed.restored, 'restored')
  const checks = readRecord(parsed.checks, 'checks')

  return {
    editedCell: readEditedCell(parsed.editedCell),
    before,
    after,
    restored,
    checks: {
      previousValue: checks.previousValue,
      newValue: checks.newValue,
      formulasPersisted: readBoolean(checks.formulasPersisted, 'formulasPersisted'),
      restoredMatchesAfter: readBoolean(checks.restoredMatchesAfter, 'restoredMatchesAfter'),
      expectedArrChanged: readBoolean(checks.expectedArrChanged, 'expectedArrChanged'),
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

function assertOpenAiAgentsSdkMcpProof(proof: OpenAiAgentsSdkMcpProof): void {
  if (proof.apiShape !== 'OpenAI Agents SDK Agent -> MCPServerStdio -> getAllMcpTools() -> invokeFunctionTool()') {
    throw new Error('Unexpected OpenAI Agents SDK MCP API shape')
  }

  if (!sameJson(proof.rawMcpToolNames, ['read_workpaper_summary', 'set_workpaper_input_cell'])) {
    throw new Error(`Unexpected raw MCP tool names: ${JSON.stringify(proof.rawMcpToolNames)}`)
  }

  if (!sameJson(proof.functionToolNames, ['read_workpaper_summary', 'set_workpaper_input_cell'])) {
    throw new Error(`Unexpected function tool names: ${JSON.stringify(proof.functionToolNames)}`)
  }

  if (
    proof.writeResult.editedCell !== 'Inputs!B3' ||
    proof.writeResult.before.expectedArr !== 60000 ||
    proof.writeResult.after.expectedArr !== 96000 ||
    proof.writeResult.restored.expectedArr !== 96000 ||
    proof.writeResult.checks.previousValue !== 0.25 ||
    proof.writeResult.checks.newValue !== 0.4 ||
    !proof.writeResult.checks.formulasPersisted ||
    !proof.writeResult.checks.restoredMatchesAfter ||
    !proof.writeResult.checks.expectedArrChanged
  ) {
    throw new Error(`OpenAI Agents SDK MCP WorkPaper proof failed: ${JSON.stringify(proof)}`)
  }

  if (!proof.finalText.includes('MCP edited Inputs!B3')) {
    throw new Error(`Unexpected OpenAI Agents SDK MCP final text: ${proof.finalText}`)
  }
}

function readSummary(value: unknown, label: string): WorkPaperMcpWriteResult['before'] {
  const record = readRecord(value, label)
  return {
    expectedArr: readNumber(record.expectedArr, `${label}.expectedArr`),
    targetGap: readNumber(record.targetGap, `${label}.targetGap`),
  }
}

function readEditedCell(value: unknown): 'Inputs!B3' {
  if (value !== 'Inputs!B3') {
    throw new Error(`Expected edited cell Inputs!B3, received ${JSON.stringify(value)}`)
  }
  return value
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
