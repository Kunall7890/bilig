import {
  buildWorkbook,
  createMcpWorkPaperToolServer,
  type JsonRpcId,
  type JsonRpcRequest,
  type McpCapabilities,
  type McpJsonRpcResponse,
} from './mcp-tool-server.ts'

interface InitializeResult {
  protocolVersion: '2025-06-18'
  capabilities: McpCapabilities
  serverInfo: {
    name: 'bilig-headless-workpaper-example'
    version: '0.1.0'
  }
}

interface JsonRpcErrorResponse {
  jsonrpc: '2.0'
  id: JsonRpcId
  error: {
    code: number
    message: string
  }
}

type InitializeResponse = {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  result: InitializeResult
}

type StdioJsonRpcResponse = JsonRpcErrorResponse | McpJsonRpcResponse | InitializeResponse

const server = createMcpWorkPaperToolServer(buildWorkbook())
let inputBuffer = ''

process.stdin.setEncoding('utf8')
process.stdin.on('data', (chunk: string) => {
  inputBuffer += chunk
  drainInputLines(false)
})
process.stdin.on('end', () => {
  drainInputLines(true)
})

function drainInputLines(flush: boolean): void {
  let newlineIndex = inputBuffer.indexOf('\n')
  while (newlineIndex !== -1) {
    const line = inputBuffer.slice(0, newlineIndex).trim()
    inputBuffer = inputBuffer.slice(newlineIndex + 1)
    if (line.length > 0) {
      handleLine(line)
    }
    newlineIndex = inputBuffer.indexOf('\n')
  }

  const trailingLine = inputBuffer.trim()
  if (flush && trailingLine.length > 0) {
    inputBuffer = ''
    handleLine(trailingLine)
  }
}

function handleLine(line: string): void {
  let request: JsonRpcRequest

  try {
    request = parseJsonRpcLine(line)
  } catch (error) {
    writeJsonRpcError(null, -32700, `Parse error: ${errorMessage(error)}`)
    return
  }

  try {
    const response = dispatchJsonRpc(request)
    if (response !== undefined) {
      writeJson(response)
    }
  } catch (error) {
    writeJsonRpcError(request.id ?? null, -32603, errorMessage(error))
  }
}

function parseJsonRpcLine(line: string): JsonRpcRequest {
  const value: unknown = JSON.parse(line)
  if (!isRecord(value)) {
    throw new Error('Invalid JSON-RPC 2.0 request')
  }

  const candidate = value
  if (candidate.jsonrpc !== '2.0' || typeof candidate.method !== 'string') {
    throw new Error('Invalid JSON-RPC 2.0 request')
  }

  const id = candidate.id
  if (id !== undefined && id !== null && typeof id !== 'string' && typeof id !== 'number') {
    throw new Error(`Unsupported JSON-RPC id: ${JSON.stringify(id)}`)
  }

  return {
    jsonrpc: '2.0',
    id,
    method: candidate.method,
    params: parseOptionalRecord(candidate.params, 'params'),
  }
}

function parseOptionalRecord(value: unknown, label: string): Record<string, unknown> | undefined {
  if (value === undefined) {
    return undefined
  }

  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object`)
  }

  return value
}

function dispatchJsonRpc(request: JsonRpcRequest): InitializeResponse | McpJsonRpcResponse | undefined {
  if (request.method === 'initialize') {
    return {
      jsonrpc: '2.0',
      id: request.id,
      result: {
        protocolVersion: '2025-06-18',
        capabilities: server.capabilities,
        serverInfo: {
          name: 'bilig-headless-workpaper-example',
          version: '0.1.0',
        },
      },
    }
  }

  if (request.method === 'notifications/initialized' || request.id === undefined) {
    return undefined
  }

  return server.handleJsonRpc(request)
}

function writeJsonRpcError(id: JsonRpcId, code: number, message: string): void {
  writeJson({
    jsonrpc: '2.0',
    id,
    error: {
      code,
      message,
    },
  })
}

function writeJson(value: StdioJsonRpcResponse): void {
  process.stdout.write(`${JSON.stringify(value)}\n`)
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
