import { buildDemoWorkPaper, createWorkPaperMcpToolServer, type WorkPaperMcpToolServer } from './work-paper-mcp-server.js'
import type { WorkPaper } from './work-paper.js'
import type { Readable, Writable } from 'node:stream'

interface WorkPaperMcpStdioOptions {
  input?: Readable
  output?: Writable
  server?: WorkPaperMcpToolServer
  workbook?: WorkPaper
  serverName?: string
  serverVersion?: string
}

function runDemoWorkPaperMcpStdioServer(options: WorkPaperMcpStdioOptions = {}): void {
  const input = options.input ?? process.stdin
  const output = options.output ?? process.stdout
  const server = options.server ?? createWorkPaperMcpToolServer(options.workbook ?? buildDemoWorkPaper())
  const serverName = options.serverName ?? 'bilig-headless-workpaper'
  const serverVersion = options.serverVersion ?? '0.1.0'
  let inputBuffer = ''

  input.setEncoding('utf8')
  input.on('data', (chunk: string) => {
    inputBuffer += chunk
    drainInputLines(false)
  })
  input.on('end', () => {
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
    let request: unknown

    try {
      request = JSON.parse(line)
    } catch (error) {
      writeJsonRpcError(null, -32700, `Parse error: ${errorMessage(error)}`)
      return
    }

    if (!isJsonRpcRequest(request)) {
      const id = isRecord(request) ? request['id'] : null
      writeJsonRpcError(isJsonRpcId(id) ? id : null, -32600, 'Invalid JSON-RPC 2.0 request')
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

  function dispatchJsonRpc(request: JsonRpcRequest): unknown {
    if (request.method === 'initialize') {
      return {
        jsonrpc: '2.0',
        id: request.id,
        result: {
          protocolVersion: '2025-06-18',
          capabilities: server.capabilities,
          serverInfo: {
            name: serverName,
            version: serverVersion,
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

  function writeJson(value: unknown): void {
    output.write(`${JSON.stringify(value)}\n`)
  }
}

type JsonRpcId = string | number | null

interface JsonRpcRequest {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  method: string
  params?: Record<string, unknown>
}

function isJsonRpcRequest(value: unknown): value is JsonRpcRequest {
  if (!isRecord(value) || value['jsonrpc'] !== '2.0' || typeof value['method'] !== 'string') {
    return false
  }

  return value['id'] === undefined || isJsonRpcId(value['id'])
}

function isJsonRpcId(value: unknown): value is JsonRpcId {
  return value === null || typeof value === 'string' || typeof value === 'number'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export { runDemoWorkPaperMcpStdioServer, type WorkPaperMcpStdioOptions }
