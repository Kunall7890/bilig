import readline from 'node:readline'
import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import type {
  CodexDynamicToolCallRequest,
  CodexDynamicToolCallResult,
  CodexDynamicToolSpec,
  CodexInitializeCapabilities,
  CodexInitializeResponse,
  CodexJsonRpcResponse,
  CodexRequestId,
  CodexServerNotification,
  CodexThread,
  CodexTurn,
  CodexUserInput,
} from '@bilig/agent-api'
import {
  expectInitializeResponse,
  expectThreadStartResponse,
  expectTurnStartResponse,
  isDynamicToolCallServerRequest,
  parseJsonRpcResponse,
  parseServerNotification,
  parseServerRequest,
  type ParsedServerRequest,
} from './codex-app-server-message-parsers.js'

export interface CodexAppServerClientOptions {
  command?: string
  args?: string[]
  cwd?: string
  env?: NodeJS.ProcessEnv
  onLog?: (message: string) => void
  handleDynamicToolCall: (request: CodexDynamicToolCallRequest) => Promise<CodexDynamicToolCallResult>
}

export type CodexAppServerJsonValue =
  | boolean
  | number
  | string
  | null
  | CodexAppServerJsonValue[]
  | { [key: string]: CodexAppServerJsonValue }

export type CodexAppServerApprovalPolicy =
  | 'untrusted'
  | 'on-failure'
  | 'on-request'
  | 'never'
  | {
      granular: {
        mcp_elicitations: boolean
        request_permissions?: boolean
        rules: boolean
        sandbox_approval: boolean
        skill_approval?: boolean
      }
    }
export type CodexAppServerSandboxMode = 'read-only' | 'workspace-write' | 'danger-full-access'
export type CodexAppServerWebSearchMode = 'live' | 'disabled' | 'off' | boolean
export interface CodexAppServerToolsConfig {
  readonly view_image?: boolean
}
export type CodexAppServerThreadConfig = {
  readonly approval_policy?: CodexAppServerApprovalPolicy
  readonly sandbox_mode?: CodexAppServerSandboxMode
  readonly network_access?: boolean
  readonly web_search?: CodexAppServerWebSearchMode
  readonly tools?: CodexAppServerToolsConfig
} & { readonly [key: string]: CodexAppServerJsonValue | CodexAppServerToolsConfig | undefined }

export interface CodexAppServerTransport {
  ensureReady(): Promise<CodexInitializeResponse>
  subscribe(listener: (notification: CodexServerNotification) => void): () => void
  threadStart(input: {
    model: string
    approvalPolicy: CodexAppServerApprovalPolicy
    sandbox: CodexAppServerSandboxMode
    config?: CodexAppServerThreadConfig
    baseInstructions: string
    developerInstructions: string
    dynamicTools: readonly CodexDynamicToolSpec[]
  }): Promise<CodexThread>
  threadResume(input: { threadId: string; baseInstructions: string; developerInstructions: string }): Promise<CodexThread>
  turnStart(input: { threadId: string; prompt: string }): Promise<CodexTurn>
  turnInterrupt(threadId: string): Promise<void>
  close(): Promise<void>
}

interface PendingResponse {
  readonly resolve: (value: unknown) => void
  readonly reject: (error: Error) => void
}

const JSON_RPC_INTERNAL_ERROR = -32603
const JSON_RPC_METHOD_NOT_FOUND = -32601
const CODEX_INITIALIZE_CAPABILITIES: CodexInitializeCapabilities = {
  experimentalApi: true,
}
const TELEMETRY_ENV_PREFIXES = ['OTEL_'] as const

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asError(value: unknown, fallback: string): Error {
  return value instanceof Error ? value : new Error(fallback)
}

function stripTelemetryEnv(inputEnv: NodeJS.ProcessEnv | undefined): NodeJS.ProcessEnv {
  const nextEnv: NodeJS.ProcessEnv = {
    ...(inputEnv ?? process.env),
    OTEL_SDK_DISABLED: 'true',
  }
  for (const key of Object.keys(nextEnv)) {
    if (key === 'OTEL_SDK_DISABLED') {
      continue
    }
    if (TELEMETRY_ENV_PREFIXES.some((prefix) => key.startsWith(prefix))) {
      delete nextEnv[key]
    }
  }
  return nextEnv
}

export class CodexAppServerClient implements CodexAppServerTransport {
  private readonly command: string
  private readonly args: string[]
  private readonly cwd: string | undefined
  private readonly env: NodeJS.ProcessEnv
  private readonly onLog: ((message: string) => void) | undefined
  private readonly handleDynamicToolCall
  private readonly pending = new Map<CodexRequestId, PendingResponse>()
  private readonly notificationListeners = new Set<(notification: CodexServerNotification) => void>()
  private process: ChildProcessWithoutNullStreams | null = null
  private reader: readline.Interface | null = null
  private nextRequestId = 1
  private initializePromise: Promise<CodexInitializeResponse> | null = null

  constructor(options: CodexAppServerClientOptions) {
    this.command = options.command ?? 'codex'
    this.args = options.args ?? ['app-server']
    this.cwd = options.cwd
    this.env = stripTelemetryEnv(options.env)
    this.onLog = options.onLog
    this.handleDynamicToolCall = options.handleDynamicToolCall
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.notificationListeners.add(listener)
    return () => {
      this.notificationListeners.delete(listener)
    }
  }

  async ensureReady(): Promise<CodexInitializeResponse> {
    if (this.initializePromise) {
      return await this.initializePromise
    }

    this.initializePromise = this.startProcess()
    try {
      return await this.initializePromise
    } catch (error) {
      this.initializePromise = null
      throw error
    }
  }

  async threadStart(input: {
    model: string
    approvalPolicy: CodexAppServerApprovalPolicy
    sandbox: CodexAppServerSandboxMode
    config?: CodexAppServerThreadConfig
    baseInstructions: string
    developerInstructions: string
    dynamicTools: readonly CodexDynamicToolSpec[]
  }): Promise<CodexThread> {
    await this.ensureReady()
    const result = expectThreadStartResponse(
      await this.request('thread/start', {
        model: input.model,
        approvalPolicy: input.approvalPolicy,
        sandbox: input.sandbox,
        ...(input.config === undefined ? {} : { config: input.config }),
        baseInstructions: input.baseInstructions,
        developerInstructions: input.developerInstructions,
        dynamicTools: [...input.dynamicTools],
        experimentalRawEvents: false,
        persistExtendedHistory: true,
      }),
    )
    return result.thread
  }

  async threadResume(input: { threadId: string; baseInstructions: string; developerInstructions: string }): Promise<CodexThread> {
    await this.ensureReady()
    const result = expectThreadStartResponse(
      await this.request('thread/resume', {
        threadId: input.threadId,
        baseInstructions: input.baseInstructions,
        developerInstructions: input.developerInstructions,
        persistExtendedHistory: true,
      }),
    )
    return result.thread
  }

  async turnStart(input: { threadId: string; prompt: string }): Promise<CodexTurn> {
    await this.ensureReady()
    const result = expectTurnStartResponse(
      await this.request('turn/start', {
        threadId: input.threadId,
        input: [
          {
            type: 'text',
            text: input.prompt,
          } satisfies CodexUserInput,
        ],
      }),
    )
    return result.turn
  }

  async turnInterrupt(threadId: string): Promise<void> {
    await this.ensureReady()
    await this.request('turn/interrupt', { threadId })
  }

  async close(): Promise<void> {
    const activeProcess = this.process
    this.process = null
    this.reader?.close()
    this.reader = null
    this.initializePromise = null
    this.rejectAllPending(new Error('Codex app-server client closed.'))
    if (!activeProcess) {
      return
    }
    if (!activeProcess.killed) {
      activeProcess.kill('SIGTERM')
    }
    await new Promise<void>((resolve) => {
      activeProcess.once('close', () => resolve())
      activeProcess.once('error', () => resolve())
      setTimeout(resolve, 1_000)
    })
  }

  private async startProcess(): Promise<CodexInitializeResponse> {
    const child = spawn(this.command, this.args, {
      stdio: ['pipe', 'pipe', 'pipe'],
      ...(this.cwd ? { cwd: this.cwd } : {}),
      env: this.env,
    })
    this.process = child

    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      this.onLog?.(chunk.trim())
    })

    child.once('error', (error) => {
      this.handleProcessFailure(child, error instanceof Error ? error : new Error(String(error)))
    })
    child.stdin.once('error', (error) => {
      this.handleProcessFailure(child, error instanceof Error ? error : new Error(String(error)))
    })
    child.once('close', (code, signal) => {
      this.handleProcessFailure(child, new Error(`Codex app-server exited unexpectedly (${signal ?? 'code'}:${String(code ?? 'unknown')})`))
    })

    this.reader = readline.createInterface({ input: child.stdout })
    this.reader.on('line', (line) => {
      if (line.trim().length === 0) {
        return
      }
      void this.handleLine(line)
    })

    const initialized = expectInitializeResponse(
      await this.request('initialize', {
        clientInfo: {
          name: 'monolith',
          title: 'Bilig Monolith',
          version: '0.1.0',
        },
        capabilities: CODEX_INITIALIZE_CAPABILITIES,
      }),
    )
    this.notify('initialized', {})
    return initialized
  }

  private async handleLine(line: string): Promise<void> {
    let message: unknown
    try {
      message = JSON.parse(line)
    } catch (error) {
      this.onLog?.(`Failed to parse Codex app-server message: ${String(error)}`)
      return
    }

    if (!isRecord(message)) {
      return
    }

    const response = parseJsonRpcResponse(message)
    if (response) {
      this.handleResponse(response)
      return
    }

    const request = parseServerRequest(message)
    if (request) {
      await this.handleServerRequest(request)
      return
    }

    const notification = parseServerNotification(message)
    if (notification) {
      this.emitNotification(notification)
    }
  }

  private handleResponse(message: CodexJsonRpcResponse<unknown>): void {
    const pending = this.pending.get(message.id)
    if (!pending) {
      return
    }
    this.pending.delete(message.id)
    if (message.error) {
      pending.reject(new Error(message.error.message))
      return
    }
    pending.resolve(message.result)
  }

  private async handleServerRequest(request: ParsedServerRequest): Promise<void> {
    try {
      if (!isDynamicToolCallServerRequest(request)) {
        this.respondWithError(request.id, JSON_RPC_METHOD_NOT_FOUND, `Unsupported server request: ${request.method}`)
        return
      }
      const result = await this.handleDynamicToolCall(request.params)
      this.respondWithResult(request.id, result)
    } catch (error) {
      this.respondWithError(request.id, JSON_RPC_INTERNAL_ERROR, asError(error, 'Dynamic tool call failed').message)
    }
  }

  private emitNotification(notification: CodexServerNotification): void {
    this.notificationListeners.forEach((listener) => {
      listener(notification)
    })
  }

  private async request(method: string, params: unknown): Promise<unknown> {
    const activeProcess = this.process
    if (!activeProcess) {
      throw new Error('Codex app-server process is not running')
    }
    const id = this.nextRequestId++
    const response = new Promise<unknown>((resolve, reject) => {
      this.pending.set(id, {
        resolve,
        reject,
      })
    })
    try {
      this.write({
        method,
        id,
        params,
      })
    } catch (error) {
      this.pending.delete(id)
      throw asError(error, 'Failed to write Codex app-server request')
    }
    return await response
  }

  private notify(method: string, params: unknown): void {
    this.write({
      method,
      params,
    })
  }

  private write(payload: Record<string, unknown>): void {
    const activeProcess = this.process
    if (!activeProcess) {
      throw new Error('Codex app-server process is not running')
    }
    activeProcess.stdin.write(`${JSON.stringify(payload)}\n`)
  }

  private respondWithResult(id: CodexRequestId, result: CodexDynamicToolCallResult): void {
    this.write({
      id,
      result,
    })
  }

  private respondWithError(id: CodexRequestId, code: number, message: string): void {
    this.write({
      id,
      error: {
        code,
        message,
      },
    })
  }

  private rejectAllPending(error: Error): void {
    if (this.pending.size === 0) {
      return
    }
    const entries = [...this.pending.values()]
    this.pending.clear()
    entries.forEach((pending) => {
      pending.reject(error)
    })
  }

  private handleProcessFailure(child: ChildProcessWithoutNullStreams, error: Error): void {
    if (this.process !== child) {
      return
    }
    this.process = null
    this.reader?.close()
    this.reader = null
    this.initializePromise = null
    this.rejectAllPending(error)
  }
}
