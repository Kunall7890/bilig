import type {
  CodexDynamicToolCallRequest,
  CodexDynamicToolCallResult,
  CodexInitializeResponse,
  CodexJsonRpcError,
  CodexJsonRpcResponse,
  CodexRequestId,
  CodexServerNotification,
  CodexThread,
  CodexThreadItem,
  CodexThreadStartResponse,
  CodexTurn,
  CodexTurnStartResponse,
  CodexUserInput,
  JsonValue,
} from '@bilig/agent-api'

type ParsedJsonValue = JsonValue

type ParsedThreadItem = CodexThreadItem

export type ParsedServerRequest =
  | {
      method: 'item/tool/call'
      id: CodexRequestId
      params: CodexDynamicToolCallRequest
    }
  | {
      method: string
      id: CodexRequestId
      params?: ParsedJsonValue
    }

export function isDynamicToolCallServerRequest(
  request: ParsedServerRequest,
): request is Extract<ParsedServerRequest, { method: 'item/tool/call' }> {
  return (
    request.method === 'item/tool/call' &&
    typeof request.params === 'object' &&
    request.params !== null &&
    'threadId' in request.params &&
    typeof request.params.threadId === 'string'
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isString(value: unknown): value is string {
  return typeof value === 'string'
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isRequestId(value: unknown): value is CodexRequestId {
  return isString(value) || isFiniteNumber(value)
}

function isCodexCommandExecutionStatus(value: unknown): value is 'inProgress' | 'completed' | 'failed' {
  return value === 'inProgress' || value === 'completed' || value === 'failed'
}

function isJsonValue(value: unknown): value is ParsedJsonValue {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return true
  }
  if (Array.isArray(value)) {
    return value.every((entry) => isJsonValue(entry))
  }
  if (!isRecord(value)) {
    return false
  }
  return Object.values(value).every((entry) => isJsonValue(entry))
}

function parseInitializeResponse(value: unknown): CodexInitializeResponse | null {
  if (!isRecord(value)) {
    return null
  }
  const { codexHome, platformFamily, platformOs, userAgent } = value
  if (!isString(userAgent) || !isString(codexHome) || !isString(platformFamily) || !isString(platformOs)) {
    return null
  }
  return {
    userAgent,
    codexHome,
    platformFamily,
    platformOs,
  }
}

function parseTurnError(value: unknown): NonNullable<CodexTurn['error']> | null {
  if (value === null) {
    return null
  }
  if (!isRecord(value) || !isString(value['message'])) {
    return null
  }
  return {
    message: value['message'],
    ...(isString(value['additionalDetails']) ? { additionalDetails: value['additionalDetails'] } : {}),
    ...(isJsonValue(value['codexErrorInfo']) ? { codexErrorInfo: value['codexErrorInfo'] } : {}),
  }
}

function parseToolContentItem(value: unknown): CodexDynamicToolCallResult['contentItems'][number] | null {
  if (!isRecord(value) || !isString(value['type'])) {
    return null
  }
  if (value['type'] === 'inputText' && isString(value['text'])) {
    return {
      type: 'inputText',
      text: value['text'],
    }
  }
  if (value['type'] === 'inputImage' && isString(value['imageUrl'])) {
    return {
      type: 'inputImage',
      imageUrl: value['imageUrl'],
    }
  }
  return null
}

function parseUserInput(value: unknown): CodexUserInput | null {
  if (!isRecord(value) || !isString(value['type'])) {
    return null
  }
  switch (value['type']) {
    case 'text': {
      if (!isString(value['text'])) {
        return null
      }
      const textElements = value['text_elements']
      if (textElements !== undefined && (!Array.isArray(textElements) || !textElements.every((entry) => isJsonValue(entry)))) {
        return null
      }
      return {
        type: 'text',
        text: value['text'],
        ...(textElements === undefined ? {} : { text_elements: textElements }),
      }
    }
    case 'image':
      return isString(value['url'])
        ? {
            type: 'image',
            url: value['url'],
          }
        : null
    case 'localImage':
      return isString(value['path'])
        ? {
            type: 'localImage',
            path: value['path'],
          }
        : null
    case 'skill':
      return isString(value['name']) && isString(value['path'])
        ? {
            type: 'skill',
            name: value['name'],
            path: value['path'],
          }
        : null
    case 'mention':
      return isString(value['name']) && isString(value['path'])
        ? {
            type: 'mention',
            name: value['name'],
            path: value['path'],
          }
        : null
    default:
      return null
  }
}

function parseThreadItem(value: unknown): ParsedThreadItem | null {
  if (!isRecord(value) || !isString(value['type']) || !isString(value['id'])) {
    return null
  }
  const type = value['type']
  const id = value['id']
  switch (type) {
    case 'userMessage': {
      if (!Array.isArray(value['content'])) {
        return null
      }
      const content: CodexUserInput[] = []
      for (const entry of value['content']) {
        const item = parseUserInput(entry)
        if (!item) {
          return null
        }
        content.push(item)
      }
      return {
        type,
        id,
        content,
      }
    }
    case 'agentMessage': {
      if (!isString(value['text'])) {
        return null
      }
      const phase = value['phase']
      if (phase !== undefined && phase !== null && !isString(phase)) {
        return null
      }
      return {
        type,
        id,
        text: value['text'],
        phase: phase ?? null,
        memoryCitation: isJsonValue(value['memoryCitation']) ? value['memoryCitation'] : null,
      }
    }
    case 'plan': {
      if (!isString(value['text'])) {
        return null
      }
      return {
        type,
        id,
        text: value['text'],
      }
    }
    case 'dynamicToolCall': {
      const namespace = value['namespace']
      const success = value['success']
      const durationMs = value['durationMs']
      if (
        !isString(value['tool']) ||
        !isJsonValue(value['arguments']) ||
        (value['status'] !== 'inProgress' && value['status'] !== 'completed' && value['status'] !== 'failed') ||
        (namespace !== undefined && namespace !== null && !isString(namespace)) ||
        (success !== undefined && success !== null && typeof success !== 'boolean') ||
        (durationMs !== undefined && durationMs !== null && !isFiniteNumber(durationMs))
      ) {
        return null
      }
      const contentItemsValue = value['contentItems']
      if (contentItemsValue !== undefined && contentItemsValue !== null && !Array.isArray(contentItemsValue)) {
        return null
      }
      const contentItems =
        contentItemsValue === undefined || contentItemsValue === null ? null : contentItemsValue.map((entry) => parseToolContentItem(entry))
      if (contentItems && contentItems.some((entry) => entry === null)) {
        return null
      }
      return {
        type,
        id,
        tool: value['tool'],
        arguments: value['arguments'],
        namespace: namespace ?? null,
        status: value['status'],
        contentItems,
        success: success ?? null,
        durationMs: durationMs ?? null,
      }
    }
    case 'commandExecution': {
      const processId = value['processId']
      const aggregatedOutput = value['aggregatedOutput']
      const exitCode = value['exitCode']
      const durationMs = value['durationMs']
      if (
        !isString(value['command']) ||
        !isString(value['cwd']) ||
        (processId !== undefined && processId !== null && !isString(processId)) ||
        !isCodexCommandExecutionStatus(value['status']) ||
        !Array.isArray(value['commandActions']) ||
        !value['commandActions'].every((entry) => isJsonValue(entry)) ||
        (aggregatedOutput !== undefined && aggregatedOutput !== null && !isString(aggregatedOutput)) ||
        (exitCode !== undefined && exitCode !== null && !isFiniteNumber(exitCode)) ||
        (durationMs !== undefined && durationMs !== null && !isFiniteNumber(durationMs))
      ) {
        return null
      }
      return {
        type,
        id,
        command: value['command'],
        cwd: value['cwd'],
        processId: processId ?? null,
        status: value['status'],
        commandActions: value['commandActions'],
        aggregatedOutput: aggregatedOutput ?? null,
        exitCode: exitCode ?? null,
        durationMs: durationMs ?? null,
      }
    }
    default: {
      const additionalEntries: Record<string, ParsedJsonValue | undefined> = {}
      for (const [key, entry] of Object.entries(value)) {
        if (key === 'type' || key === 'id') {
          continue
        }
        if (entry !== undefined && !isJsonValue(entry)) {
          return null
        }
        additionalEntries[key] = entry
      }
      return {
        type,
        id,
        ...additionalEntries,
      }
    }
  }
}

function parseTurn(value: unknown): CodexTurn | null {
  if (!isRecord(value) || !isString(value['id']) || !Array.isArray(value['items'])) {
    return null
  }
  const status = value['status']
  if (status !== 'completed' && status !== 'interrupted' && status !== 'failed' && status !== 'inProgress') {
    return null
  }
  const items: ParsedThreadItem[] = []
  for (const entry of value['items']) {
    const item = parseThreadItem(entry)
    if (!item) {
      return null
    }
    items.push(item)
  }
  const error = parseTurnError(value['error'])
  if (error === null && value['error'] !== undefined && value['error'] !== null) {
    return null
  }
  return {
    id: value['id'],
    status,
    items,
    error,
  }
}

function parseThread(value: unknown): CodexThread | null {
  if (!isRecord(value) || !isString(value['id']) || !isString(value['preview'])) {
    return null
  }
  if (!Array.isArray(value['turns'])) {
    return null
  }
  const turns: CodexTurn[] = []
  for (const entry of value['turns']) {
    const turn = parseTurn(entry)
    if (!turn) {
      return null
    }
    turns.push(turn)
  }
  return {
    id: value['id'],
    preview: value['preview'],
    turns,
  }
}

function parseThreadStartResponse(value: unknown): CodexThreadStartResponse | null {
  if (!isRecord(value)) {
    return null
  }
  const thread = parseThread(value['thread'])
  return thread ? { thread } : null
}

function parseTurnStartResponse(value: unknown): CodexTurnStartResponse | null {
  if (!isRecord(value)) {
    return null
  }
  const turn = parseTurn(value['turn'])
  return turn ? { turn } : null
}

function parseJsonRpcError(value: unknown): CodexJsonRpcError | null {
  if (!isRecord(value) || !isFiniteNumber(value['code']) || !isString(value['message'])) {
    return null
  }
  if (value['data'] !== undefined && !isJsonValue(value['data'])) {
    return null
  }
  return {
    code: value['code'],
    message: value['message'],
    ...(value['data'] !== undefined ? { data: value['data'] } : {}),
  }
}

export function parseJsonRpcResponse(value: unknown): CodexJsonRpcResponse<unknown> | null {
  if (!isRecord(value) || !isRequestId(value['id'])) {
    return null
  }
  const hasResult = Object.hasOwn(value, 'result')
  const hasError = Object.hasOwn(value, 'error')
  if (!hasResult && !hasError) {
    return null
  }
  const response: CodexJsonRpcResponse<unknown> = {
    id: value['id'],
  }
  if (hasResult) {
    response.result = value['result']
  }
  if (hasError) {
    const error = parseJsonRpcError(value['error'])
    if (!error) {
      return null
    }
    response.error = error
  }
  return response
}

function parseDynamicToolCallRequest(value: unknown): CodexDynamicToolCallRequest | null {
  const namespace = isRecord(value) ? value['namespace'] : undefined
  if (
    !isRecord(value) ||
    !isString(value['threadId']) ||
    !isString(value['turnId']) ||
    !isString(value['callId']) ||
    !isString(value['tool']) ||
    !isJsonValue(value['arguments']) ||
    (namespace !== undefined && namespace !== null && !isString(namespace))
  ) {
    return null
  }
  return {
    threadId: value['threadId'],
    turnId: value['turnId'],
    callId: value['callId'],
    tool: value['tool'],
    arguments: value['arguments'],
    namespace: namespace ?? null,
  }
}

export function parseServerRequest(value: unknown): ParsedServerRequest | null {
  if (!isRecord(value) || !isRequestId(value['id']) || !isString(value['method'])) {
    return null
  }
  if (value['method'] === 'item/tool/call') {
    const params = parseDynamicToolCallRequest(value['params'])
    return params
      ? {
          method: 'item/tool/call',
          id: value['id'],
          params,
        }
      : null
  }
  if (value['params'] !== undefined && !isJsonValue(value['params'])) {
    return null
  }
  return {
    method: value['method'],
    id: value['id'],
    ...(value['params'] !== undefined ? { params: value['params'] } : {}),
  }
}

export function parseServerNotification(value: unknown): CodexServerNotification | null {
  if (!isRecord(value) || !isString(value['method']) || !isRecord(value['params'])) {
    return null
  }
  const method = value['method']
  const params = value['params']
  switch (method) {
    case 'thread/started': {
      const thread = parseThread(params['thread'])
      return thread ? { method, params: { thread } } : null
    }
    case 'turn/started':
    case 'turn/completed': {
      if (!isString(params['threadId'])) {
        return null
      }
      const turn = parseTurn(params['turn'])
      return turn
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turn,
            },
          }
        : null
    }
    case 'item/started':
    case 'item/completed': {
      if (!isString(params['threadId']) || !isString(params['turnId'])) {
        return null
      }
      const item = parseThreadItem(params['item'])
      return item
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turnId: params['turnId'],
              item,
            },
          }
        : null
    }
    case 'item/agentMessage/delta':
    case 'item/plan/delta':
    case 'item/reasoning/delta':
    case 'item/reasoning/textDelta':
    case 'item/reasoning/summaryTextDelta':
      return isString(params['threadId']) && isString(params['turnId']) && isString(params['itemId']) && isString(params['delta'])
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turnId: params['turnId'],
              itemId: params['itemId'],
              delta: params['delta'],
            },
          }
        : null
    case 'item/commandExecution/outputDelta':
      return isString(params['threadId']) && isString(params['turnId']) && isString(params['itemId']) && isString(params['delta'])
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turnId: params['turnId'],
              itemId: params['itemId'],
              delta: params['delta'],
            },
          }
        : null
    case 'item/commandExecution/terminalInteraction': {
      const processId = params['processId']
      return isString(params['threadId']) &&
        isString(params['turnId']) &&
        isString(params['itemId']) &&
        (processId === undefined || processId === null || isString(processId)) &&
        isString(params['stdin'])
        ? {
            method,
            params: {
              threadId: params['threadId'],
              turnId: params['turnId'],
              itemId: params['itemId'],
              processId: processId ?? null,
              stdin: params['stdin'],
            },
          }
        : null
    }
    case 'error': {
      if (params['message'] !== undefined && !isString(params['message'])) {
        return null
      }
      const extraParams: Record<string, ParsedJsonValue | undefined> = {}
      for (const [key, entry] of Object.entries(params)) {
        if (entry !== undefined && !isJsonValue(entry)) {
          return null
        }
        extraParams[key] = entry
      }
      return {
        method,
        params: extraParams,
      }
    }
    default:
      return null
  }
}

export function expectInitializeResponse(value: unknown): CodexInitializeResponse {
  const response = parseInitializeResponse(value)
  if (!response) {
    throw new Error('Invalid Codex initialize response')
  }
  return response
}

export function expectThreadStartResponse(value: unknown): CodexThreadStartResponse {
  const response = parseThreadStartResponse(value)
  if (!response) {
    throw new Error('Invalid Codex thread response')
  }
  return response
}

export function expectTurnStartResponse(value: unknown): CodexTurnStartResponse {
  const response = parseTurnStartResponse(value)
  if (!response) {
    throw new Error('Invalid Codex turn response')
  }
  return response
}
