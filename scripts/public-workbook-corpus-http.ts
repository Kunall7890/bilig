export interface FetchBodyBytesResult {
  readonly response: Response
  readonly bytes: Uint8Array
}

export interface FetchBodyBytesArgs {
  readonly timeoutMs: number
  readonly maxBytes?: number
  readonly maxBytesLabel?: string
  readonly validateResponse?: (response: Response) => void
}

export async function fetchBodyBytesWithTimeout(
  input: string | URL,
  init: RequestInit,
  args: FetchBodyBytesArgs,
): Promise<FetchBodyBytesResult> {
  return withRequestTimeout(args.timeoutMs, async (signal) => {
    const response = await fetch(input, { ...init, signal })
    args.validateResponse?.(response)
    const bytes = await readResponseBodyBytes(response, signal, args)
    return { response, bytes }
  })
}

export async function fetchJsonWithTimeout(input: string | URL, init: RequestInit, args: FetchBodyBytesArgs): Promise<unknown> {
  const { bytes } = await fetchBodyBytesWithTimeout(input, init, args)
  return JSON.parse(new TextDecoder().decode(bytes)) as unknown
}

async function withRequestTimeout<T>(timeoutMs: number, operation: (signal: AbortSignal) => Promise<T>): Promise<T> {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    const result = await operation(controller.signal)
    if (controller.signal.aborted) {
      throw new Error(requestTimeoutMessage(timeoutMs))
    }
    return result
  } catch (error) {
    if (controller.signal.aborted) {
      throw new Error(requestTimeoutMessage(timeoutMs), { cause: error })
    }
    throw error
  } finally {
    clearTimeout(timer)
  }
}

async function readResponseBodyBytes(response: Response, signal: AbortSignal, args: FetchBodyBytesArgs): Promise<Uint8Array> {
  const body = response.body
  if (!body) {
    const bytes = new Uint8Array(await response.arrayBuffer())
    assertMaxBytes(bytes.byteLength, args)
    return bytes
  }

  const chunks: Uint8Array[] = []
  let byteLength = 0
  const reader = body.getReader()
  const abortPromise = new Promise<never>((_, reject) => {
    signal.addEventListener('abort', () => reject(new Error(requestTimeoutMessage(args.timeoutMs))), { once: true })
  })
  let completed = false
  const readNextChunk = async (): Promise<void> => {
    const { done, value } = await Promise.race([reader.read(), abortPromise])
    if (done) {
      completed = true
      return
    }
    byteLength += value.byteLength
    assertMaxBytes(byteLength, args)
    chunks.push(value)
    await readNextChunk()
  }
  try {
    await readNextChunk()
  } finally {
    if (!completed) {
      await reader.cancel().catch(() => undefined)
    }
    reader.releaseLock()
  }

  const bytes = new Uint8Array(byteLength)
  let offset = 0
  for (const chunk of chunks) {
    bytes.set(chunk, offset)
    offset += chunk.byteLength
  }
  return bytes
}

function assertMaxBytes(byteLength: number, args: FetchBodyBytesArgs): void {
  const maxBytes = args.maxBytes
  if (maxBytes !== undefined && byteLength > maxBytes) {
    throw new Error(`${args.maxBytesLabel ?? 'Response body'} exceeds max byte size: ${String(byteLength)} > ${String(maxBytes)}`)
  }
}

function requestTimeoutMessage(timeoutMs: number): string {
  return `Request timed out after ${String(timeoutMs)}ms`
}
