import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import { decodeAgentFrame, encodeAgentFrame } from '@bilig/agent-api'
import type { RuntimeSession } from '@bilig/contracts'
import { createRuntimeSession, type DocumentControlService, resolveRequestBaseUrl, runPromise } from '@bilig/runtime-kernel'
import type { BiligRuntimeConfig } from '@bilig/zero-sync'
import type { WorkbookAgentService } from '../codex-app/workbook-agent-service.js'
import { resolveRequestSession, resolveSessionIdentity } from './session.js'

function resolveBooleanEnv(value: string | undefined, fallback: boolean, name: string): boolean {
  if (value === undefined || value.length === 0) {
    return fallback
  }
  if (value === 'true' || value === '1') {
    return true
  }
  if (value === 'false' || value === '0') {
    return false
  }
  throw new Error(`${name} must be "1", "true", "0", or "false" when set, got ${value}`)
}

function resolveWebRuntimeConfig(env: Record<string, string | undefined>): Omit<BiligRuntimeConfig, 'currentUserId'> {
  const zeroCacheUrl = env['BILIG_ZERO_CACHE_URL']?.trim() || '/zero'
  const defaultDocumentId = env['BILIG_DEFAULT_DOCUMENT_ID']?.trim() || 'bilig-demo'

  return {
    zeroCacheUrl,
    defaultDocumentId,
    persistState: resolveBooleanEnv(env['BILIG_PERSIST_STATE'], true, 'BILIG_PERSIST_STATE'),
  }
}

export function registerSyncServerRuntimeRoutes(
  app: FastifyInstance,
  options: {
    documentService: DocumentControlService
    workbookAgentService?: WorkbookAgentService
    env: Record<string, string | undefined>
    runtimeConfig: {
      readonly browserAppBaseUrl?: string
    }
    webEnabled: boolean
  },
): void {
  const webRuntimeConfig = resolveWebRuntimeConfig(options.env)

  app.get('/healthz', async () => ({
    ok: true,
    service: 'bilig-app',
    zeroSync: false,
    web: options.webEnabled,
    workbookAgent: options.workbookAgentService?.getObservabilitySnapshot() ?? { enabled: false },
  }))

  app.get('/runtime-config.json', async (request, reply) => {
    const session = resolveSessionIdentity(request, reply)
    reply.header('cache-control', 'no-store')
    return {
      ...webRuntimeConfig,
      currentUserId: session.userID,
      workbookAgentEnabled: options.workbookAgentService?.enabled ?? false,
    } satisfies BiligRuntimeConfig
  })

  const handleSessionRequest = async (request: FastifyRequest, reply: FastifyReply) => {
    const session = resolveSessionIdentity(request, reply)
    const requestSession = resolveRequestSession(request)
    return createRuntimeSession({
      authToken: session.userID,
      userId: session.userID,
      roles: requestSession.roles,
      isAuthenticated: requestSession.isAuthenticated,
      authSource: requestSession.authSource,
    }) satisfies RuntimeSession
  }
  app.get('/v2/session', handleSessionRequest)

  app.post('/v2/agent/frames', async (request: FastifyRequest<{ Body: Buffer }>, reply: FastifyReply) => {
    const response = await runPromise(
      options.documentService.handleAgentFrame(decodeAgentFrame(request.body), {
        serverUrl: resolveRequestBaseUrl(request, '127.0.0.1:4321'),
        ...(options.runtimeConfig.browserAppBaseUrl ? { browserAppBaseUrl: options.runtimeConfig.browserAppBaseUrl } : {}),
      }),
    )
    reply.header('content-type', 'application/octet-stream')
    return Buffer.from(encodeAgentFrame(response))
  })
}
