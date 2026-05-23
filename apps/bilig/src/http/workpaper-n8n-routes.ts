import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import { createN8nForecastProof, type N8nForecastRequestBody } from '@bilig/headless'

export function registerWorkPaperN8nRoutes(app: FastifyInstance): void {
  app.post('/api/workpaper/n8n/forecast', handleN8nForecastRequest)
}

function handleN8nForecastRequest(request: FastifyRequest<{ Body: N8nForecastRequestBody }>, reply: FastifyReply) {
  try {
    reply.header('cache-control', 'no-store')
    reply.header('content-type', 'application/json; charset=utf-8')
    return createN8nForecastProof(request.body ?? {})
  } catch (error) {
    reply.code(400)
    return {
      verified: false,
      error: error instanceof Error ? error.message : 'Invalid n8n WorkPaper forecast request',
    }
  }
}
