import Fastify from 'fastify'
import { describe, expect, it } from 'vitest'
import { registerWorkPaperOpenApiRoutes } from './workpaper-openapi-routes.js'

function readObject(value: unknown): Record<string, unknown> {
  if (!isJsonObject(value)) {
    throw new Error(`Expected object, received ${JSON.stringify(value)}`)
  }
  return value
}

function isJsonObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

describe('WorkPaper OpenAPI tool routes', () => {
  it('serves a CORS-safe OpenAPI spec for Open WebUI tool server discovery', async () => {
    const app = Fastify({ logger: false })
    registerWorkPaperOpenApiRoutes(app)

    try {
      const response = await app.inject({
        method: 'GET',
        url: '/openapi/workpaper/openapi.json',
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['content-type']).toContain('application/json')
      expect(response.headers['access-control-allow-origin']).toBe('*')
      expect(response.headers['cache-control']).toBe('public, max-age=300')

      const spec = response.json()
      expect(spec).toMatchObject({
        openapi: '3.1.0',
        info: {
          title: 'Bilig WorkPaper OpenAPI Tool Server',
        },
        servers: [
          {
            url: 'https://bilig.proompteng.ai',
          },
        ],
      })
      expect(Object.keys(readObject(spec.paths))).toEqual([
        '/openapi/workpaper/list-sheets',
        '/openapi/workpaper/read-range',
        '/openapi/workpaper/set-cell-and-readback',
      ])
      expect(readObject(readObject(readObject(spec.paths)['/openapi/workpaper/set-cell-and-readback'])['post'])['operationId']).toBe(
        'set_workpaper_cell_and_readback',
      )
    } finally {
      await app.close()
    }
  })

  it('handles Open WebUI browser preflight requests', async () => {
    const app = Fastify({ logger: false })
    registerWorkPaperOpenApiRoutes(app)

    try {
      const response = await app.inject({
        method: 'OPTIONS',
        url: '/openapi/workpaper/set-cell-and-readback',
        headers: {
          origin: 'https://openwebui.com',
          'access-control-request-method': 'POST',
          'access-control-request-headers': 'content-type',
        },
      })

      expect(response.statusCode).toBe(204)
      expect(response.headers['access-control-allow-origin']).toBe('*')
      expect(response.headers['access-control-allow-methods']).toBe('POST, GET, OPTIONS')
      expect(response.headers['access-control-allow-headers']).toBe('accept, content-type')
    } finally {
      await app.close()
    }
  })

  it('edits the hosted demo WorkPaper and returns dependent readback proof in one OpenAPI call', async () => {
    const app = Fastify({ logger: false })
    registerWorkPaperOpenApiRoutes(app)

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/openapi/workpaper/set-cell-and-readback',
        headers: {
          'content-type': 'application/json',
        },
        payload: JSON.stringify({
          sheetName: 'Inputs',
          address: 'B3',
          value: 0.4,
          readbackRange: 'Summary!A1:B3',
        }),
      })

      expect(response.statusCode).toBe(200)
      expect(response.headers['access-control-allow-origin']).toBe('*')
      expect(response.headers['cache-control']).toBe('no-store')
      expect(response.json()).toMatchObject({
        editedCell: 'Inputs!B3',
        readbackRange: 'Summary!A1:B3',
        before: {
          serialized: 0.25,
        },
        after: {
          serialized: 0.4,
        },
        persistence: {
          persisted: false,
        },
        checks: {
          persisted: false,
          readbackChanged: true,
          restoredReadbackMatchesAfter: true,
          previousSerialized: 0.25,
          newSerialized: 0.4,
        },
      })
    } finally {
      await app.close()
    }
  })

  it('returns 400 JSON errors for invalid WorkPaper OpenAPI requests', async () => {
    const app = Fastify({ logger: false })
    registerWorkPaperOpenApiRoutes(app)

    try {
      const response = await app.inject({
        method: 'POST',
        url: '/openapi/workpaper/read-range',
        headers: {
          'content-type': 'application/json',
        },
        payload: JSON.stringify({
          range: 'Missing!A1:B2',
        }),
      })

      expect(response.statusCode).toBe(400)
      expect(response.headers['access-control-allow-origin']).toBe('*')
      expect(response.json()).toMatchObject({
        verified: false,
        error: expect.stringContaining('Invalid range: Missing!A1:B2'),
      })
    } finally {
      await app.close()
    }
  })
})
