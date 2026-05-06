import type { FastifyReply } from 'fastify'

const STATIC_CONNECT_SOURCES = ["'self'"] as const

function originFromEnvUrl(value: string | undefined): string | null {
  const trimmed = value?.trim()
  if (!trimmed || trimmed.startsWith('/')) {
    return null
  }
  try {
    const url = new URL(trimmed)
    if (url.protocol !== 'http:' && url.protocol !== 'https:' && url.protocol !== 'ws:' && url.protocol !== 'wss:') {
      return null
    }
    return url.origin
  } catch {
    return null
  }
}

function uniqueSources(values: readonly (string | null)[]): string[] {
  return [...new Set(values.filter((value): value is string => value !== null))]
}

export function buildSyncServerContentSecurityPolicy(env: Record<string, string | undefined> = process.env): string {
  const connectSources = uniqueSources([
    ...STATIC_CONNECT_SOURCES,
    originFromEnvUrl(env['BILIG_PUBLIC_SERVER_URL']),
    originFromEnvUrl(env['BILIG_WEB_APP_BASE_URL']),
    originFromEnvUrl(env['BILIG_ZERO_CACHE_URL']),
  ])

  return [
    "default-src 'self'",
    "object-src 'none'",
    "base-uri 'none'",
    "frame-ancestors 'self'",
    "script-src 'self' 'wasm-unsafe-eval'",
    "worker-src 'self' blob:",
    "style-src 'self' 'unsafe-inline'",
    "img-src 'self' data: blob: https:",
    "font-src 'self' data:",
    'connect-src ' + connectSources.join(' '),
    "media-src 'self' data: blob:",
    "manifest-src 'self'",
    "form-action 'self'",
  ].join('; ')
}

export function applySyncServerSecurityHeaders(reply: FastifyReply, env: Record<string, string | undefined> = process.env): void {
  reply.header('Cross-Origin-Opener-Policy', 'same-origin')
  reply.header('Cross-Origin-Embedder-Policy', 'require-corp')
  reply.header('Origin-Agent-Cluster', '?1')
  reply.header('Content-Security-Policy', buildSyncServerContentSecurityPolicy(env))
}
