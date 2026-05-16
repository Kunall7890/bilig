export interface BiligAppRuntimeConfig {
  readonly host: string
  readonly appPort: number
  readonly publicServerUrl: string
  readonly browserAppBaseUrl: string
  readonly maxImportBytes?: number
}

export function resolveBiligAppRuntimeConfig(env: Readonly<Record<string, string | undefined>> = process.env): BiligAppRuntimeConfig {
  const host = env['HOST'] ?? '0.0.0.0'
  const appPort = parseTcpPort(env['PORT'] ?? '4321', 'PORT')
  const publicServerUrl = env['BILIG_PUBLIC_SERVER_URL'] ?? `http://127.0.0.1:${appPort}`
  const browserAppBaseUrl = env['BILIG_WEB_APP_BASE_URL'] ?? publicServerUrl
  const maxImportBytes = parseOptionalPositiveInteger(env['BILIG_AGENT_IMPORT_MAX_BYTES'], 'BILIG_AGENT_IMPORT_MAX_BYTES')

  return {
    host,
    appPort,
    publicServerUrl,
    browserAppBaseUrl,
    ...(maxImportBytes !== undefined ? { maxImportBytes } : {}),
  }
}

function parseTcpPort(value: string, name: string): number {
  if (!/^(?:[1-9]\d*)$/u.test(value)) {
    throw new Error(`${name} must be a TCP port between 1 and 65535, got ${value}`)
  }

  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed) || parsed > 65_535) {
    throw new Error(`${name} must be a TCP port between 1 and 65535, got ${value}`)
  }

  return parsed
}

function parseOptionalPositiveInteger(value: string | undefined, name: string): number | undefined {
  if (value === undefined || value.length === 0) {
    return undefined
  }

  if (!/^(?:[1-9]\d*)$/u.test(value)) {
    throw new Error(`${name} must be a positive integer, got ${value}`)
  }

  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed)) {
    throw new Error(`${name} must be a safe integer, got ${value}`)
  }

  return parsed
}
