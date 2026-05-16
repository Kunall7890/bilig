export function resolvePreferredPort(configuredPort: string | undefined, fallbackPort: number): number {
  return parseTcpPort(configuredPort ?? String(fallbackPort))
}

export function resolvePreferredZeroPort(
  configuredZeroPort: string | undefined,
  configuredZeroProxyUpstream: string | undefined,
  fallbackPort: number,
): number {
  const upstreamPort = configuredZeroProxyUpstream ? new URL(configuredZeroProxyUpstream).port || undefined : undefined
  return parseTcpPort(configuredZeroPort ?? upstreamPort ?? String(fallbackPort))
}

export async function resolveRequestedOrAvailablePort(options: {
  readonly preferredPort: number
  readonly explicitPort: string | undefined
  readonly label: string
  readonly canUseRequestedPort: (port: number) => Promise<boolean>
  readonly remainingOffsets?: number
}): Promise<number> {
  const { preferredPort, explicitPort, label, canUseRequestedPort, remainingOffsets = 10 } = options
  if (explicitPort) {
    if (!(await canUseRequestedPort(preferredPort))) {
      throw new Error(`${label} ${preferredPort} is already in use.`)
    }
    return preferredPort
  }

  return findAvailablePort(preferredPort, remainingOffsets, label, canUseRequestedPort)
}

export async function canUsePort(options: {
  readonly port: number
  readonly bindProbe: (port: number) => Promise<boolean>
}): Promise<boolean> {
  return options.bindProbe(options.port)
}

async function findAvailablePort(
  startPort: number,
  remainingOffsets: number,
  label: string,
  canUseRequestedPort: (port: number) => Promise<boolean>,
  offset = 0,
): Promise<number> {
  if (offset >= remainingOffsets) {
    throw new Error(`Unable to find an available ${label}.`)
  }
  const candidate = startPort + offset
  if (await canUseRequestedPort(candidate)) {
    return candidate
  }
  return findAvailablePort(startPort, remainingOffsets, label, canUseRequestedPort, offset + 1)
}

function parseTcpPort(value: string): number {
  if (!/^(?:[1-9]\d*)$/u.test(value)) {
    throw new Error(`Expected a decimal TCP port, got ${value}`)
  }
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed) || parsed > 65_535) {
    throw new Error(`Expected a TCP port between 1 and 65535, got ${value}`)
  }
  return parsed
}
