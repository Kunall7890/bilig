import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

const forbiddenMcpRegistryClaimPatterns = [
  /compatible with Excel/iu,
  /Excel-compatible/iu,
  /Excel parity verified/iu,
  /Excel readiness score/iu,
  /Excel certification/iu,
  /\bcertified\b/iu,
  /\bguaranteed\b/iu,
  /100%\s+compatible/iu,
  /compatibility\s+score/iu,
  /compatibilityScore/u,
  /excelCompatibilityPercent/u,
  /verifies workbook compatibility/iu,
  /guarantees workbook execution/iu,
  /detects all workbook issues/iu,
] as const

export function syncStagedMcpServerMetadata(packageName: string, stagedPackageDir: string, targetVersion: string): void {
  const manifest = readPackageManifest(stagedPackageDir)
  if (!shouldValidateMcpMetadata(packageName, manifest)) {
    return
  }

  const serverJsonPath = join(stagedPackageDir, 'server.json')
  const serverJson = readMcpServerJson(serverJsonPath)
  serverJson['version'] = targetVersion

  const npmPackage = findNpmPackageEntry(serverJson, manifest['name'])
  if (!npmPackage) {
    throw new Error(`Staged ${packageName} server.json must include an npm package entry for ${String(manifest['name'])}`)
  }
  npmPackage['version'] = targetVersion

  writeFileSync(serverJsonPath, `${JSON.stringify(serverJson, null, 2)}\n`)
}

export function validateStagedMcpServerMetadata(packageName: string, stagedPackageDir: string, expectedVersion: string): void {
  const manifest = readPackageManifest(stagedPackageDir)
  if (!shouldValidateMcpMetadata(packageName, manifest)) {
    return
  }

  const serverJsonPath = join(stagedPackageDir, 'server.json')
  const serverJson = readMcpServerJson(serverJsonPath)

  if (serverJson['name'] !== manifest['mcpName']) {
    throw new Error(
      `Staged ${packageName} server.json name must match package.json mcpName: ${String(serverJson['name'])} !== ${String(
        manifest['mcpName'],
      )}`,
    )
  }
  if (serverJson['version'] !== expectedVersion) {
    throw new Error(
      `Staged ${packageName} server.json version must match package version: ${String(serverJson['version'])} !== ${expectedVersion}`,
    )
  }
  if (typeof serverJson['description'] !== 'string' || serverJson['description'].length > 100) {
    throw new Error(`Staged ${packageName} server.json description must be a string no longer than 100 characters`)
  }
  assertNoMcpRegistryOverclaimText(packageName, serverJson)

  const npmPackage = findNpmPackageEntry(serverJson, manifest['name'])
  if (!npmPackage) {
    throw new Error(`Staged ${packageName} server.json must include an npm package entry for ${String(manifest['name'])}`)
  }
  if (npmPackage['version'] !== expectedVersion) {
    throw new Error(
      `Staged ${packageName} server.json npm package version must match package version: ${String(npmPackage['version'])} !== ${expectedVersion}`,
    )
  }

  if (!findHostedRemoteEntry(serverJson)) {
    throw new Error(`Staged ${packageName} server.json must include the hosted Streamable HTTP remote endpoint`)
  }
}

function shouldValidateMcpMetadata(_packageName: string, manifest: Record<string, unknown>): manifest is { name: string; mcpName: string } {
  return typeof manifest['name'] === 'string' && typeof manifest['mcpName'] === 'string'
}

function readPackageManifest(stagedPackageDir: string): Record<string, unknown> {
  const manifestPath = join(stagedPackageDir, 'package.json')
  return readJsonRecord(manifestPath)
}

function readMcpServerJson(serverJsonPath: string): Record<string, unknown> {
  if (!existsSync(serverJsonPath)) {
    throw new Error(`Staged MCP package is missing server.json: ${serverJsonPath}`)
  }
  return readJsonRecord(serverJsonPath)
}

function readJsonRecord(path: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isRecord(parsed)) {
    throw new Error(`Expected JSON object in ${path}`)
  }
  return parsed
}

function findNpmPackageEntry(serverJson: Record<string, unknown>, packageName: string): Record<string, unknown> | undefined {
  const packages = serverJson['packages']
  if (!Array.isArray(packages)) {
    return undefined
  }
  return packages.find(
    (entry): entry is Record<string, unknown> => isRecord(entry) && entry['registryType'] === 'npm' && entry['identifier'] === packageName,
  )
}

function findHostedRemoteEntry(serverJson: Record<string, unknown>): Record<string, unknown> | undefined {
  const remotes = serverJson['remotes']
  if (!Array.isArray(remotes)) {
    return undefined
  }
  return remotes.find(
    (entry): entry is Record<string, unknown> =>
      isRecord(entry) && entry['type'] === 'streamable-http' && entry['url'] === 'https://bilig.proompteng.ai/mcp',
  )
}

function assertNoMcpRegistryOverclaimText(packageName: string, serverJson: Record<string, unknown>): void {
  const metadataText = collectJsonStrings(serverJson).join('\n')
  const matchedPattern = forbiddenMcpRegistryClaimPatterns.find((pattern) => pattern.test(metadataText))
  if (matchedPattern) {
    throw new Error(`Staged ${packageName} server.json contains overclaiming MCP Registry wording: ${matchedPattern.source}`)
  }
}

function collectJsonStrings(value: unknown): string[] {
  if (typeof value === 'string') {
    return [value]
  }
  if (Array.isArray(value)) {
    return value.flatMap((entry) => collectJsonStrings(entry))
  }
  if (isRecord(value)) {
    return Object.values(value).flatMap((entry) => collectJsonStrings(entry))
  }
  return []
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
