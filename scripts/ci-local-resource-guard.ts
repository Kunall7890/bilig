import { existsSync, readdirSync, readFileSync, type Dirent } from 'node:fs'
import { join } from 'node:path'

export const localCiResourceGuardOverrideEnv = 'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD'

export interface LocalCiResourceGuardStatus {
  readonly activeMarkerPaths: readonly string[]
}

export interface LocalCiResourceGuardOptions {
  readonly runLabel?: string
}

const coordinationDirectoryName = '.agent-coordination'
const guardMarkerNamePatterns = [/stop-interactive-corpus-runs\.md$/u, /memory-guard\.md$/u, /memory-pressure-stop\.md$/u]

export function readLocalCiResourceGuardStatus(rootDir: string): LocalCiResourceGuardStatus {
  const coordinationDirectory = join(rootDir, coordinationDirectoryName)
  if (!existsSync(coordinationDirectory)) {
    return { activeMarkerPaths: [] }
  }

  const activeMarkerPaths = readCoordinationEntries(coordinationDirectory)
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .filter((fileName) => guardMarkerNamePatterns.some((pattern) => pattern.test(fileName)))
    .filter((fileName) => isActiveGuardMarker(join(coordinationDirectory, fileName)))
    .map((fileName) => `${coordinationDirectoryName}/${fileName}`)
    .toSorted()

  return { activeMarkerPaths }
}

export function assertLocalCiResourceGuardAllowsRun(
  rootDir: string,
  env: Readonly<Record<string, string | undefined>> = process.env,
  options: LocalCiResourceGuardOptions = {},
): void {
  if (env[localCiResourceGuardOverrideEnv] === '1') {
    return
  }

  const status = readLocalCiResourceGuardStatus(rootDir)
  if (status.activeMarkerPaths.length === 0) {
    return
  }

  throw new Error(
    [
      `Refusing to start ${options.runLabel ?? 'broad CI'} while the local resource guard is active.`,
      'Active coordination markers:',
      ...status.activeMarkerPaths.map((markerPath) => `- ${markerPath}`),
      `Set ${localCiResourceGuardOverrideEnv}=1 only for an intentional broad run after confirming host capacity.`,
    ].join('\n'),
  )
}

function isActiveGuardMarker(filePath: string): boolean {
  try {
    const content = readFileSync(filePath, 'utf8')
    return /^Status:\s*active\b/imu.test(content) || /^#\s+Memory guard active\b/imu.test(content)
  } catch {
    return false
  }
}

function readCoordinationEntries(coordinationDirectory: string): Dirent[] {
  try {
    return readdirSync(coordinationDirectory, { withFileTypes: true })
  } catch {
    return []
  }
}
