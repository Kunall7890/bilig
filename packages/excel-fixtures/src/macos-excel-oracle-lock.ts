import { mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const defaultLockWaitMs = 120_000
const defaultStaleLockMs = 10 * 60_000
const lockPollMs = 250

export interface MacosExcelOracleLockOptions {
  readonly label?: string
  readonly lockRoot?: string
  readonly now?: () => number
  readonly processAlive?: (pid: number) => boolean
  readonly processId?: number
  readonly sleep?: (ms: number) => void
  readonly staleLockMs?: number
  readonly timeoutMs?: number | undefined
}

interface MacosExcelOracleLockOwner {
  readonly acquiredAt: string
  readonly acquiredAtMs: number
  readonly label?: string
  readonly pid: number
}

export function withMacosExcelOracleLock<T>(options: MacosExcelOracleLockOptions, run: () => T): T {
  const release = acquireMacosExcelOracleLock(options)
  try {
    return run()
  } finally {
    release()
  }
}

export function acquireMacosExcelOracleLock(options: MacosExcelOracleLockOptions = {}): () => void {
  const now = options.now ?? Date.now
  const sleep = options.sleep ?? sleepSync
  const lockRoot = options.lockRoot ?? defaultMacosExcelOracleLockRoot()
  const lockDir = macosExcelOracleLockDir(lockRoot)
  const ownerPath = join(lockDir, 'owner.json')
  const timeoutMs = Math.max(options.timeoutMs ?? defaultLockWaitMs, defaultLockWaitMs)
  const deadline = now() + timeoutMs
  let lastError: unknown

  mkdirSync(lockRoot, { recursive: true })

  while (now() <= deadline) {
    try {
      mkdirSync(lockDir)
      writeFileSync(ownerPath, JSON.stringify(lockOwner(options, now()), null, 2))
      return () => {
        rmSync(lockDir, { recursive: true, force: true })
      }
    } catch (error) {
      lastError = error
      if (!isNodeError(error) || error.code !== 'EEXIST') {
        throw error
      }
      removeStaleMacosExcelOracleLock(lockDir, options)
    }

    const remainingMs = deadline - now()
    if (remainingMs <= 0) {
      throw new Error(`Timed out waiting for macOS Desktop Excel oracle lock at ${lockDir}: ${String(lastError)}`)
    }
    sleep(Math.min(lockPollMs, remainingMs))
  }

  throw new Error(`Timed out waiting for macOS Desktop Excel oracle lock at ${lockDir}: ${String(lastError)}`)
}

export function macosExcelOracleLockDir(lockRoot: string = defaultMacosExcelOracleLockRoot()): string {
  return join(lockRoot, 'macos-excel-oracle.lock')
}

function removeStaleMacosExcelOracleLock(lockDir: string, options: MacosExcelOracleLockOptions): void {
  const owner = readLockOwner(lockDir)
  const processAlive = options.processAlive ?? defaultProcessAlive
  if (owner) {
    if (!processAlive(owner.pid)) {
      rmSync(lockDir, { recursive: true, force: true })
    }
    return
  }

  const now = options.now ?? Date.now
  const staleLockMs = options.staleLockMs ?? defaultStaleLockMs
  const lockAgeMs = now() - lockMtimeMs(lockDir)
  if (lockAgeMs >= staleLockMs) {
    rmSync(lockDir, { recursive: true, force: true })
  }
}

function readLockOwner(lockDir: string): MacosExcelOracleLockOwner | undefined {
  try {
    const parsed: unknown = JSON.parse(readFileSync(join(lockDir, 'owner.json'), 'utf8'))
    if (!isRecord(parsed) || typeof parsed['pid'] !== 'number' || typeof parsed['acquiredAtMs'] !== 'number') {
      return undefined
    }
    return {
      acquiredAt: typeof parsed['acquiredAt'] === 'string' ? parsed['acquiredAt'] : new Date(parsed['acquiredAtMs']).toISOString(),
      acquiredAtMs: parsed['acquiredAtMs'],
      ...(typeof parsed['label'] === 'string' ? { label: parsed['label'] } : {}),
      pid: parsed['pid'],
    }
  } catch {
    return undefined
  }
}

function lockMtimeMs(lockDir: string): number {
  try {
    return statSync(lockDir).mtimeMs
  } catch {
    return 0
  }
}

function lockOwner(options: MacosExcelOracleLockOptions, acquiredAtMs: number): MacosExcelOracleLockOwner {
  return {
    acquiredAt: new Date(acquiredAtMs).toISOString(),
    acquiredAtMs,
    ...(options.label ? { label: options.label } : {}),
    pid: options.processId ?? process.pid,
  }
}

function defaultMacosExcelOracleLockRoot(): string {
  return join(tmpdir(), 'bilig-excel-oracle')
}

function defaultProcessAlive(pid: number): boolean {
  if (!Number.isInteger(pid) || pid <= 0) {
    return false
  }
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return isNodeError(error) && error.code === 'EPERM'
  }
}

function sleepSync(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isNodeError(error: unknown): error is NodeJS.ErrnoException {
  return isRecord(error) && typeof error['code'] === 'string'
}
