import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'

export async function withPublicWorkbookCorpusCacheLock<T>(cacheDir: string, owner: string, task: () => Promise<T>): Promise<T> {
  const lockDir = join(cacheDir, '.public-workbook-corpus.lock')
  acquireLock(lockDir, owner)
  try {
    return await task()
  } finally {
    rmSync(lockDir, { recursive: true, force: true })
  }
}

function acquireLock(lockDir: string, owner: string): void {
  mkdirSync(dirname(lockDir), { recursive: true })
  for (;;) {
    try {
      mkdirSync(lockDir)
      writeFileSync(
        join(lockDir, 'owner.json'),
        `${JSON.stringify(
          {
            pid: process.pid,
            owner,
            startedAt: new Date().toISOString(),
          },
          null,
          2,
        )}\n`,
      )
      return
    } catch (error) {
      if (!isFileExistsError(error)) {
        throw error
      }
      const ownerInfo = readLockOwner(lockDir)
      if (ownerInfo.pid !== null && isProcessRunning(ownerInfo.pid)) {
        throw new Error(
          `Public workbook corpus cache is already locked by pid ${String(ownerInfo.pid)} for ${ownerInfo.owner ?? 'unknown owner'}`,
          { cause: error },
        )
      }
      rmSync(lockDir, { recursive: true, force: true })
    }
  }
}

function readLockOwner(lockDir: string): { readonly pid: number | null; readonly owner: string | null } {
  const ownerPath = join(lockDir, 'owner.json')
  if (!existsSync(ownerPath)) {
    return { pid: null, owner: null }
  }
  try {
    const parsed = JSON.parse(readFileSync(ownerPath, 'utf8')) as unknown
    if (parsed && typeof parsed === 'object') {
      const pid = Reflect.get(parsed, 'pid')
      const owner = Reflect.get(parsed, 'owner')
      return {
        pid: Number.isInteger(pid) ? pid : null,
        owner: typeof owner === 'string' ? owner : null,
      }
    }
  } catch {
    return { pid: null, owner: null }
  }
  return { pid: null, owner: null }
}

function isProcessRunning(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return typeof error === 'object' && error !== null && Reflect.get(error, 'code') === 'EPERM'
  }
}

function isFileExistsError(error: unknown): boolean {
  return typeof error === 'object' && error !== null && Reflect.get(error, 'code') === 'EEXIST'
}
