import type { WorkbookAgentThreadScope } from '@bilig/contracts'
import { logDebug } from './runtime-logger.js'

const STORAGE_KEY_PREFIX = 'bilig:workbook-agent:'
const DRAFT_STORAGE_KEY_PREFIX = 'bilig:workbook-agent-drafts:'

export interface StoredWorkbookAgentThreadRef {
  threadId: string
}

function storageKey(documentId: string): string {
  return `${STORAGE_KEY_PREFIX}${documentId}`
}

function draftStorageKey(documentId: string): string {
  return `${DRAFT_STORAGE_KEY_PREFIX}${documentId}`
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function getSessionStorage(documentId: string): Storage | null {
  try {
    return window.sessionStorage
  } catch (error) {
    logDebug('Failed to access workbook agent storage', { documentId, error })
    return null
  }
}

function readSessionStorageItem(documentId: string, key: string): string | null {
  const storage = getSessionStorage(documentId)
  if (!storage) {
    return null
  }
  try {
    return storage.getItem(key)
  } catch (error) {
    logDebug('Failed to read workbook agent storage', { documentId, key, error })
    return null
  }
}

function writeSessionStorageItem(documentId: string, key: string, value: string): void {
  const storage = getSessionStorage(documentId)
  if (!storage) {
    return
  }
  try {
    storage.setItem(key, value)
  } catch (error) {
    logDebug('Failed to persist workbook agent storage', { documentId, key, error })
  }
}

function removeSessionStorageItem(documentId: string, key: string): void {
  const storage = getSessionStorage(documentId)
  if (!storage) {
    return
  }
  try {
    storage.removeItem(key)
  } catch (error) {
    logDebug('Failed to clear workbook agent storage', { documentId, key, error })
  }
}

function parseStoredWorkbookAgentSession(value: unknown): StoredWorkbookAgentThreadRef | null {
  if (!isRecord(value) || typeof value['threadId'] !== 'string') {
    return null
  }
  const threadId = value['threadId'].trim()
  return threadId.length > 0 ? { threadId } : null
}

export function loadStoredSession(documentId: string): StoredWorkbookAgentThreadRef | null {
  const key = storageKey(documentId)
  try {
    const raw = readSessionStorageItem(documentId, key)
    if (!raw) {
      return null
    }
    const parsed = JSON.parse(raw) as unknown
    const storedSession = parseStoredWorkbookAgentSession(parsed)
    if (storedSession) {
      return storedSession
    }
    removeSessionStorageItem(documentId, key)
  } catch (error) {
    logDebug('Failed to load stored workbook agent session', { documentId, error })
    removeSessionStorageItem(documentId, key)
  }
  return null
}

export function persistStoredSession(documentId: string, value: StoredWorkbookAgentThreadRef): void {
  const storedSession = parseStoredWorkbookAgentSession(value)
  if (!storedSession) {
    removeSessionStorageItem(documentId, storageKey(documentId))
    return
  }
  writeSessionStorageItem(documentId, storageKey(documentId), JSON.stringify(storedSession))
}

export function clearStoredSession(documentId: string): void {
  removeSessionStorageItem(documentId, storageKey(documentId))
}

export function loadStoredDrafts(documentId: string): Record<string, string> {
  const key = draftStorageKey(documentId)
  try {
    const raw = readSessionStorageItem(documentId, key)
    if (!raw) {
      return {}
    }
    const parsed = JSON.parse(raw) as unknown
    if (!isRecord(parsed)) {
      removeSessionStorageItem(documentId, key)
      return {}
    }
    const drafts = Object.fromEntries(
      Object.entries(parsed).flatMap(([storedDraftKey, value]) => (typeof value === 'string' ? ([[storedDraftKey, value]] as const) : [])),
    )
    if (Object.keys(drafts).length !== Object.keys(parsed).length) {
      persistStoredDrafts(documentId, drafts)
    }
    return drafts
  } catch (error) {
    logDebug('Failed to load stored workbook agent draft', { documentId, error })
    removeSessionStorageItem(documentId, key)
    return {}
  }
}

export function persistStoredDrafts(documentId: string, drafts: Record<string, string>): void {
  const entries = Object.entries(drafts).filter((entry) => entry[1].length > 0)
  if (entries.length === 0) {
    removeSessionStorageItem(documentId, draftStorageKey(documentId))
    return
  }
  writeSessionStorageItem(documentId, draftStorageKey(documentId), JSON.stringify(Object.fromEntries(entries)))
}

export function clearStoredDraft(documentId: string, key: string): void {
  const drafts = loadStoredDrafts(documentId)
  if (!(key in drafts)) {
    return
  }
  delete drafts[key]
  persistStoredDrafts(documentId, drafts)
}

export function draftKey(threadId: string | null, scope: WorkbookAgentThreadScope): string {
  return threadId ? `thread:${threadId}` : `new:${scope}`
}
