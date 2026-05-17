const WORKBOOK_REPLICA_ID_STORAGE_KEY_PREFIX = 'bilig:workbook-replica-id:'
const WORKBOOK_REPLICA_ID_PREFIX = 'browser:'

function createWorkbookReplicaId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${WORKBOOK_REPLICA_ID_PREFIX}${crypto.randomUUID()}`
  }
  return `${WORKBOOK_REPLICA_ID_PREFIX}${Math.random().toString(36).slice(2)}`
}

function parseWorkbookReplicaId(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const trimmedValue = value.trim()
  return trimmedValue.startsWith(WORKBOOK_REPLICA_ID_PREFIX) &&
    trimmedValue.length > WORKBOOK_REPLICA_ID_PREFIX.length &&
    !/\s/u.test(trimmedValue)
    ? trimmedValue
    : null
}

function storageKey(documentId: string, userId: string): string {
  return `${WORKBOOK_REPLICA_ID_STORAGE_KEY_PREFIX}${encodeURIComponent(documentId)}:${encodeURIComponent(userId)}`
}

function removeStoredWorkbookReplicaId(key: string): void {
  try {
    sessionStorage.removeItem(key)
  } catch {
    // Ignore storage cleanup failures and keep workbook startup usable.
  }
}

export function loadOrCreateWorkbookReplicaId(documentId: string, userId: string): string {
  if (typeof sessionStorage === 'undefined') {
    return createWorkbookReplicaId()
  }
  const key = storageKey(documentId, userId)
  try {
    const storedValue = parseWorkbookReplicaId(sessionStorage.getItem(key))
    if (storedValue) {
      return storedValue
    }
    removeStoredWorkbookReplicaId(key)
    const nextValue = createWorkbookReplicaId()
    sessionStorage.setItem(key, nextValue)
    return nextValue
  } catch {
    return createWorkbookReplicaId()
  }
}
