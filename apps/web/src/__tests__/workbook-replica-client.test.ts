import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { loadOrCreateWorkbookReplicaId } from '../workbook-replica-client.js'

function createStorage() {
  const values = new Map<string, string>()
  const getItem = vi.fn((key: string) => values.get(key) ?? null)
  const setItem = vi.fn((key: string, value: string) => {
    values.set(key, value)
  })
  const removeItem = vi.fn((key: string) => {
    values.delete(key)
  })
  const storage = {
    getItem,
    setItem,
    removeItem,
    clear: vi.fn(() => {
      values.clear()
    }),
    key: vi.fn((index: number) => [...values.keys()][index] ?? null),
    get length() {
      return values.size
    },
  } satisfies Storage
  return { removeItem, setItem, storage, values }
}

describe('workbook replica client id', () => {
  let storage: Storage
  let values: Map<string, string>
  let removeItem: ReturnType<typeof vi.fn>
  let setItem: ReturnType<typeof vi.fn>

  beforeEach(() => {
    const created = createStorage()
    storage = created.storage
    values = created.values
    removeItem = created.removeItem
    setItem = created.setItem
    vi.stubGlobal('sessionStorage', storage)
    vi.stubGlobal('crypto', {
      randomUUID: vi.fn(() => 'stable-replica'),
    })
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('keeps a stable browser replica id for the same document and user within a tab', () => {
    const first = loadOrCreateWorkbookReplicaId('doc-1', 'user-1')
    const second = loadOrCreateWorkbookReplicaId('doc-1', 'user-1')

    expect(first).toBe('browser:stable-replica')
    expect(second).toBe(first)
    expect(setItem).toHaveBeenCalledTimes(1)
  })

  it('scopes replica ids by document and user', () => {
    values.set('bilig:workbook-replica-id:doc-1:user-1', 'browser:stored-user-1')
    values.set('bilig:workbook-replica-id:doc-1:user-2', 'browser:stored-user-2')

    expect(loadOrCreateWorkbookReplicaId('doc-1', 'user-1')).toBe('browser:stored-user-1')
    expect(loadOrCreateWorkbookReplicaId('doc-1', 'user-2')).toBe('browser:stored-user-2')
  })

  it('replaces corrupt stored replica ids', () => {
    values.set('bilig:workbook-replica-id:doc-1:user-1', 'bad replica')

    expect(loadOrCreateWorkbookReplicaId('doc-1', 'user-1')).toBe('browser:stable-replica')
    expect(removeItem).toHaveBeenCalledWith('bilig:workbook-replica-id:doc-1:user-1')
  })
})
