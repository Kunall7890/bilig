import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  clearStoredSession,
  loadStoredDrafts,
  loadStoredSession,
  persistStoredDrafts,
  persistStoredSession,
} from '../workbook-agent-pane-storage.js'

describe('workbook agent pane storage', () => {
  const storage = new Map<string, string>()

  beforeEach(() => {
    storage.clear()
    vi.stubGlobal('window', {
      sessionStorage: {
        getItem(key: string) {
          return storage.get(key) ?? null
        },
        removeItem(key: string) {
          storage.delete(key)
        },
        setItem(key: string, value: string) {
          storage.set(key, value)
        },
      },
    })
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('removes corrupt stored session JSON after falling back', () => {
    storage.set('bilig:workbook-agent:doc-1', '{')

    expect(loadStoredSession('doc-1')).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(false)
  })

  it('rejects and removes blank stored thread ids', () => {
    storage.set('bilig:workbook-agent:doc-1', JSON.stringify({ threadId: '   ' }))

    expect(loadStoredSession('doc-1')).toBeNull()
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(false)

    persistStoredSession('doc-1', { threadId: '   ' })
    expect(storage.has('bilig:workbook-agent:doc-1')).toBe(false)
  })

  it('normalizes and persists valid stored thread ids', () => {
    persistStoredSession('doc-1', { threadId: '  thr-1  ' })

    expect(loadStoredSession('doc-1')).toEqual({ threadId: 'thr-1' })
    expect(storage.get('bilig:workbook-agent:doc-1')).toBe(JSON.stringify({ threadId: 'thr-1' }))
  })

  it('removes corrupt stored draft JSON after falling back', () => {
    storage.set('bilig:workbook-agent-drafts:doc-1', '{')

    expect(loadStoredDrafts('doc-1')).toEqual({})
    expect(storage.has('bilig:workbook-agent-drafts:doc-1')).toBe(false)
  })

  it('self-heals stored draft maps with non-string values', () => {
    storage.set('bilig:workbook-agent-drafts:doc-1', JSON.stringify({ keep: 'draft', drop: 42 }))

    expect(loadStoredDrafts('doc-1')).toEqual({ keep: 'draft' })
    expect(storage.get('bilig:workbook-agent-drafts:doc-1')).toBe(JSON.stringify({ keep: 'draft' }))
  })

  it('does not throw when session storage writes fail', () => {
    vi.stubGlobal('window', {
      sessionStorage: {
        getItem() {
          return null
        },
        removeItem() {
          throw new Error('storage denied')
        },
        setItem() {
          throw new Error('storage denied')
        },
      },
    })

    expect(() => persistStoredSession('doc-1', { threadId: 'thr-1' })).not.toThrow()
    expect(() => persistStoredDrafts('doc-1', { key: 'draft' })).not.toThrow()
    expect(() => clearStoredSession('doc-1')).not.toThrow()
  })
})
