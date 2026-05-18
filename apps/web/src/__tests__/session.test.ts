import { afterEach, describe, expect, test, vi } from 'vitest'
import { createLocalOnlyRuntimeSession, loadRuntimeSession, type BiligRuntimeSession } from '../session.js'

const jsonResponse = (body: unknown) =>
  new Response(JSON.stringify(body), {
    headers: {
      'content-type': 'application/json',
    },
  })

const failingFetchImpl = async () => new Response('forbidden', { status: 403 })

describe('loadRuntimeSession', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  test('creates a local-only guest session for standalone web development', () => {
    expect(createLocalOnlyRuntimeSession('guest:local-dev')).toEqual<BiligRuntimeSession>({
      authToken: 'guest:local-dev',
      userId: 'guest:local-dev',
      roles: ['editor'],
      isAuthenticated: false,
      authSource: 'guest',
    })
  })

  test('uses the server-provided auth token when present', async () => {
    const response = {
      authToken: 'token-123',
      userId: 'user-123',
      roles: ['editor'],
      isAuthenticated: true,
      authSource: 'header',
    }
    const fetchImpl = async () => jsonResponse(response)

    const session = await loadRuntimeSession(fetchImpl)

    expect(session).toEqual<BiligRuntimeSession>({
      authToken: 'token-123',
      userId: 'user-123',
      roles: ['editor'],
      isAuthenticated: true,
      authSource: 'header',
    })
  })

  test('keeps the browser fetch receiver intact when using the default implementation', async () => {
    const response = {
      authToken: 'token-123',
      userId: 'user-123',
      roles: ['editor'],
      isAuthenticated: true,
      authSource: 'header',
    }
    const fetchImpl = vi.fn(async function (this: typeof globalThis, input: RequestInfo | URL, init?: RequestInit) {
      expect(this).toBe(globalThis)
      expect(input).toBe('/v2/session')
      expect(init).toEqual({
        credentials: 'include',
        headers: {
          accept: 'application/json',
        },
      })
      return jsonResponse(response)
    })
    vi.stubGlobal('fetch', fetchImpl)

    await expect(loadRuntimeSession()).resolves.toEqual<BiligRuntimeSession>(response)
  })

  test('falls back to the resolved user id when auth token is missing', async () => {
    const fetchImpl = async () =>
      jsonResponse({
        userId: 'guest:abc',
        authToken: 'guest:abc',
        roles: ['editor'],
        isAuthenticated: false,
        authSource: 'guest',
      })

    const session = await loadRuntimeSession(fetchImpl)

    expect(session.authToken).toBe('guest:abc')
    expect(session.userId).toBe('guest:abc')
    expect(session.authSource).toBe('guest')
  })

  test('fails when the runtime session endpoint returns an error response', async () => {
    await expect(loadRuntimeSession(failingFetchImpl)).rejects.toThrow(/Runtime session request failed/)
  })
})
