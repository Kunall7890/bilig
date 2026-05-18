import { describe, expect, it } from 'vitest'
import { expectThreadStartResponse, parseServerNotification } from './codex-app-server-message-parsers.js'

describe('Codex app-server message parsers', () => {
  it('preserves thread lifecycle status on thread responses and notifications', () => {
    const response = expectThreadStartResponse({
      thread: {
        id: 'thr-1',
        preview: '',
        status: {
          type: 'active',
          activeFlags: ['waitingOnUserInput'],
        },
        turns: [],
      },
    })

    expect(response.thread.status).toEqual({
      type: 'active',
      activeFlags: ['waitingOnUserInput'],
    })

    const notification = parseServerNotification({
      method: 'thread/started',
      params: {
        thread: {
          id: 'thr-1',
          preview: '',
          status: {
            type: 'idle',
          },
          turns: [],
        },
      },
    })

    expect(notification).toEqual({
      method: 'thread/started',
      params: {
        thread: {
          id: 'thr-1',
          preview: '',
          status: {
            type: 'idle',
          },
          turns: [],
        },
      },
    })
  })

  it('rejects invalid thread lifecycle status payloads', () => {
    expect(() =>
      expectThreadStartResponse({
        thread: {
          id: 'thr-1',
          preview: '',
          status: {
            type: 'active',
            activeFlags: ['unknown-flag'],
          },
          turns: [],
        },
      }),
    ).toThrow('Invalid Codex thread response')
  })
})
