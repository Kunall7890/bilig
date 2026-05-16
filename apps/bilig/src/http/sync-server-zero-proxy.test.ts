import { describe, expect, it } from 'vitest'

import { resolveZeroProxyUpstream } from './sync-server-zero-proxy.js'

describe('sync server zero proxy config', () => {
  it('ignores missing and blank zero proxy upstream config', () => {
    expect(resolveZeroProxyUpstream({})).toBeUndefined()
    expect(resolveZeroProxyUpstream({ BILIG_ZERO_PROXY_UPSTREAM: '   ' })).toBeUndefined()
  })

  it('trims absolute http zero proxy upstream config', () => {
    expect(resolveZeroProxyUpstream({ BILIG_ZERO_PROXY_UPSTREAM: ' https://zero.example.com/cache ' })).toBe(
      'https://zero.example.com/cache',
    )
  })

  it.each(['/zero', 'zero.example.com', 'ftp://zero.example.com/cache'])('rejects malformed zero proxy upstream config %s', (upstream) => {
    expect(() => resolveZeroProxyUpstream({ BILIG_ZERO_PROXY_UPSTREAM: upstream })).toThrow(
      `BILIG_ZERO_PROXY_UPSTREAM must be an absolute http(s) URL, got ${upstream}`,
    )
  })
})
