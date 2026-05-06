import { describe, expect, it } from 'vitest'

import { buildSyncServerContentSecurityPolicy } from './sync-server-security-headers.js'

describe('sync server security headers', () => {
  it('builds a restrictive workbook browser CSP with required runtime allowances', () => {
    const policy = buildSyncServerContentSecurityPolicy({
      BILIG_PUBLIC_SERVER_URL: 'https://bilig.example.com',
      BILIG_WEB_APP_BASE_URL: 'https://workbooks.example.com',
      BILIG_ZERO_CACHE_URL: 'https://zero.example.com/zero',
    })

    expect(policy).toContain("default-src 'self'")
    expect(policy).toContain("object-src 'none'")
    expect(policy).toContain("base-uri 'none'")
    expect(policy).toContain("frame-ancestors 'self'")
    expect(policy).toContain("script-src 'self' 'wasm-unsafe-eval'")
    expect(policy).toContain("worker-src 'self' blob:")
    expect(policy).toContain('connect-src')
    expect(policy).toContain('https://bilig.example.com')
    expect(policy).toContain('https://workbooks.example.com')
    expect(policy).toContain('https://zero.example.com')
  })
})
