import { describe, expect, it } from 'vitest'

import { resolveBiligAppRuntimeConfig } from '../app-runtime-config.js'

describe('bilig app runtime config', () => {
  it('uses safe defaults', () => {
    const config = resolveBiligAppRuntimeConfig({})

    expect(config).toStrictEqual({
      host: '0.0.0.0',
      appPort: 4321,
      publicServerUrl: 'http://127.0.0.1:4321',
      browserAppBaseUrl: 'http://127.0.0.1:4321',
    })
  })

  it('preserves explicit runtime configuration', () => {
    const config = resolveBiligAppRuntimeConfig({
      HOST: '127.0.0.1',
      PORT: '4567',
      BILIG_PUBLIC_SERVER_URL: ' https://api.example.test ',
      BILIG_WEB_APP_BASE_URL: ' https://workbooks.example.test ',
      BILIG_AGENT_IMPORT_MAX_BYTES: '1048576',
    })

    expect(config).toStrictEqual({
      host: '127.0.0.1',
      appPort: 4567,
      publicServerUrl: 'https://api.example.test',
      browserAppBaseUrl: 'https://workbooks.example.test',
      maxImportBytes: 1_048_576,
    })
  })

  it.each(['4321abc', '0', '70000', '-1', ''])('rejects malformed PORT=%s', (port) => {
    expect(() => resolveBiligAppRuntimeConfig({ PORT: port })).toThrow('PORT must be a TCP port between 1 and 65535')
  })

  it.each(['10mb', '0', '-1', '9007199254740992'])('rejects malformed BILIG_AGENT_IMPORT_MAX_BYTES=%s', (maxImportBytes) => {
    expect(() => resolveBiligAppRuntimeConfig({ BILIG_AGENT_IMPORT_MAX_BYTES: maxImportBytes })).toThrow(
      /BILIG_AGENT_IMPORT_MAX_BYTES must be/u,
    )
  })

  it.each([
    ['BILIG_PUBLIC_SERVER_URL', ''],
    ['BILIG_PUBLIC_SERVER_URL', '/relative'],
    ['BILIG_PUBLIC_SERVER_URL', 'ftp://api.example.test'],
    ['BILIG_WEB_APP_BASE_URL', '   '],
    ['BILIG_WEB_APP_BASE_URL', 'workbooks.example.test'],
    ['BILIG_WEB_APP_BASE_URL', 'file:///tmp/app'],
  ])('rejects malformed %s=%s', (name, value) => {
    expect(() => resolveBiligAppRuntimeConfig({ [name]: value })).toThrow(`${name} must be an absolute http(s) URL`)
  })
})
