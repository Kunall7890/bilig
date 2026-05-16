import { describe, expect, it } from 'vitest'

import {
  resolveDevAppRuntimeBuildEnabled,
  resolveDevAppServerMode,
  resolveDevCleanupCompose,
  resolveDevDisableCompose,
  resolveDevWebPreviewBuildEnabled,
  resolveDevWebServerMode,
} from '../dev-web-local-config.js'

describe('dev web local config', () => {
  it('defaults local web server mode to dev and accepts explicit modes', () => {
    expect(resolveDevWebServerMode({})).toBe('dev')
    expect(resolveDevWebServerMode({ BILIG_DEV_WEB_SERVER_MODE: 'dev' })).toBe('dev')
    expect(resolveDevWebServerMode({ BILIG_DEV_WEB_SERVER_MODE: 'preview' })).toBe('preview')
  })

  it('rejects malformed local web server modes', () => {
    expect(() => resolveDevWebServerMode({ BILIG_DEV_WEB_SERVER_MODE: 'prevew' })).toThrow(
      'BILIG_DEV_WEB_SERVER_MODE must be "dev" or "preview", got prevew',
    )
  })

  it('defaults local app server mode to watch and accepts explicit modes', () => {
    expect(resolveDevAppServerMode({})).toBe('watch')
    expect(resolveDevAppServerMode({ BILIG_DEV_APP_SERVER_MODE: 'watch' })).toBe('watch')
    expect(resolveDevAppServerMode({ BILIG_DEV_APP_SERVER_MODE: 'run' })).toBe('run')
  })

  it('rejects malformed local app server modes', () => {
    expect(() => resolveDevAppServerMode({ BILIG_DEV_APP_SERVER_MODE: 'rn' })).toThrow(
      'BILIG_DEV_APP_SERVER_MODE must be "watch" or "run", got rn',
    )
  })

  it('resolves local dev boolean flags with explicit defaults', () => {
    expect(resolveDevDisableCompose({})).toBe(false)
    expect(resolveDevDisableCompose({ BILIG_DEV_DISABLE_COMPOSE: '1' })).toBe(true)
    expect(resolveDevAppRuntimeBuildEnabled({})).toBe(true)
    expect(resolveDevAppRuntimeBuildEnabled({ BILIG_DEV_APP_RUNTIME_BUILD: '0' })).toBe(false)
    expect(resolveDevWebPreviewBuildEnabled({})).toBe(true)
    expect(resolveDevWebPreviewBuildEnabled({ BILIG_DEV_WEB_PREVIEW_BUILD: 'false' })).toBe(false)
    expect(resolveDevCleanupCompose({})).toBe(false)
    expect(resolveDevCleanupCompose({ BILIG_DEV_CLEANUP_COMPOSE: 'true' })).toBe(true)
  })

  it('rejects malformed local dev boolean flags', () => {
    expect(() => resolveDevDisableCompose({ BILIG_DEV_DISABLE_COMPOSE: 'yes' })).toThrow(
      'BILIG_DEV_DISABLE_COMPOSE must be "1", "true", "0", or "false" when set, got yes',
    )
    expect(() => resolveDevAppRuntimeBuildEnabled({ BILIG_DEV_APP_RUNTIME_BUILD: 'no' })).toThrow(
      'BILIG_DEV_APP_RUNTIME_BUILD must be "1", "true", "0", or "false" when set, got no',
    )
  })
})
