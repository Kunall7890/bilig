import { parseStrictBooleanEnvFlag } from './strict-env.js'

export type DevWebServerMode = 'dev' | 'preview'
export type DevAppServerMode = 'run' | 'watch'

export function resolveDevWebServerMode(env: { BILIG_DEV_WEB_SERVER_MODE?: string | undefined }): DevWebServerMode {
  const value = env.BILIG_DEV_WEB_SERVER_MODE
  if (value === undefined || value === 'dev') {
    return 'dev'
  }
  if (value === 'preview') {
    return 'preview'
  }

  throw new Error(`BILIG_DEV_WEB_SERVER_MODE must be "dev" or "preview", got ${value}`)
}

export function resolveDevAppServerMode(env: { BILIG_DEV_APP_SERVER_MODE?: string | undefined }): DevAppServerMode {
  const value = env.BILIG_DEV_APP_SERVER_MODE
  if (value === undefined || value === 'watch') {
    return 'watch'
  }
  if (value === 'run') {
    return 'run'
  }

  throw new Error(`BILIG_DEV_APP_SERVER_MODE must be "watch" or "run", got ${value}`)
}

export function resolveDevDisableCompose(env: { BILIG_DEV_DISABLE_COMPOSE?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_DEV_DISABLE_COMPOSE, 'BILIG_DEV_DISABLE_COMPOSE', false)
}

export function resolveDevAppRuntimeBuildEnabled(env: { BILIG_DEV_APP_RUNTIME_BUILD?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_DEV_APP_RUNTIME_BUILD, 'BILIG_DEV_APP_RUNTIME_BUILD', true)
}

export function resolveDevWebPreviewBuildEnabled(env: { BILIG_DEV_WEB_PREVIEW_BUILD?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_DEV_WEB_PREVIEW_BUILD, 'BILIG_DEV_WEB_PREVIEW_BUILD', true)
}

export function resolveDevCleanupCompose(env: { BILIG_DEV_CLEANUP_COMPOSE?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_DEV_CLEANUP_COMPOSE, 'BILIG_DEV_CLEANUP_COMPOSE', false)
}
