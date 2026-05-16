import type { BiligRuntimeConfig } from '@bilig/zero-sync'

export interface RuntimeConfig {
  documentId: string
  persistState: boolean
  currentUserId: string
  workbookAgentEnabled: boolean
}

export function createLocalOnlyRuntimeConfig(currentUserId = 'local:user'): BiligRuntimeConfig {
  return {
    zeroCacheUrl: '/zero',
    defaultDocumentId: 'local-workbook',
    persistState: true,
    currentUserId,
    workbookAgentEnabled: false,
  }
}

export function normalizeRuntimeConfigUserId<T extends { currentUserId: string }>(
  config: T,
  session: {
    readonly userId: string
  },
): T {
  if (config.currentUserId === session.userId) {
    return config
  }
  return {
    ...config,
    currentUserId: session.userId,
  }
}

function resolvePersistState(configuredPersistState: boolean, searchParams: URLSearchParams): boolean {
  const explicitPersistState = searchParams.get('persist')
  if (explicitPersistState === '0' || explicitPersistState === 'false') {
    return false
  }
  if (explicitPersistState === '1' || explicitPersistState === 'true') {
    return true
  }
  return configuredPersistState
}

export function resolveRuntimeConfig(config: BiligRuntimeConfig): RuntimeConfig {
  const searchParams = typeof window === 'undefined' ? new URLSearchParams() : new URLSearchParams(window.location.search)
  const explicitDocumentId = searchParams.get('document')
  const persistState = resolvePersistState(config.persistState, searchParams)

  if (explicitDocumentId) {
    return {
      documentId: explicitDocumentId,
      persistState,
      currentUserId: config.currentUserId,
      workbookAgentEnabled: config.workbookAgentEnabled === true,
    }
  }

  return {
    documentId: config.defaultDocumentId,
    persistState,
    currentUserId: config.currentUserId,
    workbookAgentEnabled: config.workbookAgentEnabled === true,
  }
}

export function resolveRemoteSyncEnabled(env: { readonly DEV?: boolean; readonly VITE_BILIG_REMOTE_SYNC?: string | undefined }): boolean {
  const configured = env.VITE_BILIG_REMOTE_SYNC
  if (configured !== undefined) {
    return configured !== '0'
  }
  return env.DEV !== true
}
