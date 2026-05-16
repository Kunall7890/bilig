import { createSyncServer } from './http/sync-server.js'
import { DocumentSessionManager } from './workbook-runtime/document-session-manager.js'
import { SyncDocumentSupervisor } from './workbook-runtime/sync-document-supervisor.js'
import { LocalDocumentSupervisor } from './workbook-runtime/local-document-supervisor.js'
import { LocalWorkbookSessionManager } from './workbook-runtime/local-workbook-session-manager.js'
import { createInProcessWorksheetExecutor } from './workbook-runtime/worksheet-executor.js'
import { createZeroSyncService } from './zero/service.js'
import { createWorkbookAgentService } from './codex-app/workbook-agent-service.js'
import { logError } from './runtime-logger.js'
import { resolveBiligAppRuntimeConfig } from './app-runtime-config.js'

async function main() {
  const { host, appPort, publicServerUrl, browserAppBaseUrl, maxImportBytes } = resolveBiligAppRuntimeConfig()
  const worksheetHostSessionManager = new LocalWorkbookSessionManager({
    publicServerUrl,
    browserAppBaseUrl,
    ...(maxImportBytes !== undefined ? { maxImportBytes } : {}),
  })
  const worksheetHostDocumentService = new LocalDocumentSupervisor(worksheetHostSessionManager)

  const sessionManager = new DocumentSessionManager(
    undefined,
    undefined,
    createInProcessWorksheetExecutor({
      documentService: worksheetHostDocumentService,
      serverUrl: publicServerUrl,
      browserAppBaseUrl,
    }),
    {
      publicServerUrl,
      browserAppBaseUrl,
      ...(maxImportBytes !== undefined ? { maxImportBytes } : {}),
    },
  )
  const documentService = new SyncDocumentSupervisor(sessionManager)
  const zeroSyncService = createZeroSyncService()
  const workbookAgentService = createWorkbookAgentService(zeroSyncService)

  await zeroSyncService.initialize()

  const { app: syncApp } = createSyncServer({
    sessionManager,
    documentService,
    zeroSyncService,
    workbookAgentService,
  })

  try {
    await syncApp.listen({ host, port: appPort })
    syncApp.log.info({ host, appPort, zeroSync: zeroSyncService.enabled }, 'bilig app listening')
  } catch (error) {
    try {
      await workbookAgentService.close()
    } catch (closeError) {
      logError('Failed to close workbook agent service', closeError)
    }
    try {
      await zeroSyncService.close()
    } catch (closeError) {
      logError('Failed to close zero sync service', closeError)
    }
    logError(error)
    process.exit(1)
  }
}

void (async () => {
  try {
    await main()
  } catch (error) {
    logError(error)
    process.exit(1)
  }
})()
