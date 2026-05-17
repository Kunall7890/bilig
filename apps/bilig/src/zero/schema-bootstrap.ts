import { ensureWorkbookPresenceSchema } from './presence-store.js'
import { ensureWorkbookChangeSchema } from './workbook-change-store.js'
import { ensureWorkbookAgentRunSchema } from './workbook-agent-run-store.js'
import { ensureWorkbookChatThreadSchema } from './workbook-chat-thread-store.js'
import { ensureWorkbookWorkflowRunSchema } from './workbook-workflow-run-store.js'
import { ensureZeroSyncSchema } from './zero-schema-store.js'
import type { Queryable } from './store.js'

export async function ensureZeroServiceSchema(db: Queryable): Promise<void> {
  await ensureZeroSyncSchema(db)
  await ensureWorkbookPresenceSchema(db)
  await ensureWorkbookChangeSchema(db)
  await ensureWorkbookAgentRunSchema(db)
  await ensureWorkbookChatThreadSchema(db)
  await ensureWorkbookWorkflowRunSchema(db)
}
