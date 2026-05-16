import type { WorkbookAgentReviewQueueItem } from '@bilig/agent-api'
import type {
  WorkbookAgentExecutionPolicy,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
} from '@bilig/contracts'
import type { Queryable } from './store.js'
import {
  defaultExecutionPolicyForScope,
  hasToolCallState,
  isExecutionPolicy,
  isWorkbookAgentUiContext,
  normalizeReviewQueueItem,
  normalizeThreadSummary,
  normalizeTimelineEntry,
  normalizeTimelineToolName,
  normalizeToolCallRow,
  parseNumericValue,
  type WorkbookChatItemRow,
  type WorkbookChatThreadRow,
  type WorkbookChatThreadScope,
  type WorkbookChatThreadSummaryRow,
  type WorkbookChatToolCallRow,
  type WorkbookReviewQueueItemRow,
} from './workbook-chat-thread-normalizers.js'

export type { WorkbookChatThreadScope } from './workbook-chat-thread-normalizers.js'

export interface WorkbookAgentThreadStateRecord {
  readonly documentId: string
  readonly threadId: string
  readonly actorUserId: string
  readonly scope: WorkbookChatThreadScope
  readonly executionPolicy: WorkbookAgentExecutionPolicy
  readonly context: WorkbookAgentUiContext | null
  readonly entries: readonly WorkbookAgentTimelineEntry[]
  readonly reviewQueueItems: readonly WorkbookAgentReviewQueueItem[]
  readonly updatedAtUnixMs: number
}

function dedupeTimelineEntries(entries: readonly WorkbookAgentTimelineEntry[]): WorkbookAgentTimelineEntry[] {
  const deduped: WorkbookAgentTimelineEntry[] = []
  const indexById = new Map<string, number>()
  for (const entry of entries) {
    const existingIndex = indexById.get(entry.id)
    if (existingIndex === undefined) {
      indexById.set(entry.id, deduped.length)
      deduped.push(entry)
      continue
    }
    deduped[existingIndex] = entry
  }
  return deduped
}

export async function ensureWorkbookChatThreadSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_chat_thread (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      scope TEXT NOT NULL DEFAULT 'private',
      execution_policy TEXT NOT NULL DEFAULT 'autoApplyAll',
      context_json JSONB,
      entry_count BIGINT NOT NULL DEFAULT 0,
      latest_entry_text TEXT,
      updated_at_unix_ms BIGINT NOT NULL,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id)
    )
  `)
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS execution_policy TEXT;
  `)
  await db.query(`
    UPDATE workbook_chat_thread
    SET execution_policy = CASE WHEN scope = 'shared' THEN 'ownerReview' ELSE 'autoApplyAll' END
    WHERE execution_policy IS NULL;
  `)
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ALTER COLUMN execution_policy SET DEFAULT 'autoApplyAll';
  `)
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ALTER COLUMN execution_policy SET NOT NULL;
  `)
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS entry_count BIGINT NOT NULL DEFAULT 0;
  `)
  await db.query(`
    ALTER TABLE workbook_chat_thread
      ADD COLUMN IF NOT EXISTS latest_entry_text TEXT;
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_chat_item (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      entry_id TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      turn_id TEXT,
      kind TEXT NOT NULL,
      text TEXT,
      phase TEXT,
      tool_name TEXT,
      tool_status TEXT,
      arguments_text TEXT,
      output_text TEXT,
      success BOOLEAN,
      citations_json JSONB,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id, entry_id)
    )
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_chat_tool_call (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      entry_id TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      turn_id TEXT,
      tool_name TEXT,
      tool_status TEXT,
      arguments_text TEXT,
      output_text TEXT,
      success BOOLEAN,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id, entry_id)
    )
  `)
  await db.query(`
    ALTER TABLE workbook_chat_item
      ADD COLUMN IF NOT EXISTS citations_json JSONB;
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_review_queue_item (
      workbook_id TEXT NOT NULL,
      thread_id TEXT NOT NULL,
      actor_user_id TEXT NOT NULL,
      review_item_id TEXT NOT NULL,
      turn_id TEXT NOT NULL,
      goal_text TEXT NOT NULL,
      summary TEXT NOT NULL,
      scope TEXT NOT NULL,
      risk_class TEXT NOT NULL,
      review_mode TEXT NOT NULL,
      owner_user_id TEXT,
      status TEXT NOT NULL,
      decided_by_user_id TEXT,
      decided_at_unix_ms BIGINT,
      base_revision BIGINT NOT NULL,
      created_at_unix_ms BIGINT NOT NULL,
      context_json JSONB,
      commands_json JSONB NOT NULL,
      affected_ranges_json JSONB NOT NULL,
      estimated_affected_cells BIGINT,
      recommendations_json JSONB NOT NULL DEFAULT '[]'::jsonb,
      PRIMARY KEY (workbook_id, thread_id, actor_user_id, review_item_id)
    )
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_chat_thread_document_actor_updated_idx
      ON workbook_chat_thread (workbook_id, actor_user_id, updated_at_unix_ms DESC)
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_chat_tool_call_thread_order_idx
      ON workbook_chat_tool_call (workbook_id, thread_id, actor_user_id, sort_order ASC)
  `)
  await db.query(`
    CREATE INDEX IF NOT EXISTS workbook_review_queue_item_thread_created_idx
      ON workbook_review_queue_item (workbook_id, thread_id, actor_user_id, created_at_unix_ms ASC)
  `)
}

export async function saveWorkbookAgentThreadState(db: Queryable, record: WorkbookAgentThreadStateRecord): Promise<void> {
  const persistedEntries = dedupeTimelineEntries(record.entries)
  const latestEntryText =
    [...persistedEntries].toReversed().find((entry) => typeof entry.text === 'string' && entry.text.trim().length > 0)?.text ?? null
  await db.query(
    `
      INSERT INTO workbook_chat_thread (
        workbook_id,
        thread_id,
        actor_user_id,
        scope,
        execution_policy,
        context_json,
        entry_count,
        latest_entry_text,
        updated_at_unix_ms
      )
      VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9)
      ON CONFLICT (workbook_id, thread_id, actor_user_id)
      DO UPDATE SET
        scope = EXCLUDED.scope,
        execution_policy = EXCLUDED.execution_policy,
        context_json = EXCLUDED.context_json,
        entry_count = EXCLUDED.entry_count,
        latest_entry_text = EXCLUDED.latest_entry_text,
        updated_at_unix_ms = EXCLUDED.updated_at_unix_ms
    `,
    [
      record.documentId,
      record.threadId,
      record.actorUserId,
      record.scope,
      record.executionPolicy,
      JSON.stringify(record.context),
      persistedEntries.length,
      latestEntryText,
      record.updatedAtUnixMs,
    ],
  )
  await db.query(
    `
      DELETE FROM workbook_chat_item
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  )
  await db.query(
    `
      DELETE FROM workbook_chat_tool_call
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  )
  await Promise.all(
    persistedEntries.map(async (entry, index) => {
      await db.query(
        `
          INSERT INTO workbook_chat_item (
            workbook_id,
            thread_id,
            actor_user_id,
            entry_id,
            sort_order,
            turn_id,
            kind,
            text,
            phase,
            tool_name,
            tool_status,
            arguments_text,
            output_text,
            success,
            citations_json
          )
          VALUES (
            $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15::jsonb
          )
          ON CONFLICT (workbook_id, thread_id, actor_user_id, entry_id)
          DO UPDATE SET
            sort_order = EXCLUDED.sort_order,
            turn_id = EXCLUDED.turn_id,
            kind = EXCLUDED.kind,
            text = EXCLUDED.text,
            phase = EXCLUDED.phase,
            tool_name = EXCLUDED.tool_name,
            tool_status = EXCLUDED.tool_status,
            arguments_text = EXCLUDED.arguments_text,
            output_text = EXCLUDED.output_text,
            success = EXCLUDED.success,
            citations_json = EXCLUDED.citations_json
        `,
        [
          record.documentId,
          record.threadId,
          record.actorUserId,
          entry.id,
          index,
          entry.turnId,
          entry.kind,
          entry.text,
          entry.phase,
          normalizeTimelineToolName(entry.toolName),
          entry.toolStatus,
          entry.argumentsText,
          entry.outputText,
          entry.success,
          JSON.stringify(entry.citations),
        ],
      )
      if (!hasToolCallState(entry)) {
        return
      }
      await db.query(
        `
          INSERT INTO workbook_chat_tool_call (
            workbook_id,
            thread_id,
            actor_user_id,
            entry_id,
            sort_order,
            turn_id,
            tool_name,
            tool_status,
            arguments_text,
            output_text,
            success
          )
          VALUES (
            $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11
          )
          ON CONFLICT (workbook_id, thread_id, actor_user_id, entry_id)
          DO UPDATE SET
            sort_order = EXCLUDED.sort_order,
            turn_id = EXCLUDED.turn_id,
            tool_name = EXCLUDED.tool_name,
            tool_status = EXCLUDED.tool_status,
            arguments_text = EXCLUDED.arguments_text,
            output_text = EXCLUDED.output_text,
            success = EXCLUDED.success
        `,
        [
          record.documentId,
          record.threadId,
          record.actorUserId,
          entry.id,
          index,
          entry.turnId,
          normalizeTimelineToolName(entry.toolName),
          entry.toolStatus,
          entry.argumentsText,
          entry.outputText,
          entry.success,
        ],
      )
    }),
  )
  await db.query(
    `
      DELETE FROM workbook_review_queue_item
      WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
    `,
    [record.documentId, record.threadId, record.actorUserId],
  )
  await Promise.all(
    record.reviewQueueItems.map(async (reviewItem) => {
      await db.query(
        `
          INSERT INTO workbook_review_queue_item (
            workbook_id,
            thread_id,
            actor_user_id,
            review_item_id,
            turn_id,
            goal_text,
            summary,
            scope,
            risk_class,
            review_mode,
            owner_user_id,
            status,
            decided_by_user_id,
            decided_at_unix_ms,
            base_revision,
            created_at_unix_ms,
            context_json,
            commands_json,
            affected_ranges_json,
            estimated_affected_cells,
            recommendations_json
          )
          VALUES (
            $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17::jsonb, $18::jsonb, $19::jsonb, $20, $21::jsonb
          )
        `,
        [
          record.documentId,
          record.threadId,
          record.actorUserId,
          reviewItem.id,
          reviewItem.turnId,
          reviewItem.goalText,
          reviewItem.summary,
          reviewItem.scope,
          reviewItem.riskClass,
          reviewItem.reviewMode,
          reviewItem.ownerUserId,
          reviewItem.status,
          reviewItem.decidedByUserId,
          reviewItem.decidedAtUnixMs,
          reviewItem.baseRevision,
          reviewItem.createdAtUnixMs,
          JSON.stringify(reviewItem.context),
          JSON.stringify(reviewItem.commands),
          JSON.stringify(reviewItem.affectedRanges),
          reviewItem.estimatedAffectedCells,
          JSON.stringify(reviewItem.recommendations),
        ],
      )
    }),
  )
}

export async function loadWorkbookAgentThreadState(
  db: Queryable,
  input: {
    documentId: string
    threadId: string
    actorUserId: string
  },
): Promise<WorkbookAgentThreadStateRecord | null> {
  const threadResult = await db.query<WorkbookChatThreadRow>(
    `
      SELECT
        workbook_id AS "workbookId",
        thread_id AS "threadId",
        actor_user_id AS "actorUserId",
        scope AS "scope",
        execution_policy AS "executionPolicy",
        context_json AS "contextJson",
        updated_at_unix_ms AS "updatedAtUnixMs"
      FROM workbook_chat_thread
      WHERE workbook_id = $1
        AND thread_id = $2
        AND (actor_user_id = $3 OR scope = 'shared')
      ORDER BY
        CASE WHEN actor_user_id = $3 THEN 0 ELSE 1 END ASC,
        updated_at_unix_ms DESC
      LIMIT 1
    `,
    [input.documentId, input.threadId, input.actorUserId],
  )
  const thread = threadResult.rows[0]
  const updatedAtUnixMs = parseNumericValue(thread?.updatedAtUnixMs)
  const executionPolicy = isExecutionPolicy(thread?.executionPolicy)
    ? thread.executionPolicy
    : thread?.scope === 'private' || thread?.scope === 'shared'
      ? defaultExecutionPolicyForScope(thread.scope)
      : null
  if (
    !thread ||
    typeof thread.workbookId !== 'string' ||
    typeof thread.threadId !== 'string' ||
    typeof thread.actorUserId !== 'string' ||
    (thread.scope !== 'private' && thread.scope !== 'shared') ||
    updatedAtUnixMs === null ||
    executionPolicy === null
  ) {
    return null
  }
  const [itemResult, toolCallResult, reviewQueueItemResult] = await Promise.all([
    db.query<WorkbookChatItemRow>(
      `
        SELECT
          entry_id AS "entryId",
          turn_id AS "turnId",
          kind AS "kind",
          text AS "text",
          phase AS "phase",
          tool_name AS "toolName",
          tool_status AS "toolStatus",
          arguments_text AS "argumentsText",
          output_text AS "outputText",
          success AS "success",
          citations_json AS "citationsJson",
          sort_order AS "sortOrder"
        FROM workbook_chat_item
        WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
        ORDER BY sort_order ASC
      `,
      [thread.workbookId, thread.threadId, thread.actorUserId],
    ),
    db.query<WorkbookChatToolCallRow>(
      `
        SELECT
          entry_id AS "entryId",
          turn_id AS "turnId",
          tool_name AS "toolName",
          tool_status AS "toolStatus",
          arguments_text AS "argumentsText",
          output_text AS "outputText",
          success AS "success",
          sort_order AS "sortOrder"
        FROM workbook_chat_tool_call
        WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
        ORDER BY sort_order ASC
      `,
      [thread.workbookId, thread.threadId, thread.actorUserId],
    ),
    db.query<WorkbookReviewQueueItemRow>(
      `
        SELECT
          review_item_id AS "reviewItemId",
          workbook_id AS "workbookId",
          thread_id AS "threadId",
          actor_user_id AS "actorUserId",
          turn_id AS "turnId",
          goal_text AS "goalText",
          summary AS "summary",
          scope AS "scope",
          risk_class AS "riskClass",
          review_mode AS "reviewMode",
          owner_user_id AS "ownerUserId",
          status AS "status",
          decided_by_user_id AS "decidedByUserId",
          decided_at_unix_ms AS "decidedAtUnixMs",
          base_revision AS "baseRevision",
          created_at_unix_ms AS "createdAtUnixMs",
          context_json AS "contextJson",
          commands_json AS "commandsJson",
          affected_ranges_json AS "affectedRangesJson",
          estimated_affected_cells AS "estimatedAffectedCells",
          recommendations_json AS "recommendationsJson"
        FROM workbook_review_queue_item
        WHERE workbook_id = $1 AND thread_id = $2 AND actor_user_id = $3
        ORDER BY created_at_unix_ms ASC, review_item_id ASC
      `,
      [thread.workbookId, thread.threadId, thread.actorUserId],
    ),
  ])
  const toolCallsByEntryId = new Map(
    toolCallResult.rows.flatMap((row) => {
      const normalized = normalizeToolCallRow(row)
      return normalized ? [[normalized.entryId, normalized] as const] : []
    }),
  )
  const entries = itemResult.rows
    .map((row) =>
      normalizeTimelineEntry({
        ...row,
        ...(typeof row.entryId === 'string' ? toolCallsByEntryId.get(row.entryId) : undefined),
      }),
    )
    .filter((entry): entry is WorkbookAgentTimelineEntry => entry !== null)
  const reviewQueueItems = reviewQueueItemResult.rows.flatMap((row) => {
    const normalized = normalizeReviewQueueItem(row)
    return normalized ? [normalized] : []
  })
  return {
    documentId: thread.workbookId,
    threadId: thread.threadId,
    actorUserId: thread.actorUserId,
    scope: thread.scope,
    executionPolicy,
    context: isWorkbookAgentUiContext(thread.contextJson) ? thread.contextJson : null,
    entries,
    reviewQueueItems,
    updatedAtUnixMs,
  }
}

export async function listWorkbookAgentThreadSummaries(
  db: Queryable,
  input: {
    documentId: string
    actorUserId: string
  },
): Promise<WorkbookAgentThreadSummary[]> {
  const result = await db.query<WorkbookChatThreadSummaryRow>(
    `
      SELECT
        thread.thread_id AS "threadId",
        thread.scope AS "scope",
        thread.actor_user_id AS "ownerUserId",
        thread.updated_at_unix_ms AS "updatedAtUnixMs",
        COALESCE(item_counts.entry_count, 0) AS "entryCount",
        COALESCE(review_counts.review_queue_item_count, 0) AS "reviewQueueItemCount",
        latest_item.text AS "latestEntryText"
      FROM (
        SELECT ranked.workbook_id, ranked.thread_id, ranked.actor_user_id, ranked.scope, ranked.updated_at_unix_ms
        FROM (
          SELECT
            workbook_id,
            thread_id,
            actor_user_id,
            scope,
            updated_at_unix_ms,
            ROW_NUMBER() OVER (
              PARTITION BY thread_id
              ORDER BY
                CASE WHEN actor_user_id = $2 THEN 0 ELSE 1 END ASC,
                updated_at_unix_ms DESC
            ) AS row_rank
          FROM workbook_chat_thread
          WHERE workbook_id = $1
            AND (actor_user_id = $2 OR scope = 'shared')
        ) AS ranked
        WHERE ranked.row_rank = 1
      ) AS thread
      LEFT JOIN (
        SELECT workbook_id, thread_id, actor_user_id, COUNT(*)::integer AS entry_count
        FROM workbook_chat_item
        GROUP BY workbook_id, thread_id, actor_user_id
      ) AS item_counts
        ON item_counts.workbook_id = thread.workbook_id
       AND item_counts.thread_id = thread.thread_id
       AND item_counts.actor_user_id = thread.actor_user_id
      LEFT JOIN (
        SELECT workbook_id, thread_id, actor_user_id, COUNT(*)::integer AS review_queue_item_count
        FROM workbook_review_queue_item
        GROUP BY workbook_id, thread_id, actor_user_id
      ) AS review_counts
        ON review_counts.workbook_id = thread.workbook_id
       AND review_counts.thread_id = thread.thread_id
       AND review_counts.actor_user_id = thread.actor_user_id
      LEFT JOIN LATERAL (
        SELECT text
        FROM workbook_chat_item
        WHERE workbook_id = thread.workbook_id
          AND thread_id = thread.thread_id
          AND actor_user_id = thread.actor_user_id
          AND text IS NOT NULL
        ORDER BY sort_order DESC
        LIMIT 1
      ) AS latest_item
        ON TRUE
      WHERE thread.workbook_id = $1
        AND (thread.actor_user_id = $2 OR thread.scope = 'shared')
      ORDER BY thread.updated_at_unix_ms DESC
    `,
    [input.documentId, input.actorUserId],
  )
  return result.rows.map((row) => normalizeThreadSummary(row)).filter((row): row is WorkbookAgentThreadSummary => row !== null)
}
