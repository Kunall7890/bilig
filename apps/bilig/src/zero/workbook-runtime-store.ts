import { SpreadsheetEngine, type EngineReplicaSnapshot } from '@bilig/core'
import { isWorkbookSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import { applyWorkbookEvent } from '@bilig/zero-sync'
import { parseCheckpointPayload, parseCheckpointReplicaState, parseNonNegativeInteger, parsePositiveInteger } from './store-support.js'
import {
  loadWorkbookEventRecordsAfter,
  type Queryable,
  type QueryResultRow,
  type WorkbookRuntimeMetadata,
  type WorkbookRuntimeState,
  type WorkbookRuntimeStoreConnection,
} from './store.js'

interface WorkbookCheckpointRecord {
  revision: number
  checkpointPayload: WorkbookSnapshot
  replicaState: EngineReplicaSnapshot | null
}

interface InlineWorkbookCheckpointRecord {
  checkpointPayload: WorkbookSnapshot | null
  replicaState: EngineReplicaSnapshot | null
}

interface WorkbookMetadataRow extends QueryResultRow {
  head_revision?: unknown
  calculated_revision?: unknown
  owner_user_id?: unknown
}

function parseWorkbookRuntimeMetadata(row: WorkbookMetadataRow | undefined, documentId: string): WorkbookRuntimeMetadata {
  if (!row) {
    return {
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'system',
    }
  }

  const headRevision = parseNonNegativeInteger(row.head_revision)
  const calculatedRevision = parseNonNegativeInteger(row.calculated_revision)
  if (headRevision === null) {
    throw new Error(`Invalid Zero workbook head revision for ${documentId}`)
  }
  if (calculatedRevision === null) {
    throw new Error(`Invalid Zero workbook calculated revision for ${documentId}`)
  }
  if (calculatedRevision > headRevision) {
    throw new Error(`Invalid Zero workbook revision order for ${documentId}`)
  }
  if (typeof row.owner_user_id !== 'string' || row.owner_user_id.trim() === '') {
    throw new Error(`Invalid Zero workbook owner for ${documentId}`)
  }

  return {
    headRevision,
    calculatedRevision,
    ownerUserId: row.owner_user_id,
  }
}

async function loadZeroWorkbookMetadata(db: WorkbookRuntimeStoreConnection, documentId: string): Promise<WorkbookRuntimeMetadata> {
  const result = await db.query<WorkbookMetadataRow>(
    `
      SELECT head_revision, calculated_revision, owner_user_id
      FROM workbooks
      WHERE id = $1
      LIMIT 1
    `,
    [documentId],
  )
  return parseWorkbookRuntimeMetadata(result.rows[0], documentId)
}

async function loadInlineWorkbookCheckpoint(db: Queryable, documentId: string): Promise<InlineWorkbookCheckpointRecord> {
  const result = await db.query<{
    snapshot: unknown
    replica_snapshot: unknown
  }>(`SELECT snapshot, replica_snapshot FROM workbooks WHERE id = $1 LIMIT 1`, [documentId])
  const row = result.rows[0]
  return {
    checkpointPayload: isWorkbookSnapshot(row?.snapshot) ? row.snapshot : null,
    replicaState: parseCheckpointReplicaState(row?.replica_snapshot),
  }
}

async function loadLatestWorkbookCheckpoint(db: Queryable, documentId: string): Promise<WorkbookCheckpointRecord | null> {
  const result = await db.query<{
    revision: number | string | null
    payload: unknown
    replica_snapshot: unknown
  }>(
    `
      SELECT revision, payload, replica_snapshot
      FROM workbook_snapshot
      WHERE workbook_id = $1
        AND format = 'json-v1'
      ORDER BY revision DESC
      LIMIT 1
    `,
    [documentId],
  )
  const row = result.rows[0]
  if (!row || !isWorkbookSnapshot(row.payload)) {
    return null
  }
  const revision = parsePositiveInteger(row.revision)
  if (revision === null) {
    return null
  }
  return {
    revision,
    checkpointPayload: row.payload,
    replicaState: parseCheckpointReplicaState(row.replica_snapshot),
  }
}

export async function loadWorkbookState(db: WorkbookRuntimeStoreConnection, documentId: string): Promise<WorkbookRuntimeState> {
  const metadata = await loadZeroWorkbookMetadata(db, documentId)
  const { checkpointPayload: inlineCheckpointPayload, replicaState: inlineReplicaState } = await loadInlineWorkbookCheckpoint(
    db,
    documentId,
  )
  const { headRevision, calculatedRevision, ownerUserId } = metadata

  if (inlineCheckpointPayload) {
    return {
      snapshot: inlineCheckpointPayload,
      replicaSnapshot: inlineReplicaState,
      headRevision,
      calculatedRevision,
      ownerUserId,
    }
  }

  const checkpoint = await loadLatestWorkbookCheckpoint(db, documentId)
  const baseRevision = checkpoint?.revision ?? 0
  const baseCheckpointPayload = parseCheckpointPayload(checkpoint?.checkpointPayload, documentId)
  const baseReplicaState = parseCheckpointReplicaState(checkpoint?.replicaState)

  if (headRevision <= baseRevision) {
    return {
      snapshot: baseCheckpointPayload,
      replicaSnapshot: baseReplicaState,
      headRevision,
      calculatedRevision,
      ownerUserId,
    }
  }

  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `checkpoint-replay:${documentId}:${headRevision}`,
  })
  await engine.ready()
  engine.importSnapshot(baseCheckpointPayload)
  if (baseReplicaState) {
    engine.importReplicaSnapshot(baseReplicaState)
  }
  const events = await loadWorkbookEventRecordsAfter(db, documentId, baseRevision)
  for (const event of events) {
    applyWorkbookEvent(engine, event.payload)
  }

  return {
    snapshot: engine.exportSnapshot(),
    replicaSnapshot: null,
    headRevision,
    calculatedRevision,
    ownerUserId,
  }
}

export async function loadWorkbookRuntimeMetadata(
  db: WorkbookRuntimeStoreConnection,
  documentId: string,
): Promise<WorkbookRuntimeMetadata> {
  return await loadZeroWorkbookMetadata(db, documentId)
}

export async function acquireWorkbookMutationLock(db: Queryable, documentId: string): Promise<void> {
  await db.query(`SELECT pg_advisory_xact_lock(hashtext($1))`, [documentId])
}
