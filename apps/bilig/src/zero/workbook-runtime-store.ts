import { SpreadsheetEngine, type EngineReplicaSnapshot } from '@bilig/core'
import { isWorkbookSnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import { applyWorkbookEvent, zql } from '@bilig/zero-sync'
import type { Row } from '@rocicorp/zero'
import { parseCheckpointPayload, parseCheckpointReplicaState, parsePositiveInteger } from './store-support.js'
import {
  loadWorkbookEventRecordsAfter,
  type Queryable,
  type WorkbookRuntimeMetadata,
  type WorkbookRuntimeState,
  type WorkbookRuntimeStoreConnection,
} from './store.js'

interface WorkbookCheckpointRecord {
  revision: number
  checkpointPayload: WorkbookSnapshot
  replicaState: EngineReplicaSnapshot | null
}

type ZeroWorkbookRow = Row['workbooks']

interface InlineWorkbookCheckpointRecord {
  checkpointPayload: WorkbookSnapshot | null
  replicaState: EngineReplicaSnapshot | null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function parseWorkbookRuntimeMetadata(row: ZeroWorkbookRow | undefined, documentId: string): WorkbookRuntimeMetadata {
  if (!row) {
    return {
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'system',
    }
  }

  if (!isSafeNonNegativeInteger(row.headRevision)) {
    throw new Error(`Invalid Zero workbook head revision for ${documentId}`)
  }
  if (!isSafeNonNegativeInteger(row.calculatedRevision)) {
    throw new Error(`Invalid Zero workbook calculated revision for ${documentId}`)
  }
  if (row.calculatedRevision > row.headRevision) {
    throw new Error(`Invalid Zero workbook revision order for ${documentId}`)
  }
  if (row.ownerUserId.trim() === '') {
    throw new Error(`Invalid Zero workbook owner for ${documentId}`)
  }

  return {
    headRevision: row.headRevision,
    calculatedRevision: row.calculatedRevision,
    ownerUserId: row.ownerUserId,
  }
}

async function loadZeroWorkbookMetadata(db: WorkbookRuntimeStoreConnection, documentId: string): Promise<WorkbookRuntimeMetadata> {
  const row = await db.run(zql.workbooks.where('id', documentId).one())
  return parseWorkbookRuntimeMetadata(row, documentId)
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
