import sqlite3InitModule, { type Database, type SAHPoolUtil, type SqlValue, type Sqlite3Static } from '@sqlite.org/sqlite-wasm'
import type {
  WorkbookLocalAuthoritativeDelta,
  WorkbookLocalAuthoritativeBase,
  WorkbookLocalProjectionOverlay,
  WorkbookLocalViewportBase,
} from './workbook-local-base.js'
import {
  readWorkbookViewportProjection,
  writeWorkbookAuthoritativeBase,
  writeWorkbookAuthoritativeDelta,
  writeWorkbookProjectionOverlay,
} from './workbook-local-store-projection.js'
import { initializeWorkbookLocalStoreSchema } from './workbook-local-store-schema.js'
import { WorkbookLocalMutationJournalStore, type WorkbookLocalMutationRecord } from './workbook-local-mutation-journal-store.js'

export type { WorkbookLocalMutationRecord } from './workbook-local-mutation-journal-store.js'

const WORKBOOK_VFS_NAME = 'bilig-opfs-sahpool'
const WORKBOOK_VFS_DIRECTORY = '/bilig/workbooks'
const WORKBOOK_VFS_INITIAL_CAPACITY = 12
const WORKBOOK_VFS_VERBOSITY = 0

let sqliteRuntimePromise: Promise<{ sqlite3: Sqlite3Static; poolUtil: SAHPoolUtil }> | null = null
let memorySqliteRuntimePromise: Promise<Sqlite3Static> | null = null

export class WorkbookLocalStoreLockedError extends Error {
  override readonly name = 'WorkbookLocalStoreLockedError'
}

export class WorkbookLocalStoreUnavailableError extends Error {
  override readonly name = 'WorkbookLocalStoreUnavailableError'
}

export interface WorkbookStoredState {
  readonly snapshot: unknown
  readonly replica: unknown
  readonly authoritativeRevision: number
  readonly appliedPendingLocalSeq: number
}

export interface WorkbookBootstrapState {
  readonly workbookName: string
  readonly sheetNames: readonly string[]
  readonly materializedCellCount: number
  readonly authoritativeRevision: number
  readonly appliedPendingLocalSeq: number
}

export interface WorkbookLocalStore {
  loadBootstrapState(): Promise<WorkbookBootstrapState | null>
  loadState(): Promise<WorkbookStoredState | null>
  persistProjectionState(input: {
    readonly state: WorkbookStoredState
    readonly authoritativeBase: WorkbookLocalAuthoritativeBase
    readonly projectionOverlay: WorkbookLocalProjectionOverlay
  }): Promise<void>
  ingestAuthoritativeDelta(input: {
    readonly state: WorkbookStoredState
    readonly authoritativeDelta: WorkbookLocalAuthoritativeDelta
    readonly projectionOverlay: WorkbookLocalProjectionOverlay
    readonly removePendingMutationIds?: readonly string[]
  }): Promise<void>
  listPendingMutations(): Promise<WorkbookLocalMutationRecord[]>
  listMutationJournalEntries(): Promise<WorkbookLocalMutationRecord[]>
  appendPendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void>
  updatePendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void>
  removePendingMutation(id: string): Promise<void>
  readViewportProjection(
    sheetName: string,
    viewport: {
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    },
  ): WorkbookLocalViewportBase | null
  close(): void
}

export interface WorkbookLocalStoreFactory {
  open(documentId: string): Promise<WorkbookLocalStore>
}

export interface OpfsWorkbookLocalStoreFactoryOptions {
  vfsName?: string
  directory?: string
  initialCapacity?: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function supportsWorkerOpfs(): boolean {
  const scope = globalThis as typeof globalThis & {
    navigator?: Navigator
    document?: Document
  }
  if (typeof scope.document !== 'undefined') {
    return false
  }
  return typeof scope.navigator?.storage?.getDirectory === 'function'
}

function sanitizeDocumentId(documentId: string): string {
  return encodeURIComponent(documentId).replaceAll('%', '_')
}

function getErrorName(error: unknown): string | null {
  return isRecord(error) && typeof error['name'] === 'string' ? error['name'] : null
}

function getErrorMessage(error: unknown): string | null {
  return isRecord(error) && typeof error['message'] === 'string' ? error['message'] : null
}

function stringifyConsoleArg(value: unknown): string {
  if (typeof value === 'string') {
    return value
  }
  const errorName = getErrorName(value)
  const errorMessage = getErrorMessage(value)
  if (errorName && errorMessage) {
    return `${errorName}: ${errorMessage}`
  }
  try {
    return JSON.stringify(value)
  } catch {
    return String(value)
  }
}

function isOpfsLockDiagnostic(args: readonly unknown[], vfsName: string): boolean {
  const message = args.map((value) => stringifyConsoleArg(value)).join(' ')
  if (!message.includes(vfsName)) {
    return false
  }
  return message.includes('removeVfs() failed with no recovery strategy') || message.includes('Access Handles cannot be created')
}

type SqliteLogger = (...args: unknown[]) => void

function getSqliteConfigLogger(sqlite3: Sqlite3Static, key: 'error' | 'warn'): SqliteLogger | null {
  if (!isRecord(sqlite3) || !isRecord(sqlite3['config'])) {
    return null
  }
  const logger = sqlite3['config'][key]
  return typeof logger === 'function' ? (logger as SqliteLogger) : null
}

function setSqliteConfigLogger(sqlite3: Sqlite3Static, key: 'error' | 'warn', logger: SqliteLogger | null): void {
  if (!logger || !isRecord(sqlite3) || !isRecord(sqlite3['config'])) {
    return
  }
  sqlite3['config'][key] = logger
}

async function withSuppressedOpfsLockDiagnostics<T>(sqlite3: Sqlite3Static, vfsName: string, task: () => Promise<T>): Promise<T> {
  const globalConsole = globalThis.console
  const originalError = globalConsole.error
  const originalWarn = globalConsole.warn
  const originalSqliteError = getSqliteConfigLogger(sqlite3, 'error')
  const originalSqliteWarn = getSqliteConfigLogger(sqlite3, 'warn')
  globalConsole.error = ((...args: unknown[]) => {
    if (!isOpfsLockDiagnostic(args, vfsName)) {
      originalError(...args)
    }
  }) as typeof globalConsole.error
  globalConsole.warn = ((...args: unknown[]) => {
    if (!isOpfsLockDiagnostic(args, vfsName)) {
      originalWarn(...args)
    }
  }) as typeof globalConsole.warn
  if (originalSqliteError) {
    setSqliteConfigLogger(sqlite3, 'error', (...args: unknown[]) => {
      if (!isOpfsLockDiagnostic(args, vfsName)) {
        originalSqliteError(...args)
      }
    })
  }
  if (originalSqliteWarn) {
    setSqliteConfigLogger(sqlite3, 'warn', (...args: unknown[]) => {
      if (!isOpfsLockDiagnostic(args, vfsName)) {
        originalSqliteWarn(...args)
      }
    })
  }
  try {
    return await task()
  } finally {
    globalConsole.error = originalError
    globalConsole.warn = originalWarn
    setSqliteConfigLogger(sqlite3, 'error', originalSqliteError)
    setSqliteConfigLogger(sqlite3, 'warn', originalSqliteWarn)
  }
}

function isAccessHandleConflict(error: unknown): boolean {
  const message = getErrorMessage(error)
  const name = getErrorName(error)
  if (!message || !name) {
    return false
  }
  return (
    (message.includes('createSyncAccessHandle') && message.includes('Access Handles cannot be created')) ||
    (name === 'NoModificationAllowedError' &&
      (message.includes('Access Handles cannot be created') ||
        (message.includes('removeEntry') && message.includes('FileSystemDirectoryHandle'))))
  )
}

function isSqliteOpenUnavailable(error: unknown): boolean {
  const message = getErrorMessage(error)
  if (!message) {
    return false
  }
  return message.includes('SQLITE_CANTOPEN') || message.includes('unable to open database file')
}

function toWorkbookLocalStoreLockedError(documentId: string | null, cause?: unknown): WorkbookLocalStoreLockedError {
  const suffix = documentId ? ` for ${documentId}` : ''
  const error = new WorkbookLocalStoreLockedError(`Workbook local store is locked by another tab${suffix}`)
  if (cause !== undefined) {
    Object.defineProperty(error, 'cause', {
      configurable: true,
      enumerable: false,
      value: cause,
      writable: true,
    })
  }
  return error
}

function toWorkbookLocalStoreUnavailableError(documentId: string | null, cause?: unknown): WorkbookLocalStoreUnavailableError {
  const suffix = documentId ? ` for ${documentId}` : ''
  const error = new WorkbookLocalStoreUnavailableError(`Workbook local store is unavailable${suffix}`)
  if (cause !== undefined) {
    Object.defineProperty(error, 'cause', {
      configurable: true,
      enumerable: false,
      value: cause,
      writable: true,
    })
  }
  return error
}

function parseWorkbookStoredState(value: unknown): WorkbookStoredState | null {
  if (
    !isRecord(value) ||
    !isSafeNonNegativeInteger(value['authoritativeRevision']) ||
    !isSafeNonNegativeInteger(value['appliedPendingLocalSeq']) ||
    !isRecord(value['snapshot']) ||
    !isRecord(value['replica'])
  ) {
    return null
  }
  return {
    snapshot: value['snapshot'],
    replica: value['replica'],
    authoritativeRevision: value['authoritativeRevision'],
    appliedPendingLocalSeq: value['appliedPendingLocalSeq'],
  }
}

function parseWorkbookBootstrapState(value: unknown): WorkbookBootstrapState | null {
  if (
    !isRecord(value) ||
    typeof value['workbookName'] !== 'string' ||
    !Array.isArray(value['sheetNames']) ||
    !value['sheetNames'].every((sheetName) => typeof sheetName === 'string') ||
    !isSafeNonNegativeInteger(value['materializedCellCount']) ||
    !isSafeNonNegativeInteger(value['authoritativeRevision']) ||
    !isSafeNonNegativeInteger(value['appliedPendingLocalSeq'])
  ) {
    return null
  }
  return {
    workbookName: value['workbookName'],
    sheetNames: [...value['sheetNames']],
    materializedCellCount: value['materializedCellCount'],
    authoritativeRevision: value['authoritativeRevision'],
    appliedPendingLocalSeq: value['appliedPendingLocalSeq'],
  }
}

function readSingleObjectRow(db: Database, sql: string, bind?: readonly SqlValue[]): Record<string, SqlValue> | null {
  const statement = db.prepare(sql)
  try {
    if (bind) {
      statement.bind([...bind])
    }
    if (!statement.step()) {
      return null
    }
    return statement.get({})
  } finally {
    statement.finalize()
  }
}

function assertWorkbookStoredState(value: WorkbookStoredState): void {
  if (!parseWorkbookStoredState(value)) {
    throw new TypeError('Invalid workbook runtime state')
  }
}

async function getSqliteRuntime(
  options: Required<OpfsWorkbookLocalStoreFactoryOptions>,
): Promise<{ sqlite3: Sqlite3Static; poolUtil: SAHPoolUtil }> {
  if (!supportsWorkerOpfs()) {
    throw new Error('Workbook local storage requires a worker with OPFS support')
  }
  if (!sqliteRuntimePromise) {
    sqliteRuntimePromise = (async () => {
      const sqlite3 = await sqlite3InitModule()
      try {
        const poolUtil = await withSuppressedOpfsLockDiagnostics(
          sqlite3,
          options.vfsName,
          async () =>
            await sqlite3.installOpfsSAHPoolVfs({
              name: options.vfsName,
              directory: options.directory,
              initialCapacity: options.initialCapacity,
              // The app translates lock conflicts into its own persistence-state path.
              // Keep sqlite-wasm from spamming expected pool diagnostics into the console.
              verbosity: WORKBOOK_VFS_VERBOSITY,
            } as Parameters<Sqlite3Static['installOpfsSAHPoolVfs']>[0] & { verbosity: number }),
        )
        return { sqlite3, poolUtil }
      } catch (error) {
        if (isAccessHandleConflict(error)) {
          throw toWorkbookLocalStoreLockedError(null, error)
        }
        if (isSqliteOpenUnavailable(error)) {
          throw toWorkbookLocalStoreUnavailableError(null, error)
        }
        throw error
      }
    })()
  }
  try {
    return await sqliteRuntimePromise
  } catch (error) {
    if (!(error instanceof WorkbookLocalStoreLockedError)) {
      sqliteRuntimePromise = null
    }
    throw error
  }
}

function extractWorkbookName(snapshot: unknown): string | null {
  if (isRecord(snapshot) && isRecord(snapshot['workbook']) && typeof snapshot['workbook']['name'] === 'string') {
    return snapshot['workbook']['name']
  }
  return null
}

class SqliteWorkbookLocalStore implements WorkbookLocalStore {
  private readonly mutationJournal: WorkbookLocalMutationJournalStore

  constructor(
    private readonly db: Database,
    private readonly defaultWorkbookName: string,
    private readonly closeDbOnClose = true,
  ) {
    this.mutationJournal = new WorkbookLocalMutationJournalStore(db)
  }

  async loadBootstrapState(): Promise<WorkbookBootstrapState | null> {
    const row = readSingleObjectRow(
      this.db,
      `
        SELECT workbook_name AS workbookName,
               (SELECT COUNT(*) FROM authoritative_cell_render) AS materializedCellCount,
               authoritative_revision AS authoritativeRevision,
               applied_pending_local_seq AS appliedPendingLocalSeq
          FROM runtime_state
         WHERE id = 1
      `,
    )
    if (!row) {
      return null
    }

    const sheetNames: string[] = []
    const statement = this.db.prepare(
      `
        SELECT name
          FROM authoritative_sheet
         ORDER BY sort_order ASC, name ASC
      `,
    )
    try {
      while (statement.step()) {
        const sheet = statement.get({})
        if (typeof sheet['name'] === 'string') {
          sheetNames.push(sheet['name'])
        }
      }
    } finally {
      statement.finalize()
    }

    return parseWorkbookBootstrapState({
      workbookName:
        typeof row['workbookName'] === 'string' && row['workbookName'].length > 0 ? row['workbookName'] : this.defaultWorkbookName,
      sheetNames,
      materializedCellCount: typeof row['materializedCellCount'] === 'number' ? row['materializedCellCount'] : 0,
      authoritativeRevision: row['authoritativeRevision'],
      appliedPendingLocalSeq: row['appliedPendingLocalSeq'],
    })
  }

  async loadState(): Promise<WorkbookStoredState | null> {
    const row = readSingleObjectRow(
      this.db,
      `
        SELECT snapshot_json AS snapshotJson,
               replica_json AS replicaJson,
               authoritative_revision AS authoritativeRevision,
               applied_pending_local_seq AS appliedPendingLocalSeq
          FROM runtime_state
         WHERE id = 1
      `,
    )
    if (!row) {
      return null
    }
    const snapshotJson = row['snapshotJson']
    const replicaJson = row['replicaJson']
    const authoritativeRevision = row['authoritativeRevision']
    const appliedPendingLocalSeq = row['appliedPendingLocalSeq']
    if (
      typeof snapshotJson !== 'string' ||
      typeof replicaJson !== 'string' ||
      typeof authoritativeRevision !== 'number' ||
      typeof appliedPendingLocalSeq !== 'number'
    ) {
      return null
    }
    try {
      return parseWorkbookStoredState({
        snapshot: JSON.parse(snapshotJson) as unknown,
        replica: JSON.parse(replicaJson) as unknown,
        authoritativeRevision,
        appliedPendingLocalSeq,
      })
    } catch {
      return null
    }
  }

  async persistProjectionState(input: {
    readonly state: WorkbookStoredState
    readonly authoritativeBase: WorkbookLocalAuthoritativeBase
    readonly projectionOverlay: WorkbookLocalProjectionOverlay
  }): Promise<void> {
    assertWorkbookStoredState(input.state)
    this.db.transaction((db) => {
      writeWorkbookAuthoritativeBase(db, input.authoritativeBase)
      writeWorkbookProjectionOverlay(db, input.projectionOverlay)
      db.exec(
        `
          INSERT INTO runtime_state (
            id,
            workbook_name,
            snapshot_json,
            replica_json,
            authoritative_revision,
            applied_pending_local_seq,
            updated_at_ms
          )
          VALUES (1, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET
            workbook_name = excluded.workbook_name,
            snapshot_json = excluded.snapshot_json,
            replica_json = excluded.replica_json,
            authoritative_revision = excluded.authoritative_revision,
            applied_pending_local_seq = excluded.applied_pending_local_seq,
            updated_at_ms = excluded.updated_at_ms
        `,
        {
          bind: [
            extractWorkbookName(input.state.snapshot) ?? this.defaultWorkbookName,
            JSON.stringify(input.state.snapshot),
            JSON.stringify(input.state.replica),
            input.state.authoritativeRevision,
            input.state.appliedPendingLocalSeq,
            Date.now(),
          ],
        },
      )
    })
  }

  async ingestAuthoritativeDelta(input: {
    readonly state: WorkbookStoredState
    readonly authoritativeDelta: WorkbookLocalAuthoritativeDelta
    readonly projectionOverlay: WorkbookLocalProjectionOverlay
    readonly removePendingMutationIds?: readonly string[]
  }): Promise<void> {
    assertWorkbookStoredState(input.state)
    this.db.transaction((db) => {
      if ((input.removePendingMutationIds?.length ?? 0) > 0) {
        this.mutationJournal.ackPendingMutations(input.removePendingMutationIds ?? [], Date.now())
      }
      writeWorkbookAuthoritativeDelta(db, input.authoritativeDelta)
      writeWorkbookProjectionOverlay(db, input.projectionOverlay)
      db.exec(
        `
          INSERT INTO runtime_state (
            id,
            workbook_name,
            snapshot_json,
            replica_json,
            authoritative_revision,
            applied_pending_local_seq,
            updated_at_ms
          )
          VALUES (1, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET
            workbook_name = excluded.workbook_name,
            snapshot_json = excluded.snapshot_json,
            replica_json = excluded.replica_json,
            authoritative_revision = excluded.authoritative_revision,
            applied_pending_local_seq = excluded.applied_pending_local_seq,
            updated_at_ms = excluded.updated_at_ms
        `,
        {
          bind: [
            extractWorkbookName(input.state.snapshot) ?? this.defaultWorkbookName,
            JSON.stringify(input.state.snapshot),
            JSON.stringify(input.state.replica),
            input.state.authoritativeRevision,
            input.state.appliedPendingLocalSeq,
            Date.now(),
          ],
        },
      )
    })
  }

  async listPendingMutations(): Promise<WorkbookLocalMutationRecord[]> {
    return this.mutationJournal.listPendingMutations()
  }

  async listMutationJournalEntries(): Promise<WorkbookLocalMutationRecord[]> {
    return this.mutationJournal.listMutationJournalEntries()
  }

  async appendPendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void> {
    this.mutationJournal.appendPendingMutation(mutation)
  }

  async updatePendingMutation(mutation: WorkbookLocalMutationRecord): Promise<void> {
    this.mutationJournal.updatePendingMutation(mutation)
  }

  async removePendingMutation(id: string): Promise<void> {
    this.mutationJournal.removePendingMutation(id)
  }

  readViewportProjection(
    sheetName: string,
    viewport: {
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    },
  ): WorkbookLocalViewportBase | null {
    return readWorkbookViewportProjection(this.db, sheetName, viewport)
  }

  close(): void {
    if (this.closeDbOnClose) {
      this.db.close()
    }
  }
}

async function getMemorySqliteRuntime(): Promise<Sqlite3Static> {
  if (!memorySqliteRuntimePromise) {
    memorySqliteRuntimePromise = (async () => {
      try {
        return await sqlite3InitModule()
      } catch (error) {
        memorySqliteRuntimePromise = null
        throw error
      }
    })()
  }
  return await memorySqliteRuntimePromise
}

export function createOpfsWorkbookLocalStoreFactory(options: OpfsWorkbookLocalStoreFactoryOptions = {}): WorkbookLocalStoreFactory {
  const resolvedOptions: Required<OpfsWorkbookLocalStoreFactoryOptions> = {
    vfsName: options.vfsName ?? WORKBOOK_VFS_NAME,
    directory: options.directory ?? WORKBOOK_VFS_DIRECTORY,
    initialCapacity: options.initialCapacity ?? WORKBOOK_VFS_INITIAL_CAPACITY,
  }

  return {
    async open(documentId: string): Promise<WorkbookLocalStore> {
      try {
        const { poolUtil } = await getSqliteRuntime(resolvedOptions)
        const path = `/workbooks/${sanitizeDocumentId(documentId)}.sqlite`
        const db = new poolUtil.OpfsSAHPoolDb(path)
        initializeWorkbookLocalStoreSchema(db)
        return new SqliteWorkbookLocalStore(db, documentId)
      } catch (error) {
        if (error instanceof WorkbookLocalStoreLockedError || isAccessHandleConflict(error)) {
          throw toWorkbookLocalStoreLockedError(documentId, error)
        }
        if (error instanceof WorkbookLocalStoreUnavailableError || isSqliteOpenUnavailable(error)) {
          throw toWorkbookLocalStoreUnavailableError(documentId, error)
        }
        throw error
      }
    },
  }
}

export function createMemoryWorkbookLocalStoreFactory(): WorkbookLocalStoreFactory {
  const databases = new Map<string, Database>()

  return {
    async open(documentId: string): Promise<WorkbookLocalStore> {
      let db = databases.get(documentId)
      if (!db) {
        const sqlite3 = await getMemorySqliteRuntime()
        db = new sqlite3.oo1.DB(':memory:', 'c')
        initializeWorkbookLocalStoreSchema(db)
        databases.set(documentId, db)
      }
      return new SqliteWorkbookLocalStore(db, documentId, false)
    },
  }
}
