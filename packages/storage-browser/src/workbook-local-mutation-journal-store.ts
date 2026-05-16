import type { Database, SqlValue } from '@sqlite.org/sqlite-wasm'

export interface WorkbookLocalMutationRecord {
  readonly id: string
  readonly localSeq: number
  readonly baseRevision: number
  readonly method: string
  readonly args: unknown[]
  readonly enqueuedAtUnixMs: number
  readonly submittedAtUnixMs: number | null
  readonly lastAttemptedAtUnixMs: number | null
  readonly ackedAtUnixMs: number | null
  readonly rebasedAtUnixMs: number | null
  readonly failedAtUnixMs: number | null
  readonly attemptCount: number
  readonly failureMessage: string | null
  readonly status: 'local' | 'submitted' | 'acked' | 'rebased' | 'failed'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isSafeNullableUnixMs(value: unknown): value is number | null {
  return value === null || isSafeNonNegativeInteger(value)
}

export function parseWorkbookLocalMutationRecord(value: unknown): WorkbookLocalMutationRecord | null {
  if (
    !isRecord(value) ||
    typeof value['id'] !== 'string' ||
    !isSafeNonNegativeInteger(value['localSeq']) ||
    !isSafeNonNegativeInteger(value['baseRevision']) ||
    typeof value['method'] !== 'string' ||
    !Array.isArray(value['args']) ||
    !isSafeNonNegativeInteger(value['enqueuedAtUnixMs']) ||
    !isSafeNullableUnixMs(value['submittedAtUnixMs']) ||
    !isSafeNullableUnixMs(value['lastAttemptedAtUnixMs']) ||
    !isSafeNullableUnixMs(value['ackedAtUnixMs']) ||
    !isSafeNullableUnixMs(value['rebasedAtUnixMs']) ||
    !isSafeNullableUnixMs(value['failedAtUnixMs']) ||
    !isSafeNonNegativeInteger(value['attemptCount']) ||
    (value['failureMessage'] !== null && typeof value['failureMessage'] !== 'string') ||
    (value['status'] !== 'local' &&
      value['status'] !== 'submitted' &&
      value['status'] !== 'acked' &&
      value['status'] !== 'rebased' &&
      value['status'] !== 'failed')
  ) {
    return null
  }
  return {
    id: value['id'],
    localSeq: value['localSeq'],
    baseRevision: value['baseRevision'],
    method: value['method'],
    args: [...value['args']],
    enqueuedAtUnixMs: value['enqueuedAtUnixMs'],
    submittedAtUnixMs: value['submittedAtUnixMs'] ?? null,
    lastAttemptedAtUnixMs: value['lastAttemptedAtUnixMs'] ?? null,
    ackedAtUnixMs: value['ackedAtUnixMs'] ?? null,
    rebasedAtUnixMs: value['rebasedAtUnixMs'] ?? null,
    failedAtUnixMs: value['failedAtUnixMs'] ?? null,
    attemptCount: value['attemptCount'],
    failureMessage: value['failureMessage'] ?? null,
    status: value['status'],
  }
}

export function assertWorkbookLocalMutationRecord(value: WorkbookLocalMutationRecord): void {
  if (!parseWorkbookLocalMutationRecord(value)) {
    throw new TypeError('Invalid workbook local mutation record')
  }
}

export class WorkbookLocalMutationJournalStore {
  constructor(private readonly db: Database) {}

  listPendingMutations(): WorkbookLocalMutationRecord[] {
    return this.readMutationRows(
      `
        SELECT op_id AS id,
               local_seq AS localSeq,
               base_revision AS baseRevision,
               method,
               args_json AS argsJson,
               enqueued_at_ms AS enqueuedAtUnixMs,
               submitted_at_ms AS submittedAtUnixMs,
               last_attempted_at_ms AS lastAttemptedAtUnixMs,
               acked_at_ms AS ackedAtUnixMs,
               rebased_at_ms AS rebasedAtUnixMs,
               failed_at_ms AS failedAtUnixMs,
               attempt_count AS attemptCount,
               failure_message AS failureMessage,
               status
          FROM pending_op
         WHERE status != 'acked'
         ORDER BY local_seq ASC
      `,
    )
  }

  listMutationJournalEntries(): WorkbookLocalMutationRecord[] {
    return this.readMutationRows(
      `
        SELECT op_id AS id,
               local_seq AS localSeq,
               base_revision AS baseRevision,
               method,
               args_json AS argsJson,
               enqueued_at_ms AS enqueuedAtUnixMs,
               submitted_at_ms AS submittedAtUnixMs,
               last_attempted_at_ms AS lastAttemptedAtUnixMs,
               acked_at_ms AS ackedAtUnixMs,
               rebased_at_ms AS rebasedAtUnixMs,
               failed_at_ms AS failedAtUnixMs,
               attempt_count AS attemptCount,
               failure_message AS failureMessage,
               status
          FROM pending_op
         ORDER BY local_seq ASC
      `,
    )
  }

  appendPendingMutation(mutation: WorkbookLocalMutationRecord): void {
    assertWorkbookLocalMutationRecord(mutation)
    this.db.transaction((db) => {
      db.exec(
        `
          INSERT INTO pending_op (
            op_id,
            local_seq,
            base_revision,
            method,
            args_json,
            enqueued_at_ms,
            submitted_at_ms,
            last_attempted_at_ms,
            acked_at_ms,
            rebased_at_ms,
            failed_at_ms,
            attempt_count,
            failure_message,
            status
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
        {
          bind: [
            mutation.id,
            mutation.localSeq,
            mutation.baseRevision,
            mutation.method,
            JSON.stringify(mutation.args),
            mutation.enqueuedAtUnixMs,
            mutation.submittedAtUnixMs,
            mutation.lastAttemptedAtUnixMs,
            mutation.ackedAtUnixMs,
            mutation.rebasedAtUnixMs,
            mutation.failedAtUnixMs,
            mutation.attemptCount,
            mutation.failureMessage,
            mutation.status,
          ],
        },
      )
    })
  }

  updatePendingMutation(mutation: WorkbookLocalMutationRecord): void {
    assertWorkbookLocalMutationRecord(mutation)
    this.db.exec(
      `
        UPDATE pending_op
           SET base_revision = ?,
               method = ?,
               args_json = ?,
               enqueued_at_ms = ?,
               submitted_at_ms = ?,
               last_attempted_at_ms = ?,
               acked_at_ms = ?,
               rebased_at_ms = ?,
               failed_at_ms = ?,
               attempt_count = ?,
               failure_message = ?,
               status = ?
         WHERE op_id = ?
      `,
      {
        bind: [
          mutation.baseRevision,
          mutation.method,
          JSON.stringify(mutation.args),
          mutation.enqueuedAtUnixMs,
          mutation.submittedAtUnixMs,
          mutation.lastAttemptedAtUnixMs,
          mutation.ackedAtUnixMs,
          mutation.rebasedAtUnixMs,
          mutation.failedAtUnixMs,
          mutation.attemptCount,
          mutation.failureMessage,
          mutation.status,
          mutation.id,
        ],
      },
    )
  }

  ackPendingMutations(ids: readonly string[], ackedAtUnixMs: number): void {
    if (ids.length === 0) {
      return
    }
    const ackPendingMutation = this.db.prepare(
      `
        UPDATE pending_op
           SET status = 'acked',
               acked_at_ms = ?,
               failed_at_ms = NULL,
               failure_message = NULL
         WHERE op_id = ?
      `,
    )
    try {
      ids.forEach((id) => {
        ackPendingMutation.bind([ackedAtUnixMs, id])
        ackPendingMutation.step()
        ackPendingMutation.reset()
      })
    } finally {
      ackPendingMutation.finalize()
    }
  }

  removePendingMutation(id: string): void {
    this.db.exec('DELETE FROM pending_op WHERE op_id = ?', {
      bind: [id],
    })
  }

  private readMutationRows(sql: string): WorkbookLocalMutationRecord[] {
    const rows: Record<string, SqlValue>[] = []
    const statement = this.db.prepare(sql)
    try {
      while (statement.step()) {
        rows.push(statement.get({}))
      }
    } finally {
      statement.finalize()
    }
    return rows.flatMap((row) => {
      const argsJson = row['argsJson']
      if (typeof argsJson !== 'string') {
        return []
      }
      try {
        const parsed = parseWorkbookLocalMutationRecord({
          ...row,
          args: JSON.parse(argsJson) as unknown,
        })
        return parsed ? [parsed] : []
      } catch {
        return []
      }
    })
  }
}
