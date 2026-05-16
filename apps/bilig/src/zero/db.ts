import { zeroNodePg, type NodePgTransaction } from '@rocicorp/zero/server/adapters/pg'
import { Pool } from 'pg'
import { schema } from '@bilig/zero-sync'
import { logError } from '../runtime-logger.js'
import type { WorkbookRuntimeStoreConnection } from './store.js'

export function resolveZeroDatabaseUrl(): string | null {
  return process.env['ZERO_UPSTREAM_DB'] ?? process.env['DATABASE_URL'] ?? process.env['BILIG_DATABASE_URL'] ?? null
}

export function createZeroPool(connectionString: string): Pool {
  const pool = new Pool({
    connectionString,
  })
  pool.on('error', (error) => {
    logError('Zero Postgres pool error', error)
  })
  return pool
}

export function createZeroDbProvider(connection: NodePgTransaction | string) {
  return zeroNodePg(schema, connection)
}

export type BiligDbProvider = ReturnType<typeof createZeroDbProvider>

export function createWorkbookRuntimeStoreConnection(
  connection: NodePgTransaction,
  dbProvider: BiligDbProvider,
): WorkbookRuntimeStoreConnection {
  return {
    query: (text, values) => connection.query(text, values),
    run: (query, options) => dbProvider.run(query, options),
  }
}

declare module '@rocicorp/zero' {
  interface DefaultTypes {
    dbProvider: BiligDbProvider
  }
}
