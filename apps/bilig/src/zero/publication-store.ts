import { zeroSchemaServerColumnNamesByTable, zeroSchemaTableNames } from '@bilig/zero-sync'
import type { Queryable } from './store.js'

export const DEFAULT_ZERO_PUBLICATION = 'zero_data_v2'

export const ZERO_PUBLICATION_TABLES = zeroSchemaTableNames
export const ZERO_PUBLICATION_COLUMNS_BY_TABLE = zeroSchemaServerColumnNamesByTable

const POSTGRES_IDENTIFIER_PATTERN = /^[A-Za-z_][A-Za-z0-9_]*$/

interface PublicationTableState {
  readonly tableName: string
  readonly columnNames: readonly string[] | null
}

function quoteIdentifier(identifier: string): string {
  return `"${identifier.replaceAll('"', '""')}"`
}

function formatQualifiedTable(tableName: string): string {
  return `public.${quoteIdentifier(tableName)}`
}

function formatQualifiedTableList(tableNames: readonly string[]): string {
  return tableNames.map((tableName) => formatQualifiedTable(tableName)).join(', ')
}

function normalizePublicationColumnNames(value: unknown): readonly string[] | null {
  if (!Array.isArray(value)) {
    return null
  }
  return value.every((columnName): columnName is string => typeof columnName === 'string') ? value : null
}

function parsePublicationTableRows(
  rows: readonly { tableName?: unknown; columnNames?: unknown }[],
): ReadonlyMap<string, PublicationTableState> {
  return new Map(
    rows.flatMap((row): [string, PublicationTableState][] => {
      if (typeof row.tableName !== 'string' || row.tableName.length === 0) {
        return []
      }
      return [
        [
          row.tableName,
          {
            tableName: row.tableName,
            columnNames: normalizePublicationColumnNames(row.columnNames),
          },
        ],
      ]
    }),
  )
}

function publicationTableNeedsColumnRepair(tableName: string, tableState: PublicationTableState): boolean {
  if (tableState.columnNames === null) {
    return false
  }
  const publishedColumns = new Set(tableState.columnNames)
  const expectedColumns = ZERO_PUBLICATION_COLUMNS_BY_TABLE[tableName] ?? []
  return expectedColumns.some((columnName) => !publishedColumns.has(columnName))
}

export function resolveZeroPublicationName(env: Record<string, string | undefined> = process.env): string {
  const publication = env['BILIG_ZERO_PUBLICATION']?.trim() || DEFAULT_ZERO_PUBLICATION
  if (!POSTGRES_IDENTIFIER_PATTERN.test(publication)) {
    throw new Error(`Invalid Zero publication name: ${publication}`)
  }
  return publication
}

async function publicationExists(db: Queryable, publicationName: string): Promise<boolean> {
  const result = await db.query<{ present?: unknown }>(
    `
      SELECT 1 AS present
      FROM pg_publication
      WHERE pubname = $1
      LIMIT 1
    `,
    [publicationName],
  )
  return result.rows.length > 0
}

async function loadPublicationTables(db: Queryable, publicationName: string): Promise<ReadonlyMap<string, PublicationTableState>> {
  const result = await db.query<{ tableName?: unknown; columnNames?: unknown }>(
    `
      SELECT tablename AS "tableName",
        attnames AS "columnNames"
      FROM pg_publication_tables
      WHERE pubname = $1
        AND schemaname = 'public'
    `,
    [publicationName],
  )
  return parsePublicationTableRows(result.rows)
}

export async function ensureZeroPublication(db: Queryable, publicationName = resolveZeroPublicationName()): Promise<void> {
  const quotedPublicationName = quoteIdentifier(publicationName)
  if (!(await publicationExists(db, publicationName))) {
    await db.query(`CREATE PUBLICATION ${quotedPublicationName} FOR TABLE ${formatQualifiedTableList(ZERO_PUBLICATION_TABLES)}`)
    return
  }

  const existingTables = await loadPublicationTables(db, publicationName)
  const missingTables = ZERO_PUBLICATION_TABLES.filter((tableName) => !existingTables.has(tableName))
  const columnFilteredTables = ZERO_PUBLICATION_TABLES.filter((tableName) => {
    const tableState = existingTables.get(tableName)
    return tableState ? publicationTableNeedsColumnRepair(tableName, tableState) : false
  })
  if (missingTables.length === 0 && columnFilteredTables.length === 0) {
    return
  }

  if (missingTables.length > 0) {
    await db.query(`ALTER PUBLICATION ${quotedPublicationName} ADD TABLE ${formatQualifiedTableList(missingTables)}`)
  }
  if (columnFilteredTables.length > 0) {
    await db.query(`ALTER PUBLICATION ${quotedPublicationName} DROP TABLE ${formatQualifiedTableList(columnFilteredTables)}`)
    await db.query(`ALTER PUBLICATION ${quotedPublicationName} ADD TABLE ${formatQualifiedTableList(columnFilteredTables)}`)
  }
}
