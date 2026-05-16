import type { Queryable } from './store.js'

export interface DefaultedColumn {
  readonly tableName: string
  readonly columnName: string
  readonly dataType: string
  readonly defaultSql: string
}

export interface NullableColumn {
  readonly tableName: string
  readonly columnName: string
  readonly dataType: string
}

export async function addColumnIfMissing(db: Queryable, column: NullableColumn): Promise<void> {
  await db.query(`
    ALTER TABLE ${column.tableName}
      ADD COLUMN IF NOT EXISTS ${column.columnName} ${column.dataType};
  `)
}

export async function addDefaultedColumnIfMissing(db: Queryable, column: DefaultedColumn): Promise<void> {
  await db.query(`
    ALTER TABLE ${column.tableName}
      ADD COLUMN IF NOT EXISTS ${column.columnName} ${column.dataType} DEFAULT ${column.defaultSql};
  `)
}

export async function enforceDefaultedNotNullColumn(db: Queryable, column: DefaultedColumn): Promise<void> {
  await db.query(`
    UPDATE ${column.tableName}
    SET ${column.columnName} = ${column.defaultSql}
    WHERE ${column.columnName} IS NULL;
  `)
  await db.query(`
    ALTER TABLE ${column.tableName}
      ALTER COLUMN ${column.columnName} SET DEFAULT ${column.defaultSql},
      ALTER COLUMN ${column.columnName} SET NOT NULL;
  `)
}

export async function ensureDefaultedNotNullColumn(db: Queryable, column: DefaultedColumn): Promise<void> {
  await addDefaultedColumnIfMissing(db, column)
  await enforceDefaultedNotNullColumn(db, column)
}
