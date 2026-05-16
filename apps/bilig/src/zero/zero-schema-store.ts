import type { Queryable } from './store.js'
import { ensureDefaultedNotNullColumn } from './schema-upgrade.js'

export async function ensureZeroSyncSchema(db: Queryable): Promise<void> {
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbooks (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      snapshot JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'owner_user_id',
    dataType: 'TEXT',
    defaultSql: "'system'",
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'head_revision',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'calculated_revision',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'source_projection_version',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'calc_mode',
    dataType: 'TEXT',
    defaultSql: "'automatic'",
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'compatibility_mode',
    dataType: 'TEXT',
    defaultSql: "'excel-modern'",
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'recalc_epoch',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await db.query(`ALTER TABLE workbooks ADD COLUMN IF NOT EXISTS replica_snapshot JSONB;`)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'workbooks',
    columnName: 'created_at',
    dataType: 'TIMESTAMPTZ',
    defaultSql: 'NOW()',
  })

  await db.query(`
    CREATE TABLE IF NOT EXISTS sheets (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_id INTEGER,
      name TEXT NOT NULL,
      sort_order INTEGER NOT NULL,
      PRIMARY KEY (workbook_id, name)
    );
  `)
  await db.query(`ALTER TABLE sheets ADD COLUMN IF NOT EXISTS sheet_id INTEGER;`)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'sheets',
    columnName: 'freeze_rows',
    dataType: 'INTEGER',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'sheets',
    columnName: 'freeze_cols',
    dataType: 'INTEGER',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'sheets',
    columnName: 'created_at',
    dataType: 'TIMESTAMPTZ',
    defaultSql: 'NOW()',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'sheets',
    columnName: 'updated_at',
    dataType: 'TIMESTAMPTZ',
    defaultSql: 'NOW()',
  })

  await db.query(`
    CREATE TABLE IF NOT EXISTS cells (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      input_value JSONB,
      formula TEXT,
      format TEXT,
      PRIMARY KEY (workbook_id, sheet_name, address)
    );
  `)
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS row_num INTEGER;`)
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS col_num INTEGER;`)
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS style_id TEXT;`)
  await db.query(`ALTER TABLE cells ADD COLUMN IF NOT EXISTS explicit_format_id TEXT;`)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'cells',
    columnName: 'source_revision',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'cells',
    columnName: 'updated_by',
    dataType: 'TEXT',
    defaultSql: "'system'",
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'cells',
    columnName: 'updated_at',
    dataType: 'TIMESTAMPTZ',
    defaultSql: 'NOW()',
  })

  await db.query(`
    CREATE TABLE IF NOT EXISTS cell_eval (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      address TEXT NOT NULL,
      value JSONB NOT NULL,
      flags INTEGER NOT NULL,
      version INTEGER NOT NULL,
      PRIMARY KEY (workbook_id, sheet_name, address)
    );
  `)
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS row_num INTEGER;`)
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS col_num INTEGER;`)
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS style_id TEXT;`)
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS style_json JSONB;`)
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS format_id TEXT;`)
  await db.query(`ALTER TABLE cell_eval ADD COLUMN IF NOT EXISTS format_code TEXT;`)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'cell_eval',
    columnName: 'calc_revision',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'cell_eval',
    columnName: 'updated_at',
    dataType: 'TIMESTAMPTZ',
    defaultSql: 'NOW()',
  })

  await db.query(`
    CREATE TABLE IF NOT EXISTS row_metadata (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      start_index INTEGER NOT NULL,
      count INTEGER NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (workbook_id, sheet_name, start_index)
    );
  `)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'row_metadata',
    columnName: 'source_revision',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'row_metadata',
    columnName: 'updated_at',
    dataType: 'TIMESTAMPTZ',
    defaultSql: 'NOW()',
  })

  await db.query(`
    CREATE TABLE IF NOT EXISTS column_metadata (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      sheet_name TEXT NOT NULL,
      start_index INTEGER NOT NULL,
      count INTEGER NOT NULL,
      size INTEGER,
      hidden BOOLEAN,
      PRIMARY KEY (workbook_id, sheet_name, start_index)
    );
  `)
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'column_metadata',
    columnName: 'source_revision',
    dataType: 'BIGINT',
    defaultSql: '0',
  })
  await ensureDefaultedNotNullColumn(db, {
    tableName: 'column_metadata',
    columnName: 'updated_at',
    dataType: 'TIMESTAMPTZ',
    defaultSql: 'NOW()',
  })

  await db.query(`
    CREATE TABLE IF NOT EXISTS defined_names (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      name TEXT NOT NULL,
      value JSONB NOT NULL,
      PRIMARY KEY (workbook_id, name)
    );
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_metadata (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      key TEXT NOT NULL,
      value JSONB NOT NULL,
      PRIMARY KEY (workbook_id, key)
    );
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS calculation_settings (
      workbook_id TEXT PRIMARY KEY REFERENCES workbooks(id) ON DELETE CASCADE,
      mode TEXT NOT NULL,
      recalc_epoch BIGINT NOT NULL DEFAULT 0
    );
  `)

  await db.query(`
    CREATE TABLE IF NOT EXISTS cell_styles (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      style_id TEXT NOT NULL,
      record_json JSONB NOT NULL,
      hash TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, style_id)
    );
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS cell_number_formats (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      format_id TEXT NOT NULL,
      code TEXT NOT NULL,
      kind TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, format_id)
    );
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_event (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      revision BIGINT NOT NULL,
      actor_user_id TEXT NOT NULL,
      client_mutation_id TEXT,
      txn_json JSONB NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, revision)
    );
  `)
  await db.query(`
    CREATE TABLE IF NOT EXISTS recalc_job (
      id TEXT PRIMARY KEY,
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      from_revision BIGINT NOT NULL,
      to_revision BIGINT NOT NULL,
      dirty_regions_json JSONB,
      status TEXT NOT NULL,
      attempts INTEGER NOT NULL DEFAULT 0,
      lease_until TIMESTAMPTZ,
      lease_owner TEXT,
      last_error TEXT,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `)

  await db.query(`
    CREATE TABLE IF NOT EXISTS workbook_snapshot (
      workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
      revision BIGINT NOT NULL,
      format TEXT NOT NULL,
      payload JSONB NOT NULL,
      replica_snapshot JSONB,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      PRIMARY KEY (workbook_id, revision)
    );
  `)

  await db.query(`CREATE INDEX IF NOT EXISTS sheets_workbook_sort_order_idx ON sheets(workbook_id, sort_order);`)
  await db.query(`CREATE UNIQUE INDEX IF NOT EXISTS sheets_workbook_sheet_id_idx ON sheets(workbook_id, sheet_id);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cells_workbook_sheet_idx ON cells(workbook_id, sheet_name);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cells_workbook_sheet_row_col_idx ON cells(workbook_id, sheet_name, row_num, col_num);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cell_eval_workbook_sheet_idx ON cell_eval(workbook_id, sheet_name);`)
  await db.query(`CREATE INDEX IF NOT EXISTS cell_eval_workbook_sheet_row_col_idx ON cell_eval(workbook_id, sheet_name, row_num, col_num);`)
  await db.query(`CREATE INDEX IF NOT EXISTS row_metadata_workbook_sheet_idx ON row_metadata(workbook_id, sheet_name, start_index);`)
  await db.query(`CREATE INDEX IF NOT EXISTS column_metadata_workbook_sheet_idx ON column_metadata(workbook_id, sheet_name, start_index);`)
  await db.query(`CREATE INDEX IF NOT EXISTS recalc_job_status_lease_created_idx ON recalc_job(status, lease_until, created_at);`)
  await db.query(`CREATE INDEX IF NOT EXISTS workbook_event_workbook_created_idx ON workbook_event(workbook_id, created_at);`)
  await db.query(`CREATE INDEX IF NOT EXISTS workbook_snapshot_workbook_revision_idx ON workbook_snapshot(workbook_id, revision DESC);`)

  await db.query(`
    DO $$
    BEGIN
      IF to_regclass('public.computed_cells') IS NOT NULL THEN
        INSERT INTO cell_eval (
          workbook_id,
          sheet_name,
          address,
          row_num,
          col_num,
          value,
          flags,
          version,
          calc_revision,
          updated_at
        )
        SELECT
          workbook_id,
          sheet_name,
          address,
          row_num,
          col_num,
          value,
          flags,
          version,
          calc_revision,
          updated_at
        FROM computed_cells
        ON CONFLICT (workbook_id, sheet_name, address)
        DO NOTHING;
      END IF;
    END $$;
  `)
}
