import { parseStrictBooleanEnvFlag } from './strict-env.js'

export function resolveExcelOracleDisabled(env: { BILIG_EXCEL_ORACLE_DISABLE?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_EXCEL_ORACLE_DISABLE, 'BILIG_EXCEL_ORACLE_DISABLE', false)
}
