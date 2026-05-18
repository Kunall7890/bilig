type WorkbookFormulaAuditLiteralInput = number | string | boolean | null

export type WorkbookFormulaAuditContext =
  | 'worksheet-cell'
  | 'defined-name'
  | 'conditional-format'
  | 'data-validation'
  | 'chart'
  | 'pivot-calculated'
  | 'external-name'
  | 'ole-link'

export type WorkbookFormulaCacheStatus = 'trustedCached' | 'staleRisk' | 'externalSubstitution' | 'engineRecomputed' | 'missing'

export interface WorkbookFormulaAuditAttributesSnapshot {
  aca?: boolean
  bx?: boolean
  ca?: boolean
  xmlSpace?: string
}

export interface WorkbookFormulaAuditEntrySnapshot {
  context: WorkbookFormulaAuditContext
  clause: string
  formula: string
  normalizedFormula?: string
  sheetName?: string
  address?: string
  name?: string
  sqref?: string
  formulaType?: string
  sharedIndex?: string
  ref?: string
  cellValueType?: string
  cachedValue?: WorkbookFormulaAuditLiteralInput
  cachedValueRaw?: string
  cacheStatus?: WorkbookFormulaCacheStatus
  rawFormulaXml?: string
  attributes?: WorkbookFormulaAuditAttributesSnapshot
}

export interface WorkbookFormulaDiagnosticSnapshot {
  code: string
  context: WorkbookFormulaAuditContext | 'calculation'
  clause: string
  message: string
  sheetName?: string
  address?: string
  name?: string
  formula?: string
}

export interface WorkbookCalcChainCellSnapshot {
  sheetIndex: number
  sheetName?: string
  address: string
  childChain?: boolean
  newDependencyLevel?: boolean
}

export interface WorkbookCalcChainSnapshot {
  packagePath: string
  cells: WorkbookCalcChainCellSnapshot[]
}

export interface WorkbookFormulaAuditSnapshot {
  formulas: WorkbookFormulaAuditEntrySnapshot[]
  diagnostics?: WorkbookFormulaDiagnosticSnapshot[]
  calcChain?: WorkbookCalcChainSnapshot
}
