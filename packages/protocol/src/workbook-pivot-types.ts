import type { WorkbookPackageRelationshipSnapshot } from './package-artifacts.js'
import type { CellRangeRef, LiteralInput } from './types.js'

export type PivotAggregation = 'sum' | 'count' | 'countNums' | 'average' | 'min' | 'max' | 'product'

export interface WorkbookPivotValueSnapshot {
  sourceColumn: string
  summarizeBy: PivotAggregation
  outputLabel?: string
}

export interface WorkbookPivotFilterSnapshot {
  sourceColumn: string
  includedValues?: LiteralInput[]
  hiddenValues?: LiteralInput[]
}

export interface WorkbookPivotPageFieldSnapshot {
  sourceColumn: string
  selectedValue?: LiteralInput
}

export interface WorkbookPivotHiddenItemsSnapshot {
  sourceColumn: string
  values: LiteralInput[]
}

export interface WorkbookPivotCalculatedFormulaSnapshot {
  name: string
  formula: string
  clause: '18.10' | '3.2.3.1'
}

export interface WorkbookPivotSnapshot {
  name: string
  sheetName: string
  address: string
  source?: CellRangeRef
  cacheId?: number
  sourceKind?: 'worksheet' | 'table' | 'named-range' | 'external-cache-only'
  groupBy: string[]
  columnFields?: string[]
  pageFields?: WorkbookPivotPageFieldSnapshot[]
  filters?: WorkbookPivotFilterSnapshot[]
  hiddenItems?: WorkbookPivotHiddenItemsSnapshot[]
  calculatedFields?: WorkbookPivotCalculatedFormulaSnapshot[]
  calculatedItems?: WorkbookPivotCalculatedFormulaSnapshot[]
  cacheOnly?: boolean
  cacheFields?: string[]
  cachedRecords?: LiteralInput[][]
  values: WorkbookPivotValueSnapshot[]
  rows: number
  cols: number
}

export interface WorkbookUnsupportedFormulaDependencySnapshot {
  kind: 'external-workbook-reference'
  sheetName: string
  address: string
  formula: string
  importedFormula: string
  linkedWorkbooks: WorkbookExternalWorkbookReferenceSnapshot[]
  cachedValuesUsed: boolean
  cachedFormulaValuePreserved: boolean
  cachedExternalReferenceValuesUsed: boolean
  resolvedExternalReferenceCount: number
  unresolvedExternalReferenceCount: number
  reason: string
}

export interface WorkbookExternalWorkbookReferenceSnapshot {
  bookIndex: number
  packagePath?: string
  target?: string
  targetMode?: string
  workbookName?: string
  sheetNames?: string[]
}

export interface WorkbookUnsupportedPivotSnapshot {
  kind: 'external-cache' | 'raw-part'
  reason: string
  cacheId?: number
  sourceType?: string
  cachedRecordCount?: number
  cacheFieldNames?: string[]
  cacheOnly?: boolean
  sheetName?: string
  address?: string
  name?: string
  packagePart?: string
}

export interface WorkbookPivotPackagePartSnapshot {
  path: string
  xml: string
}

export interface WorkbookPivotArtifactsSnapshot {
  parts: WorkbookPivotPackagePartSnapshot[]
  workbookPivotCachesXml?: string
  workbookRelationships?: WorkbookPackageRelationshipSnapshot[]
}

export interface WorkbookSheetPivotArtifactsSnapshot {
  relationships: WorkbookPackageRelationshipSnapshot[]
  pivotTableDefinitionsXml?: string
}
