import type { LiteralInput, WorkbookSnapshot, WorkbookTableSnapshot, WorkbookUnsupportedFormulaDependencySnapshot } from '@bilig/protocol'
import { parseFormula, type FormulaNode } from '@bilig/formula'
import {
  collectImportedFormulaExternalWorkbookReferences,
  translateImportedFormulaExternalReferences,
  type ImportedExternalCacheSheetMap,
  type ImportedExternalLinkCaches,
  type ImportedExternalWorkbookReferences,
} from './xlsx-external-references.js'
import { normalizeImportedFormulaSource, translateImportedFormulaStructuredReferences } from './xlsx-formula-translation.js'
import { formulaReferencesExternalWorkbook, formulaReferencesVolatileFunction } from './xlsx-import-warnings.js'
import type { WorksheetFormulaCell } from './xlsx-formulas.js'
import { decodeA1CellRef } from './xlsx-a1-utils.js'

type ImportedFormulaCellSnapshot = WorkbookSnapshot['sheets'][number]['cells'][number] & {
  readonly formula: string
}

export interface ImportedFormulaSnapshotCellResult {
  readonly formulaCell: ImportedFormulaCellSnapshot
  readonly hasExternalWorkbookDependency: boolean
  readonly hasVolatileFormula: boolean
  readonly hasCachedLiteral: boolean
  readonly materializedExternalCacheSheetKeys: readonly string[]
  readonly unsupportedFormulaDependency?: WorkbookUnsupportedFormulaDependencySnapshot
}

export interface BuildImportedFormulaSnapshotCellArgs {
  readonly sheetName: string
  readonly address: string
  readonly formula: string
  readonly formulaManifest: WorksheetFormulaCell | undefined
  readonly cachedLiteral: LiteralInput | undefined
  readonly tables: readonly WorkbookTableSnapshot[] | undefined
  readonly externalLinkCaches: ImportedExternalLinkCaches
  readonly externalCacheSheetNames?: ImportedExternalCacheSheetMap
  readonly externalWorkbookReferences: ImportedExternalWorkbookReferences
}

export function buildImportedFormulaSnapshotCell(
  args: BuildImportedFormulaSnapshotCellArgs,
): ImportedFormulaSnapshotCellResult | undefined {
  if (args.formula.trim().length === 0) {
    return undefined
  }

  const normalizedFormula = normalizeImportedFormulaSource(args.formula)
  const externalReferenceTranslation = translateImportedFormulaExternalReferences(
    normalizedFormula,
    args.externalLinkCaches,
    args.externalCacheSheetNames,
  )
  const hasExternalWorkbookDependency =
    externalReferenceTranslation.resolvedCount > 0 ||
    externalReferenceTranslation.unresolvedCount > 0 ||
    formulaReferencesExternalWorkbook(externalReferenceTranslation.formula)
  const importedFormula = canPreserveStructuredReferencesNatively(
    externalReferenceTranslation.formula,
    args.tables,
    args.sheetName,
    args.address,
  )
    ? externalReferenceTranslation.formula
    : translateImportedFormulaStructuredReferences({
        formula: externalReferenceTranslation.formula,
        ownerSheetName: args.sheetName,
        ownerAddress: args.address,
        tables: args.tables,
      })
  const cachedLiteral = args.formulaManifest?.cachedValue !== undefined ? args.formulaManifest.cachedValue : args.cachedLiteral
  const formulaCell: ImportedFormulaCellSnapshot = {
    address: args.address,
    formula: importedFormula,
    ...(cachedLiteral !== undefined ? { value: cachedLiteral } : {}),
  }

  const unsupportedFormulaDependency: WorkbookUnsupportedFormulaDependencySnapshot | undefined = hasExternalWorkbookDependency
    ? {
        kind: 'external-workbook-reference',
        sheetName: args.sheetName,
        address: args.address,
        formula: normalizedFormula,
        importedFormula,
        linkedWorkbooks: [...collectImportedFormulaExternalWorkbookReferences(normalizedFormula, args.externalWorkbookReferences)],
        cachedValuesUsed: cachedLiteral !== undefined || externalReferenceTranslation.resolvedCount > 0,
        cachedFormulaValuePreserved: cachedLiteral !== undefined,
        cachedExternalReferenceValuesUsed: externalReferenceTranslation.resolvedCount > 0,
        resolvedExternalReferenceCount: externalReferenceTranslation.resolvedCount,
        unresolvedExternalReferenceCount: externalReferenceTranslation.unresolvedCount,
        reason:
          'Formula depends on an external workbook reference; cached linked values are preserved but linked workbooks are not recalculated during import.',
      }
    : undefined

  return {
    formulaCell,
    hasCachedLiteral: cachedLiteral !== undefined,
    hasExternalWorkbookDependency,
    materializedExternalCacheSheetKeys: externalReferenceTranslation.materializedExternalCacheSheetKeys,
    hasVolatileFormula: formulaReferencesVolatileFunction(normalizedFormula),
    ...(unsupportedFormulaDependency ? { unsupportedFormulaDependency } : {}),
  }
}

function canPreserveStructuredReferencesNatively(
  formula: string,
  tables: readonly WorkbookTableSnapshot[] | undefined,
  ownerSheetName: string,
  ownerAddress: string,
): boolean {
  if (!formula.includes('[') || !tables || tables.length === 0) {
    return false
  }
  try {
    const ast = parseFormula(formula)
    return formulaHasStructuredReference(ast) && formulaStructuredReferencesAreNativeSupported(ast, tables, ownerSheetName, ownerAddress)
  } catch {
    return false
  }
}

function formulaStructuredReferencesAreNativeSupported(
  node: FormulaNode,
  tables: readonly WorkbookTableSnapshot[],
  ownerSheetName: string,
  ownerAddress: string,
): boolean {
  switch (node.kind) {
    case 'StructuredRef':
      return structuredReferenceIsNativeSupported(node, tables, ownerSheetName, ownerAddress)
    case 'ArrayConstant':
      return node.rows.every((row) =>
        row.every((entry) => formulaStructuredReferencesAreNativeSupported(entry, tables, ownerSheetName, ownerAddress)),
      )
    case 'UnaryExpr':
      return formulaStructuredReferencesAreNativeSupported(node.argument, tables, ownerSheetName, ownerAddress)
    case 'BinaryExpr':
      return (
        formulaStructuredReferencesAreNativeSupported(node.left, tables, ownerSheetName, ownerAddress) &&
        formulaStructuredReferencesAreNativeSupported(node.right, tables, ownerSheetName, ownerAddress)
      )
    case 'CallExpr':
      return node.args.every((arg) => formulaStructuredReferencesAreNativeSupported(arg, tables, ownerSheetName, ownerAddress))
    case 'InvokeExpr':
      return (
        formulaStructuredReferencesAreNativeSupported(node.callee, tables, ownerSheetName, ownerAddress) &&
        node.args.every((arg) => formulaStructuredReferencesAreNativeSupported(arg, tables, ownerSheetName, ownerAddress))
      )
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return true
  }
}

function structuredReferenceIsNativeSupported(
  node: Extract<FormulaNode, { kind: 'StructuredRef' }>,
  tables: readonly WorkbookTableSnapshot[],
  ownerSheetName: string,
  ownerAddress: string,
): boolean {
  const table =
    node.tableName.length > 0 ? findTableByName(tables, node.tableName) : ownerTableForAddress(tables, ownerSheetName, ownerAddress)
  if (!table) {
    return false
  }
  if (findTableColumnIndex(table, node.columnName) === -1) {
    return false
  }
  if (node.endColumnName !== undefined && findTableColumnIndex(table, node.endColumnName) === -1) {
    return false
  }
  if (node.section === 'headers') {
    return table.headerRow
  }
  if (node.section === 'totals') {
    return table.totalsRow
  }
  return true
}

function formulaHasStructuredReference(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'StructuredRef':
      return true
    case 'ArrayConstant':
      return node.rows.some((row) => row.some(formulaHasStructuredReference))
    case 'UnaryExpr':
      return formulaHasStructuredReference(node.argument)
    case 'BinaryExpr':
      return formulaHasStructuredReference(node.left) || formulaHasStructuredReference(node.right)
    case 'CallExpr':
      return node.args.some(formulaHasStructuredReference)
    case 'InvokeExpr':
      return formulaHasStructuredReference(node.callee) || node.args.some(formulaHasStructuredReference)
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return false
  }
}

function findTableByName(tables: readonly WorkbookTableSnapshot[], tableName: string): WorkbookTableSnapshot | undefined {
  return tables.find((table) => table.name.localeCompare(tableName, undefined, { sensitivity: 'accent' }) === 0)
}

function findTableColumnIndex(table: WorkbookTableSnapshot, columnName: string): number {
  const normalizedColumnName = normalizeStructuredColumnLookupName(columnName)
  return table.columnNames.findIndex((candidate) => normalizeStructuredColumnLookupName(candidate) === normalizedColumnName)
}

function normalizeStructuredColumnLookupName(columnName: string): string {
  return columnName.replace(/\s+/gu, ' ').trim().toLocaleLowerCase('en-US')
}

function ownerTableForAddress(
  tables: readonly WorkbookTableSnapshot[],
  ownerSheetName: string,
  ownerAddress: string,
): WorkbookTableSnapshot | undefined {
  const owner = decodeA1CellRef(ownerAddress)
  return tables.find((table) => {
    if (table.sheetName !== ownerSheetName) {
      return false
    }
    const start = decodeA1CellRef(table.startAddress)
    const end = decodeA1CellRef(table.endAddress)
    return owner.r >= start.r && owner.r <= end.r && owner.c >= start.c && owner.c <= end.c
  })
}
