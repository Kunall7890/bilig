import type { LiteralInput, WorkbookSnapshot, WorkbookTableSnapshot, WorkbookUnsupportedFormulaDependencySnapshot } from '@bilig/protocol'
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
  const importedFormula = translateImportedFormulaStructuredReferences({
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
