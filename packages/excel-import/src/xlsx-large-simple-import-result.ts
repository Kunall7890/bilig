import type { CellStyleRecord, WorkbookDefinedNameSnapshot, WorkbookSnapshot, WorkbookTableSnapshot } from '@bilig/protocol'
import { formulaShouldUseCachedUnsupportedFunctionValue } from '@bilig/core/headless-runtime'
import { attachImportedRuntimeImage } from './import-runtime-image.js'
import { createWorkbookPreview } from './workbook-import-preview.js'
import { XLSX_CONTENT_TYPE } from './workbook-import-content-types.js'
import { applyImportedAutoFilterRowVisibility } from './xlsx-autofilter-row-visibility.js'
import { precisionAsDisplayedCalculationWarning, manualCalculationModeWarning } from './xlsx-calculation-settings.js'
import { canPreserveStructuredReferencesNatively } from './xlsx-import-formula-cells.js'
import { unsupportedFormulaCachesWarning } from './xlsx-import-warnings.js'
import { translateImportedFormulaStructuredReferences } from './xlsx-formula-translation.js'
import { buildLargeSimpleRuntimeSheetCells } from './xlsx-large-simple-runtime-sheet-cells.js'
import type { LargeSimpleXlsxImportResult, LargeSimpleXlsxImportStats, ParsedWorksheet } from './xlsx-large-simple-import-types.js'

type WorkbookMetadata = NonNullable<WorkbookSnapshot['workbook']['metadata']>

interface BuildLargeSimpleImportResultArgs {
  readonly fileName: string
  readonly sourceByteLength: number
  readonly workbookName: string
  readonly sheetNames: readonly string[]
  readonly sheets: WorkbookSnapshot['sheets']
  readonly previewSheets: ParsedWorksheet['preview'][]
  readonly sheetStats: ParsedWorksheet['stats'][]
  readonly warnings: string[]
  readonly definedNames: WorkbookDefinedNameSnapshot[] | undefined
  readonly importedTables: WorkbookTableSnapshot[]
  readonly importedCalculationSettings: WorkbookMetadata['calculationSettings'] | undefined
  readonly importedFormulaAudit: WorkbookMetadata['formulaAudit'] | undefined
  readonly importedWorkbookProperties: WorkbookMetadata['properties'] | undefined
  readonly importedWorkbookDocumentProperties: WorkbookMetadata['documentPropertyArtifacts'] | undefined
  readonly importedDrawingArtifacts: WorkbookMetadata['drawingArtifacts'] | undefined
  readonly importedChartArtifacts: WorkbookMetadata['chartArtifacts'] | undefined
  readonly importedChartSheetArtifacts: WorkbookMetadata['chartSheetArtifacts'] | undefined
  readonly importedCharts: WorkbookMetadata['charts'] | undefined
  readonly importedPivotArtifacts: WorkbookMetadata['pivotArtifacts'] | undefined
  readonly importedControlArtifacts: WorkbookMetadata['controlArtifacts'] | undefined
  readonly importedExternalConnections: WorkbookMetadata['externalConnections'] | undefined
  readonly importedDataModelArtifacts: WorkbookMetadata['dataModelArtifacts'] | undefined
  readonly importedExternalLinkArtifacts: WorkbookMetadata['externalLinkArtifacts'] | undefined
  readonly importedSlicerConnectionArtifacts: WorkbookMetadata['slicerConnectionArtifacts'] | undefined
  readonly importedWorkbookCellMetadata: WorkbookMetadata['cellMetadata'] | undefined
  readonly styleRecords: CellStyleRecord[]
  readonly phaseTelemetry: LargeSimpleXlsxImportStats['phaseTelemetry']
}

function translateLargeSimpleStructuredReferenceFormulas(
  sheets: WorkbookSnapshot['sheets'],
  tables: readonly WorkbookTableSnapshot[] | undefined,
): WorkbookSnapshot['sheets'] {
  if (!tables || tables.length === 0) {
    return sheets
  }
  return sheets.map((sheet) => {
    let changed = false
    const cells = sheet.cells.map((cell) => {
      if (typeof cell.formula !== 'string' || !cell.formula.includes('[')) {
        return cell
      }
      const formula = canPreserveStructuredReferencesNatively(cell.formula, tables, sheet.name, cell.address)
        ? cell.formula
        : translateImportedFormulaStructuredReferences({
            formula: cell.formula,
            ownerSheetName: sheet.name,
            ownerAddress: cell.address,
            tables,
          })
      if (formula === cell.formula) {
        return cell
      }
      changed = true
      return { ...cell, formula }
    })
    return changed ? { ...sheet, cells } : sheet
  })
}

export function buildLargeSimpleImportResult(args: BuildLargeSimpleImportResultArgs): LargeSimpleXlsxImportResult {
  const sortedImportedTables =
    args.importedTables.length > 0 ? [...args.importedTables].toSorted((left, right) => left.name.localeCompare(right.name)) : undefined
  const importedSheets = translateLargeSimpleStructuredReferenceFormulas(
    args.sheets.map((sheet) => applyImportedAutoFilterRowVisibility(sheet, sortedImportedTables)),
    sortedImportedTables,
  )
  const hasFormulaCells = args.sheetStats.some((entry) => entry.formulaCellCount > 0)
  if (args.importedCalculationSettings?.mode === 'manual') {
    args.warnings.push(manualCalculationModeWarning)
  }
  if (args.importedCalculationSettings?.fullPrecision === false && hasFormulaCells) {
    args.warnings.push(precisionAsDisplayedCalculationWarning)
  }
  const definedFormulaNames = new Set((args.definedNames ?? []).map((definedName) => definedName.name.trim().toUpperCase()))
  if (
    importedSheets.some((sheet) =>
      sheet.cells.some(
        (cell) =>
          cell.value !== undefined &&
          typeof cell.formula === 'string' &&
          formulaShouldUseCachedUnsupportedFunctionValue(cell.formula, definedFormulaNames),
      ),
    )
  ) {
    args.warnings.push(unsupportedFormulaCachesWarning)
  }
  const calculationSettings =
    args.importedCalculationSettings ??
    (hasFormulaCells
      ? {
          mode: 'automatic' as const,
          compatibilityMode: 'excel-modern' as const,
          fullCalcOnLoad: false,
          forceFullCalc: false,
        }
      : undefined)
  const workbookMetadata =
    args.definedNames ||
    args.importedWorkbookProperties ||
    args.importedWorkbookDocumentProperties ||
    args.importedDrawingArtifacts ||
    args.importedChartArtifacts ||
    args.importedChartSheetArtifacts ||
    args.importedCharts ||
    args.importedPivotArtifacts ||
    args.importedControlArtifacts ||
    sortedImportedTables ||
    args.styleRecords.length > 0 ||
    args.importedExternalConnections ||
    args.importedDataModelArtifacts ||
    args.importedExternalLinkArtifacts ||
    args.importedSlicerConnectionArtifacts ||
    args.importedWorkbookCellMetadata ||
    calculationSettings ||
    args.importedFormulaAudit
      ? {
          ...(args.importedWorkbookProperties ? { properties: args.importedWorkbookProperties } : {}),
          ...(args.importedWorkbookDocumentProperties ? { documentPropertyArtifacts: args.importedWorkbookDocumentProperties } : {}),
          ...(args.definedNames ? { definedNames: args.definedNames } : {}),
          ...(args.importedDrawingArtifacts ? { drawingArtifacts: args.importedDrawingArtifacts } : {}),
          ...(args.importedChartArtifacts ? { chartArtifacts: args.importedChartArtifacts } : {}),
          ...(args.importedChartSheetArtifacts ? { chartSheetArtifacts: args.importedChartSheetArtifacts } : {}),
          ...(args.importedCharts ? { charts: args.importedCharts } : {}),
          ...(args.importedPivotArtifacts ? { pivotArtifacts: args.importedPivotArtifacts } : {}),
          ...(args.importedControlArtifacts ? { controlArtifacts: args.importedControlArtifacts } : {}),
          ...(sortedImportedTables ? { tables: sortedImportedTables } : {}),
          ...(args.styleRecords.length > 0 ? { styles: args.styleRecords } : {}),
          ...(args.importedExternalConnections ? { externalConnections: args.importedExternalConnections } : {}),
          ...(args.importedDataModelArtifacts ? { dataModelArtifacts: args.importedDataModelArtifacts } : {}),
          ...(args.importedExternalLinkArtifacts ? { externalLinkArtifacts: args.importedExternalLinkArtifacts } : {}),
          ...(args.importedSlicerConnectionArtifacts ? { slicerConnectionArtifacts: args.importedSlicerConnectionArtifacts } : {}),
          ...(args.importedWorkbookCellMetadata ? { cellMetadata: args.importedWorkbookCellMetadata } : {}),
          ...(calculationSettings ? { calculationSettings } : {}),
          ...(args.importedFormulaAudit ? { formulaAudit: args.importedFormulaAudit } : {}),
        }
      : undefined
  const runtimeSheetCells = buildLargeSimpleRuntimeSheetCells(args.sheetStats, importedSheets)
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: args.workbookName,
      ...(workbookMetadata ? { metadata: workbookMetadata } : {}),
    },
    sheets: importedSheets,
  }
  const stats: LargeSimpleXlsxImportStats = {
    sheetCount: args.sheets.length,
    cellCount: args.sheetStats.reduce((sum, entry) => sum + entry.cellCount, 0),
    formulaCellCount: args.sheetStats.reduce((sum, entry) => sum + entry.formulaCellCount, 0),
    valueCellCount: args.sheetStats.reduce((sum, entry) => sum + entry.valueCellCount, 0),
    definedNameCount: args.definedNames?.length ?? 0,
    tableCount: sortedImportedTables?.length ?? args.sheetStats.reduce((sum, entry) => sum + entry.tableCount, 0),
    mergeCount: args.sheetStats.reduce((sum, entry) => sum + entry.mergeCount, 0),
    conditionalFormatCount: args.sheetStats.reduce((sum, entry) => sum + entry.conditionalFormatCount, 0),
    dataValidationCount: args.sheetStats.reduce((sum, entry) => sum + entry.dataValidationCount, 0),
    warningCount: args.warnings.length,
    dimensions: args.sheetStats.map((entry) => entry.dimension),
    phaseTelemetry: args.phaseTelemetry,
  }

  return {
    snapshot:
      runtimeSheetCells.length > 0
        ? attachImportedRuntimeImage(snapshot, {
            version: 1,
            templateBank: [],
            formulaInstances: [],
            formulaValues: [],
            sheetCells: runtimeSheetCells,
          })
        : snapshot,
    workbookName: args.workbookName,
    sheetNames: [...args.sheetNames],
    warnings: args.warnings,
    preview: createWorkbookPreview({
      contentType: XLSX_CONTENT_TYPE,
      fileName: args.fileName,
      fileSizeBytes: args.sourceByteLength,
      workbookName: args.workbookName,
      sheets: args.previewSheets,
      warnings: args.warnings,
    }),
    stats,
  }
}
