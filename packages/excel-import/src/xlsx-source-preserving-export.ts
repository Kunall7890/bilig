import type { LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import {
  exportXlsxSourceLiteralPatches as exportBiligXlsxSourceLiteralPatches,
  exportXlsxSourceLiteralPatchesToFile as exportBiligXlsxSourceLiteralPatchesToFile,
  exportXlsxSourceLiteralPatchesToFileAsync as exportBiligXlsxSourceLiteralPatchesToFileAsync,
  type XlsxSourceLiteralPatchExportInput as BiligXlsxSourceLiteralPatchExportInput,
  type XlsxSourceTextPatch,
} from '@bilig/xlsx'

import { applyExportCalculationSettingsToWorkbookXml } from './xlsx-calculation-settings.js'
import { readImportedXlsxSourceCellPatches, type ImportedXlsxSourceCellPatch, type ImportedXlsxSourceReader } from './xlsx-source-bytes.js'

type ImportedXlsxSourceReference = Uint8Array | ImportedXlsxSourceReader

const sourcePreservingOutputCalculationSettings = Symbol.for('bilig.sourcePreservingXlsxOutputCalculationSettings')

type SnapshotWithSourcePreservingOutputCalculationSettings = WorkbookSnapshot & {
  readonly [sourcePreservingOutputCalculationSettings]?: NonNullable<WorkbookSnapshot['workbook']['metadata']>['calculationSettings']
}

export interface XlsxSourceLiteralPatch {
  readonly sheetName: string
  readonly address: string
  readonly value: LiteralInput
  readonly preserveFormula?: boolean
}

export interface XlsxSourceLiteralPatchExportInput {
  readonly source: ImportedXlsxSourceReference
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly sheetNames?: readonly string[]
  readonly workbookName?: string
}

export interface XlsxSourceLiteralPatchFileExportInput extends XlsxSourceLiteralPatchExportInput {
  readonly outputPath: string
}

export interface XlsxSourceLiteralPatchFileExportResult {
  readonly bytesWritten: number
}

export function tryExportSourcePreservingXlsx(snapshot: WorkbookSnapshot, source: ImportedXlsxSourceReference): Uint8Array | null {
  const input = sourcePreservingPatchInputFromSnapshot(snapshot, source)
  if (!input) {
    return null
  }
  try {
    return exportBiligXlsxSourceLiteralPatches(input)
  } catch {
    return null
  }
}

export function tryExportSourcePreservingXlsxToFile(
  snapshot: WorkbookSnapshot,
  source: ImportedXlsxSourceReference,
  outputPath: string,
): XlsxSourceLiteralPatchFileExportResult | null {
  const input = sourcePreservingPatchInputFromSnapshot(snapshot, source)
  if (!input) {
    return null
  }
  try {
    return exportBiligXlsxSourceLiteralPatchesToFile({ ...input, outputPath })
  } catch {
    return null
  }
}

export async function tryExportSourcePreservingXlsxToFileAsync(
  snapshot: WorkbookSnapshot,
  source: ImportedXlsxSourceReference,
  outputPath: string,
): Promise<XlsxSourceLiteralPatchFileExportResult | null> {
  const input = sourcePreservingPatchInputFromSnapshot(snapshot, source)
  if (!input) {
    return null
  }
  try {
    return await exportBiligXlsxSourceLiteralPatchesToFileAsync({ ...input, outputPath })
  } catch {
    return null
  }
}

export function exportXlsxSourceLiteralPatches(input: XlsxSourceLiteralPatchExportInput): Uint8Array {
  return exportBiligXlsxSourceLiteralPatches(input)
}

export function exportXlsxSourceLiteralPatchesToFile(input: XlsxSourceLiteralPatchFileExportInput): XlsxSourceLiteralPatchFileExportResult {
  return exportBiligXlsxSourceLiteralPatchesToFile(input)
}

export async function exportXlsxSourceLiteralPatchesToFileAsync(
  input: XlsxSourceLiteralPatchFileExportInput,
): Promise<XlsxSourceLiteralPatchFileExportResult> {
  return exportBiligXlsxSourceLiteralPatchesToFileAsync(input)
}

function sourcePreservingPatchInputFromSnapshot(
  snapshot: WorkbookSnapshot,
  source: ImportedXlsxSourceReference,
): BiligXlsxSourceLiteralPatchExportInput | null {
  const patches = readImportedXlsxSourceCellPatches(snapshot)
  if (patches.length === 0) {
    return null
  }
  const calculationTextPatches = sourcePreservingCalculationTextPatches(snapshot)
  return {
    source,
    patches: patches.map(importedPatchToLiteralPatch),
    sheetNames: orderedSheetNames(snapshot),
    workbookName: snapshot.workbook.name,
    textPatches: calculationTextPatches,
    forceWorkbookRecalculation: calculationTextPatches.length === 0,
  }
}

function orderedSheetNames(snapshot: WorkbookSnapshot): string[] {
  return snapshot.sheets.toSorted((left, right) => left.order - right.order).map((sheet) => sheet.name)
}

function importedPatchToLiteralPatch(patch: ImportedXlsxSourceCellPatch): XlsxSourceLiteralPatch {
  return {
    sheetName: patch.sheetName,
    address: patch.address,
    value: patch.value,
    ...(patch.preserveFormula === true ? { preserveFormula: true } : {}),
  }
}

function sourcePreservingCalculationTextPatches(snapshot: WorkbookSnapshot): readonly XlsxSourceTextPatch[] {
  const settings = (snapshot as SnapshotWithSourcePreservingOutputCalculationSettings)[sourcePreservingOutputCalculationSettings]
  return settings
    ? [
        {
          path: 'xl/workbook.xml',
          patchText: (text) => applyExportCalculationSettingsToWorkbookXml(text, settings),
        },
      ]
    : []
}
