import { Cause, Effect, Exit } from 'effect'
import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from './workbook-range-records.js'
import {
  cloneChartRecord,
  cloneCommentThreadRecord,
  cloneConditionalFormatArtifactsRecord,
  cloneConditionalFormatRecord,
  cloneDataValidationRecord,
  cloneDefinedNameRecord,
  cloneDefinedNameValue,
  cloneFilterRecord,
  cloneHyperlinkRecord,
  cloneImageRecord,
  cloneMacroPayloadRecord,
  cloneMergeRangeRecord,
  cloneNoteRecord,
  clonePivotRecord,
  cloneRangeProtectionRecord,
  cloneSheetProtectionRecord,
  cloneSheetDrawingArtifactsRecord,
  cloneSheetLegacyCommentVmlRecord,
  cloneShapeRecord,
  cloneSortKeyRecord,
  cloneSortRecord,
  cloneSheetThreadedCommentArtifactsRecord,
  cloneTableRecord,
  conditionalFormatKey,
  dataValidationKey,
  deleteRecordsBySheet,
  mergeRangeKey,
  rangeProtectionKey,
  rekeyRecords,
  tableKey,
} from './workbook-metadata-records.js'
import {
  createWorkbookMetadataRecord,
  macroPayloadKey,
  type WorkbookCommentThreadRecord,
  type WorkbookConditionalFormatRecord,
  type WorkbookSheetConditionalFormatArtifactsRecord,
  compareDefinedNameRecords,
  definedNameKey,
  normalizeDefinedNameScope,
  type WorkbookMacroPayloadRecord,
  type WorkbookMergeRangeRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
  type WorkbookDataValidationRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFreezePaneRecord,
  type WorkbookMetadataRecord,
} from './workbook-metadata-types.js'
import type {
  WorkbookMetadataService,
  WorkbookSheetDeletionMetadataContext,
  WorkbookSheetReorderMetadataContext,
} from './workbook-metadata-service-contract.js'
import { createWorkbookMetadataCellRecordService } from './workbook-metadata-cell-record-service.js'
import { createWorkbookMetadataDrawingService } from './workbook-metadata-drawing-service.js'
import { canonicalMergeRangeRef, isSingleCellMergeRange, rangeContainsAddress, rangesIntersect } from './workbook-merge-records.js'
import {
  assertMergeRangesDoNotOverlap,
  metadataEffect,
  renameDataValidationSourceSheet,
  rewriteDefinedNameForSheetDeletion,
} from './workbook-metadata-service-helpers.js'
import {
  deleteWorkbookFilterRecord,
  deleteWorkbookSortRecord,
  getWorkbookFilterRecord,
  getWorkbookSortRecord,
  listWorkbookFilterRecords,
  listWorkbookSortRecords,
  setWorkbookFilterRecord,
  setWorkbookSortRecord,
} from './workbook-metadata-sheet-range-records.js'
import {
  getWorkbookPropertyRecord,
  getWorkbookProtectionRecord,
  listWorkbookPropertyRecords,
  setWorkbookPropertyRecord,
  setWorkbookProtectionRecord,
} from './workbook-metadata-workbook-records.js'
import { clonePreservedSheetMetadata } from './workbook-preserved-metadata.js'
import {
  renameDrawingChartPackageArtifactsSheetReferences,
  renamePreservedChartPackageArtifactsSheetReferences,
  rewriteDrawingChartPackageArtifactsForSheetDeletion,
  rewritePreservedChartPackageArtifactsForSheetDeletion,
} from './engine/services/structure-chart-artifact-rewrite.js'
import {
  renamePreservedWorkbookMetadataSheetReferences,
  rewritePreservedPivotPackageArtifactsForSheetDeletion,
  rewritePreservedWorkbookMetadataForSheetDeletion,
  rewritePreservedWorkbookMetadataForSheetReorder,
  rewritePreservedWorkbookMetadataForTableDeletion,
} from './engine/services/structure-preserved-sheet-metadata-rewrite.js'
import { rewriteThreadedCommentArtifactsForSheetDeletion } from './engine/services/structure-threaded-comment-artifact-rewrite.js'

export { WorkbookMetadataError, type WorkbookMetadataService } from './workbook-metadata-service-contract.js'

export function createWorkbookMetadataService(metadata: WorkbookMetadataRecord): WorkbookMetadataService {
  const renameMacroPayloadSheetCodeNames = (oldSheetName: string, newSheetName: string): void => {
    for (const [key, payload] of metadata.macroPayloads.entries()) {
      if (!payload.sheetCodeNames?.some((entry) => entry.sheetName === oldSheetName)) {
        continue
      }
      metadata.macroPayloads.set(key, {
        ...payload,
        sheetCodeNames: payload.sheetCodeNames.map((entry) =>
          entry.sheetName === oldSheetName
            ? { sheetName: newSheetName, codeName: entry.codeName }
            : { sheetName: entry.sheetName, codeName: entry.codeName },
        ),
      })
    }
  }

  const deleteMacroPayloadSheetCodeNames = (sheetName: string): void => {
    for (const [key, payload] of metadata.macroPayloads.entries()) {
      if (!payload.sheetCodeNames?.some((entry) => entry.sheetName === sheetName)) {
        continue
      }
      const sheetCodeNames = payload.sheetCodeNames
        .filter((entry) => entry.sheetName !== sheetName)
        .map((entry) => ({ sheetName: entry.sheetName, codeName: entry.codeName }))
      const nextPayload = { ...payload }
      if (sheetCodeNames.length > 0) {
        nextPayload.sheetCodeNames = sheetCodeNames
      } else {
        delete nextPayload.sheetCodeNames
      }
      metadata.macroPayloads.set(key, nextPayload)
    }
  }

  const renameSheetNow = (oldSheetName: string, newSheetName: string): void => {
    rekeyRecords(metadata.freezePanes, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record))
    rekeyRecords(metadata.sheetTabColors, (record) =>
      record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : { ...record },
    )
    rekeyRecords(metadata.merges, (record) =>
      record.sheetName === oldSheetName ? { ...cloneMergeRangeRecord(record), sheetName: newSheetName } : cloneMergeRangeRecord(record),
    )
    rekeyRecords(metadata.rowMetadata, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record))
    rekeyRecords(metadata.columnMetadata, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record))
    rekeyRecords(metadata.filters, (record) =>
      record.sheetName === oldSheetName || record.range.sheetName === oldSheetName
        ? {
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            range: {
              ...record.range,
              sheetName: record.range.sheetName === oldSheetName ? newSheetName : record.range.sheetName,
            },
          }
        : cloneFilterRecord(record),
    )
    rekeyRecords(metadata.sorts, (record) =>
      record.sheetName === oldSheetName || record.range.sheetName === oldSheetName
        ? {
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            range: {
              ...record.range,
              sheetName: record.range.sheetName === oldSheetName ? newSheetName : record.range.sheetName,
            },
            keys: record.keys.map(cloneSortKeyRecord),
          }
        : cloneSortRecord(record),
    )
    rekeyRecords(metadata.dataValidations, (record) => {
      const cloned = renameDataValidationSourceSheet(record, oldSheetName, newSheetName)
      if (cloned.range.sheetName === oldSheetName) {
        cloned.range.sheetName = newSheetName
      }
      return cloned
    })
    rekeyRecords(metadata.sheetProtections, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneSheetProtectionRecord(record), sheetName: newSheetName }
        : cloneSheetProtectionRecord(record),
    )
    rekeyRecords(metadata.conditionalFormats, (record) =>
      record.range.sheetName === oldSheetName
        ? {
            ...cloneConditionalFormatRecord(record),
            range: {
              ...record.range,
              sheetName: newSheetName,
            },
          }
        : cloneConditionalFormatRecord(record),
    )
    rekeyRecords(metadata.conditionalFormatArtifacts, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneConditionalFormatArtifactsRecord(record), sheetName: newSheetName }
        : cloneConditionalFormatArtifactsRecord(record),
    )
    rekeyRecords(metadata.sheetDrawingArtifacts, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneSheetDrawingArtifactsRecord(record), sheetName: newSheetName }
        : cloneSheetDrawingArtifactsRecord(record),
    )
    rekeyRecords(metadata.sheetThreadedCommentArtifacts, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneSheetThreadedCommentArtifactsRecord(record), sheetName: newSheetName }
        : cloneSheetThreadedCommentArtifactsRecord(record),
    )
    rekeyRecords(metadata.sheetLegacyCommentVml, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneSheetLegacyCommentVmlRecord(record), sheetName: newSheetName }
        : cloneSheetLegacyCommentVmlRecord(record),
    )
    const preservedSheetMetadata = metadata.preservedSheetMetadata.get(oldSheetName)
    if (preservedSheetMetadata) {
      metadata.preservedSheetMetadata.delete(oldSheetName)
      metadata.preservedSheetMetadata.set(newSheetName, clonePreservedSheetMetadata(preservedSheetMetadata))
    }
    rekeyRecords(metadata.rangeProtections, (record) =>
      record.range.sheetName === oldSheetName
        ? {
            ...cloneRangeProtectionRecord(record),
            range: {
              ...record.range,
              sheetName: newSheetName,
            },
          }
        : cloneRangeProtectionRecord(record),
    )
    rekeyRecords(metadata.commentThreads, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneCommentThreadRecord(record), sheetName: newSheetName }
        : cloneCommentThreadRecord(record),
    )
    refreshLegacyCommentVmlCommentSignature(metadata, newSheetName)
    rekeyRecords(metadata.notes, (record) =>
      record.sheetName === oldSheetName ? { ...cloneNoteRecord(record), sheetName: newSheetName } : cloneNoteRecord(record),
    )
    rekeyRecords(metadata.hyperlinks, (record) =>
      record.sheetName === oldSheetName ? { ...cloneHyperlinkRecord(record), sheetName: newSheetName } : cloneHyperlinkRecord(record),
    )
    rekeyRecords(metadata.tables, (record) =>
      record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : cloneTableRecord(record),
    )
    rekeyRecords(metadata.spills, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : { ...record }))
    rekeyRecords(metadata.pivots, (record) =>
      record.sheetName === oldSheetName || record.source?.sheetName === oldSheetName
        ? {
            ...record,
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            ...(record.source
              ? {
                  source: {
                    ...record.source,
                    sheetName: record.source.sheetName === oldSheetName ? newSheetName : record.source.sheetName,
                  },
                }
              : {}),
            groupBy: [...record.groupBy],
            values: record.values.map((value) => ({ ...value })),
          }
        : clonePivotRecord(record),
    )
    rekeyRecords(metadata.charts, (record) =>
      record.sheetName === oldSheetName || record.source.sheetName === oldSheetName
        ? {
            ...cloneChartRecord(record),
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            source: {
              ...record.source,
              sheetName: record.source.sheetName === oldSheetName ? newSheetName : record.source.sheetName,
            },
          }
        : cloneChartRecord(record),
    )
    const drawingArtifacts = renameDrawingChartPackageArtifactsSheetReferences(metadata.drawingArtifacts, oldSheetName, newSheetName)
    if (drawingArtifacts) {
      metadata.drawingArtifacts = drawingArtifacts
    }
    const renamedPreservedWorkbookMetadata = renamePreservedWorkbookMetadataSheetReferences(
      metadata.preservedWorkbookMetadata,
      oldSheetName,
      newSheetName,
    )
    if (renamedPreservedWorkbookMetadata) {
      metadata.preservedWorkbookMetadata = renamedPreservedWorkbookMetadata
    }
    const preservedWorkbookMetadata = renamePreservedChartPackageArtifactsSheetReferences(
      metadata.preservedWorkbookMetadata,
      oldSheetName,
      newSheetName,
    )
    if (preservedWorkbookMetadata) {
      metadata.preservedWorkbookMetadata = preservedWorkbookMetadata
    }
    rekeyRecords(metadata.images, (record) =>
      record.sheetName === oldSheetName ? { ...cloneImageRecord(record), sheetName: newSheetName } : cloneImageRecord(record),
    )
    rekeyRecords(metadata.shapes, (record) =>
      record.sheetName === oldSheetName ? { ...cloneShapeRecord(record), sheetName: newSheetName } : cloneShapeRecord(record),
    )
    rekeyRecords(metadata.definedNames, (record) =>
      record.scopeSheetName === oldSheetName
        ? { ...cloneDefinedNameRecord(record), scopeSheetName: newSheetName }
        : cloneDefinedNameRecord(record),
    )
    renameMacroPayloadSheetCodeNames(oldSheetName, newSheetName)
  }

  const deleteSheetRecordsNow = (sheetName: string, context?: WorkbookSheetDeletionMetadataContext): void => {
    const deletedPreservedSheetMetadata = metadata.preservedSheetMetadata.get(sheetName)
    if (context) {
      const preservedWorkbookMetadata = rewritePreservedWorkbookMetadataForSheetDeletion(
        metadata.preservedWorkbookMetadata,
        sheetName,
        context,
      )
      if (preservedWorkbookMetadata) {
        metadata.preservedWorkbookMetadata = preservedWorkbookMetadata
      }
    }
    const pivotPreservedWorkbookMetadata = rewritePreservedPivotPackageArtifactsForSheetDeletion(
      metadata.preservedWorkbookMetadata,
      deletedPreservedSheetMetadata,
      context,
    )
    if (pivotPreservedWorkbookMetadata) {
      metadata.preservedWorkbookMetadata = pivotPreservedWorkbookMetadata
    }
    const preservedWorkbookMetadata = rewritePreservedChartPackageArtifactsForSheetDeletion(metadata.preservedWorkbookMetadata, sheetName)
    if (preservedWorkbookMetadata) {
      metadata.preservedWorkbookMetadata = preservedWorkbookMetadata
    }
    const drawingArtifacts = rewriteDrawingChartPackageArtifactsForSheetDeletion(metadata.drawingArtifacts, sheetName)
    if (drawingArtifacts) {
      metadata.drawingArtifacts = drawingArtifacts
    }
    deleteRecordsBySheet(metadata.definedNames, sheetName, (record) => record.scopeSheetName)
    rekeyRecords(metadata.definedNames, (record) => rewriteDefinedNameForSheetDeletion(record, sheetName))
    deleteRecordsBySheet(metadata.tables, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.spills, sheetName, (record) => record.sheetName)
    for (const [key, record] of metadata.pivots.entries()) {
      if (record.sheetName === sheetName || record.source?.sheetName === sheetName) {
        metadata.pivots.delete(key)
      }
    }
    for (const [key, record] of metadata.charts.entries()) {
      if (record.sheetName === sheetName || record.source.sheetName === sheetName) {
        metadata.charts.delete(key)
      }
    }
    deleteRecordsBySheet(metadata.images, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.shapes, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.rowMetadata, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.columnMetadata, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.merges, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.filters, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.sorts, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.dataValidations, sheetName, (record) => record.range.sheetName)
    metadata.sheetProtections.delete(sheetName)
    deleteRecordsBySheet(metadata.conditionalFormats, sheetName, (record) => record.range.sheetName)
    metadata.conditionalFormatArtifacts.delete(sheetName)
    metadata.sheetDrawingArtifacts.delete(sheetName)
    metadata.threadedCommentArtifacts = rewriteThreadedCommentArtifactsForSheetDeletion({
      workbookArtifacts: metadata.threadedCommentArtifacts,
      sheetArtifactsByName: metadata.sheetThreadedCommentArtifacts,
      deletedSheetName: sheetName,
    })
    metadata.sheetThreadedCommentArtifacts.delete(sheetName)
    metadata.sheetLegacyCommentVml.delete(sheetName)
    metadata.preservedSheetMetadata.delete(sheetName)
    deleteRecordsBySheet(metadata.rangeProtections, sheetName, (record) => record.range.sheetName)
    deleteRecordsBySheet(metadata.commentThreads, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.notes, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.hyperlinks, sheetName, (record) => record.sheetName)
    metadata.freezePanes.delete(sheetName)
    metadata.sheetTabColors.delete(sheetName)
    deleteMacroPayloadSheetCodeNames(sheetName)
  }

  const reorderSheetRecordsNow = (context: WorkbookSheetReorderMetadataContext): void => {
    const preservedWorkbookMetadata = rewritePreservedWorkbookMetadataForSheetReorder(metadata.preservedWorkbookMetadata, context)
    if (preservedWorkbookMetadata) {
      metadata.preservedWorkbookMetadata = preservedWorkbookMetadata
    }
  }

  const resetNow = (): void => {
    const defaults = createWorkbookMetadataRecord()
    metadata.properties.clear()
    metadata.workbookProtection = defaults.workbookProtection
    metadata.macroPayloads.clear()
    metadata.definedNames.clear()
    metadata.tables.clear()
    metadata.spills.clear()
    metadata.pivots.clear()
    metadata.charts.clear()
    metadata.images.clear()
    metadata.shapes.clear()
    metadata.drawingArtifacts = defaults.drawingArtifacts
    metadata.controlArtifacts = defaults.controlArtifacts
    metadata.externalLinkArtifacts = defaults.externalLinkArtifacts
    metadata.sheetDrawingArtifacts.clear()
    metadata.threadedCommentArtifacts = defaults.threadedCommentArtifacts
    metadata.sheetThreadedCommentArtifacts.clear()
    metadata.sheetLegacyCommentVml.clear()
    metadata.preservedWorkbookMetadata = {}
    metadata.preservedSheetMetadata.clear()
    metadata.rowMetadata.clear()
    metadata.columnMetadata.clear()
    metadata.freezePanes.clear()
    metadata.sheetTabColors.clear()
    metadata.merges.clear()
    metadata.sheetProtections.clear()
    metadata.filters.clear()
    metadata.sorts.clear()
    metadata.dataValidations.clear()
    metadata.conditionalFormats.clear()
    metadata.conditionalFormatArtifacts.clear()
    metadata.rangeProtections.clear()
    metadata.commentThreads.clear()
    metadata.notes.clear()
    metadata.hyperlinks.clear()
    metadata.calculationSettings = defaults.calculationSettings
    metadata.volatileContext = defaults.volatileContext
  }

  return {
    renameSheet(oldSheetName, newSheetName) {
      return metadataEffect('Failed to rename workbook sheet metadata', () => renameSheetNow(oldSheetName, newSheetName))
    },
    reorderSheetRecords(context) {
      return metadataEffect('Failed to reorder workbook sheet metadata', () => reorderSheetRecordsNow(context))
    },
    deleteSheetRecords(sheetName, context) {
      return metadataEffect('Failed to delete workbook sheet metadata', () => deleteSheetRecordsNow(sheetName, context))
    },
    reset() {
      return metadataEffect('Failed to reset workbook metadata', resetNow)
    },
    setWorkbookProperty(key, value) {
      return metadataEffect('Failed to set workbook property', () => setWorkbookPropertyRecord(metadata, key, value))
    },
    getWorkbookProperty(key) {
      return metadataEffect('Failed to get workbook property', () => getWorkbookPropertyRecord(metadata, key))
    },
    listWorkbookProperties() {
      return metadataEffect('Failed to list workbook properties', () => listWorkbookPropertyRecords(metadata))
    },
    setWorkbookProtection(record) {
      return metadataEffect('Failed to set workbook protection metadata', () => setWorkbookProtectionRecord(metadata, record))
    },
    getWorkbookProtection() {
      return metadataEffect('Failed to get workbook protection metadata', () => getWorkbookProtectionRecord(metadata))
    },
    setMacroPayload(record) {
      return metadataEffect('Failed to set macro payload metadata', () => {
        const stored: WorkbookMacroPayloadRecord = cloneMacroPayloadRecord(record)
        metadata.macroPayloads.set(macroPayloadKey(stored.kind), stored)
        return cloneMacroPayloadRecord(stored)
      })
    },
    listMacroPayloads() {
      return metadataEffect('Failed to list macro payload metadata', () =>
        [...metadata.macroPayloads.values()]
          .toSorted((left, right) => macroPayloadKey(left.kind).localeCompare(macroPayloadKey(right.kind)))
          .map(cloneMacroPayloadRecord),
      )
    },
    setCalculationSettings(settings) {
      return metadataEffect('Failed to set calculation settings', () => {
        metadata.calculationSettings = {
          compatibilityMode: 'excel-modern',
          ...settings,
        }
        return { ...metadata.calculationSettings }
      })
    },
    getCalculationSettings() {
      return metadataEffect('Failed to get calculation settings', () => ({ ...metadata.calculationSettings }))
    },
    setVolatileContext(context) {
      return metadataEffect('Failed to set volatile context', () => {
        metadata.volatileContext = { ...context }
        return { ...metadata.volatileContext }
      })
    },
    getVolatileContext() {
      return metadataEffect('Failed to get volatile context', () => ({ ...metadata.volatileContext }))
    },
    setDefinedName(name, value, scopeSheetName) {
      return metadataEffect('Failed to set defined name', () => {
        const trimmedName = name.trim()
        const normalizedScope = normalizeDefinedNameScope(scopeSheetName)
        const record: WorkbookDefinedNameRecord = {
          name: trimmedName,
          ...(normalizedScope !== undefined ? { scopeSheetName: normalizedScope } : {}),
          value: cloneDefinedNameValue(value),
        }
        metadata.definedNames.set(definedNameKey(trimmedName, normalizedScope), record)
        return cloneDefinedNameRecord(record)
      })
    },
    getDefinedName(name, scopeSheetName) {
      return metadataEffect('Failed to get defined name', () => {
        const scopedKey = definedNameKey(name, scopeSheetName)
        const record = metadata.definedNames.get(scopedKey) ?? metadata.definedNames.get(definedNameKey(name))
        return record ? cloneDefinedNameRecord(record) : undefined
      })
    },
    deleteDefinedName(name, scopeSheetName) {
      return metadataEffect('Failed to delete defined name', () => metadata.definedNames.delete(definedNameKey(name, scopeSheetName)))
    },
    listDefinedNames() {
      return metadataEffect('Failed to list defined names', () =>
        [...metadata.definedNames.values()].toSorted(compareDefinedNameRecords).map(cloneDefinedNameRecord),
      )
    },
    setTable(record) {
      return metadataEffect('Failed to set table metadata', () => {
        const stored = cloneTableRecord(record)
        stored.name = stored.name.trim()
        metadata.tables.set(tableKey(stored.name), stored)
        return cloneTableRecord(stored)
      })
    },
    getTable(name) {
      return metadataEffect('Failed to get table metadata', () => {
        const record = metadata.tables.get(tableKey(name))
        return record ? cloneTableRecord(record) : undefined
      })
    },
    deleteTable(name) {
      return metadataEffect('Failed to delete table metadata', () => {
        const existing = metadata.tables.get(tableKey(name))
        if (!existing) {
          return false
        }
        const preservedWorkbookMetadata = rewritePreservedWorkbookMetadataForTableDeletion(
          metadata.preservedWorkbookMetadata,
          existing.name,
        )
        if (preservedWorkbookMetadata) {
          metadata.preservedWorkbookMetadata = preservedWorkbookMetadata
        }
        return metadata.tables.delete(tableKey(existing.name))
      })
    },
    listTables() {
      return metadataEffect('Failed to list table metadata', () =>
        [...metadata.tables.values()]
          .toSorted((left, right) => tableKey(left.name).localeCompare(tableKey(right.name)))
          .map(cloneTableRecord),
      )
    },
    setFreezePane(sheetName, rows, cols, options) {
      return metadataEffect('Failed to set freeze pane metadata', () => {
        const record: WorkbookFreezePaneRecord = { sheetName, rows, cols }
        if (options?.topLeftCell !== undefined) {
          record.topLeftCell = options.topLeftCell
        }
        if (options?.activePane !== undefined) {
          record.activePane = options.activePane
        }
        metadata.freezePanes.set(sheetName, record)
        return { ...record }
      })
    },
    getFreezePane(sheetName) {
      return metadataEffect('Failed to get freeze pane metadata', () => {
        const record = metadata.freezePanes.get(sheetName)
        return record ? { ...record } : undefined
      })
    },
    clearFreezePane(sheetName) {
      return metadataEffect('Failed to clear freeze pane metadata', () => metadata.freezePanes.delete(sheetName))
    },
    setSheetTabColor(sheetName, tabColor) {
      return metadataEffect('Failed to set sheet tab color metadata', () => {
        const record = { sheetName, ...tabColor }
        metadata.sheetTabColors.set(sheetName, record)
        return { ...record }
      })
    },
    getSheetTabColor(sheetName) {
      return metadataEffect('Failed to get sheet tab color metadata', () => {
        const record = metadata.sheetTabColors.get(sheetName)
        return record ? { ...record } : undefined
      })
    },
    clearSheetTabColor(sheetName) {
      return metadataEffect('Failed to clear sheet tab color metadata', () => metadata.sheetTabColors.delete(sheetName))
    },
    setMergeRange(range) {
      return metadataEffect('Failed to set merged cell metadata', () => {
        const stored = canonicalMergeRangeRef(range)
        if (isSingleCellMergeRange(stored)) {
          throw new Error('Merged ranges must include at least two cells')
        }
        const overlapping = [...metadata.merges.values()].filter((record) => rangesIntersect(record, stored))
        if (overlapping.some((record) => mergeRangeKey(record) !== mergeRangeKey(stored))) {
          throw new Error('Merged ranges cannot overlap')
        }
        metadata.merges.set(mergeRangeKey(stored), stored)
        return cloneMergeRangeRecord(stored)
      })
    },
    setMergeRanges(sheetName, ranges) {
      return metadataEffect('Failed to set merged cell metadata ranges', () => {
        const storedRanges = ranges.map((range) => canonicalMergeRangeRef({ ...range, sheetName: range.sheetName ?? sheetName }))
        const seenKeys = new Set<string>()
        for (const stored of storedRanges) {
          if (isSingleCellMergeRange(stored)) {
            throw new Error('Merged ranges must include at least two cells')
          }
          const key = mergeRangeKey(stored)
          if (seenKeys.has(key)) {
            throw new Error('Merged ranges cannot contain duplicate ranges')
          }
          seenKeys.add(key)
        }
        assertMergeRangesDoNotOverlap(storedRanges)
        deleteRecordsBySheet(metadata.merges, sheetName, (record) => record.sheetName)
        for (const stored of storedRanges) {
          metadata.merges.set(mergeRangeKey(stored), stored)
        }
        return storedRanges.toSorted((left, right) => mergeRangeKey(left).localeCompare(mergeRangeKey(right))).map(cloneMergeRangeRecord)
      })
    },
    getMergeRange(sheetName, address) {
      return metadataEffect('Failed to get merged cell metadata', () => {
        const record = [...metadata.merges.values()].find((entry) => rangeContainsAddress(entry, sheetName, address))
        return record ? cloneMergeRangeRecord(record) : undefined
      })
    },
    getMergeRangeByRange(range) {
      return metadataEffect('Failed to get merged range metadata', () => {
        const record = metadata.merges.get(mergeRangeKey(range))
        return record ? cloneMergeRangeRecord(record) : undefined
      })
    },
    clearMergeRanges(range) {
      return metadataEffect('Failed to clear merged cell metadata', () => {
        const removed: WorkbookMergeRangeRecord[] = []
        for (const [key, record] of metadata.merges.entries()) {
          if (!rangesIntersect(record, range)) {
            continue
          }
          metadata.merges.delete(key)
          removed.push(cloneMergeRangeRecord(record))
        }
        return removed.toSorted((left, right) => mergeRangeKey(left).localeCompare(mergeRangeKey(right)))
      })
    },
    listMergeRanges(sheetName) {
      return metadataEffect('Failed to list merged cell metadata', () =>
        [...metadata.merges.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) => mergeRangeKey(left).localeCompare(mergeRangeKey(right)))
          .map(cloneMergeRangeRecord),
      )
    },
    setSheetProtection(record) {
      return metadataEffect('Failed to set sheet protection metadata', () => {
        const stored: WorkbookSheetProtectionRecord = cloneSheetProtectionRecord({
          sheetName: record.sheetName,
          ...(record.hideFormulas !== undefined ? { hideFormulas: record.hideFormulas } : {}),
          ...(record.xmlAttributes ? { xmlAttributes: record.xmlAttributes.map((attribute) => ({ ...attribute })) } : {}),
        })
        metadata.sheetProtections.set(record.sheetName, stored)
        return cloneSheetProtectionRecord(stored)
      })
    },
    getSheetProtection(sheetName) {
      return metadataEffect('Failed to get sheet protection metadata', () => {
        const record = metadata.sheetProtections.get(sheetName)
        return record ? cloneSheetProtectionRecord(record) : undefined
      })
    },
    clearSheetProtection(sheetName) {
      return metadataEffect('Failed to clear sheet protection metadata', () => metadata.sheetProtections.delete(sheetName))
    },
    setFilter(sheetName, range) {
      return metadataEffect('Failed to set filter metadata', () => setWorkbookFilterRecord(metadata, sheetName, range))
    },
    getFilter(sheetName, range) {
      return metadataEffect('Failed to get filter metadata', () => getWorkbookFilterRecord(metadata, sheetName, range))
    },
    deleteFilter(sheetName, range) {
      return metadataEffect('Failed to delete filter metadata', () => deleteWorkbookFilterRecord(metadata, sheetName, range))
    },
    listFilters(sheetName) {
      return metadataEffect('Failed to list filter metadata', () => listWorkbookFilterRecords(metadata, sheetName))
    },
    setSort(sheetName, range, keys) {
      return metadataEffect('Failed to set sort metadata', () => setWorkbookSortRecord(metadata, sheetName, range, keys))
    },
    getSort(sheetName, range) {
      return metadataEffect('Failed to get sort metadata', () => getWorkbookSortRecord(metadata, sheetName, range))
    },
    deleteSort(sheetName, range) {
      return metadataEffect('Failed to delete sort metadata', () => deleteWorkbookSortRecord(metadata, sheetName, range))
    },
    listSorts(sheetName) {
      return metadataEffect('Failed to list sort metadata', () => listWorkbookSortRecords(metadata, sheetName))
    },
    setDataValidation(record) {
      return metadataEffect('Failed to set data validation metadata', () => {
        const storedRange = canonicalWorkbookRangeRef(record.range)
        const nextRecord: WorkbookDataValidationRecord = {
          range: storedRange,
          rule: record.rule,
        }
        if (record.allowBlank !== undefined) {
          nextRecord.allowBlank = record.allowBlank
        }
        if (record.showDropdown !== undefined) {
          nextRecord.showDropdown = record.showDropdown
        }
        if (record.promptTitle !== undefined) {
          nextRecord.promptTitle = record.promptTitle
        }
        if (record.promptMessage !== undefined) {
          nextRecord.promptMessage = record.promptMessage
        }
        if (record.errorStyle !== undefined) {
          nextRecord.errorStyle = record.errorStyle
        }
        if (record.errorTitle !== undefined) {
          nextRecord.errorTitle = record.errorTitle
        }
        if (record.errorMessage !== undefined) {
          nextRecord.errorMessage = record.errorMessage
        }
        const stored: WorkbookDataValidationRecord = cloneDataValidationRecord(nextRecord)
        metadata.dataValidations.set(dataValidationKey(storedRange.sheetName, storedRange), stored)
        return cloneDataValidationRecord(stored)
      })
    },
    getDataValidation(sheetName, range) {
      return metadataEffect('Failed to get data validation metadata', () => {
        const record = metadata.dataValidations.get(dataValidationKey(sheetName, range))
        return record ? cloneDataValidationRecord(record) : undefined
      })
    },
    deleteDataValidation(sheetName, range) {
      return metadataEffect('Failed to delete data validation metadata', () =>
        metadata.dataValidations.delete(dataValidationKey(sheetName, range)),
      )
    },
    listDataValidations(sheetName) {
      return metadataEffect('Failed to list data validation metadata', () =>
        [...metadata.dataValidations.values()]
          .filter((record) => record.range.sheetName === sheetName)
          .toSorted((left, right) =>
            dataValidationKey(left.range.sheetName, left.range).localeCompare(dataValidationKey(right.range.sheetName, right.range)),
          )
          .map(cloneDataValidationRecord),
      )
    },
    setConditionalFormat(record) {
      return metadataEffect('Failed to set conditional format metadata', () => {
        const id = conditionalFormatKey(record.id)
        const nextRecord: WorkbookConditionalFormatRecord = {
          id,
          range: canonicalWorkbookRangeRef(record.range),
          rule: structuredClone(record.rule),
          style: structuredClone(record.style),
        }
        if (record.stopIfTrue !== undefined) {
          nextRecord.stopIfTrue = record.stopIfTrue
        }
        if (record.priority !== undefined) {
          nextRecord.priority = record.priority
        }
        const stored = cloneConditionalFormatRecord(nextRecord)
        metadata.conditionalFormats.set(id, stored)
        return cloneConditionalFormatRecord(stored)
      })
    },
    getConditionalFormat(id) {
      return metadataEffect('Failed to get conditional format metadata', () => {
        const record = metadata.conditionalFormats.get(conditionalFormatKey(id))
        return record ? cloneConditionalFormatRecord(record) : undefined
      })
    },
    deleteConditionalFormat(id) {
      return metadataEffect('Failed to delete conditional format metadata', () =>
        metadata.conditionalFormats.delete(conditionalFormatKey(id)),
      )
    },
    listConditionalFormats(sheetName) {
      return metadataEffect('Failed to list conditional format metadata', () =>
        [...metadata.conditionalFormats.values()]
          .filter((record) => record.range.sheetName === sheetName)
          .toSorted((left, right) => {
            const priorityCompare = (left.priority ?? Number.MAX_SAFE_INTEGER) - (right.priority ?? Number.MAX_SAFE_INTEGER)
            if (priorityCompare !== 0) {
              return priorityCompare
            }
            return left.id.localeCompare(right.id)
          })
          .map(cloneConditionalFormatRecord),
      )
    },
    setConditionalFormatArtifacts(sheetName, artifacts) {
      return metadataEffect('Failed to set conditional format artifact metadata', () => {
        const stored: WorkbookSheetConditionalFormatArtifactsRecord = {
          sheetName,
          xml: artifacts.xml,
        }
        metadata.conditionalFormatArtifacts.set(sheetName, stored)
        return cloneConditionalFormatArtifactsRecord(stored)
      })
    },
    getConditionalFormatArtifacts(sheetName) {
      return metadataEffect('Failed to get conditional format artifact metadata', () => {
        const record = metadata.conditionalFormatArtifacts.get(sheetName)
        return record ? cloneConditionalFormatArtifactsRecord(record) : undefined
      })
    },
    deleteConditionalFormatArtifacts(sheetName) {
      return metadataEffect('Failed to delete conditional format artifact metadata', () =>
        metadata.conditionalFormatArtifacts.delete(sheetName),
      )
    },
    setRangeProtection(record) {
      return metadataEffect('Failed to set range protection metadata', () => {
        const id = rangeProtectionKey(record.id)
        const stored: WorkbookRangeProtectionRecord = cloneRangeProtectionRecord({
          ...structuredClone(record),
          id,
          range: canonicalWorkbookRangeRef(record.range),
        })
        metadata.rangeProtections.set(id, stored)
        return cloneRangeProtectionRecord(stored)
      })
    },
    getRangeProtection(id) {
      return metadataEffect('Failed to get range protection metadata', () => {
        const record = metadata.rangeProtections.get(rangeProtectionKey(id))
        return record ? cloneRangeProtectionRecord(record) : undefined
      })
    },
    deleteRangeProtection(id) {
      return metadataEffect('Failed to delete range protection metadata', () => metadata.rangeProtections.delete(rangeProtectionKey(id)))
    },
    listRangeProtections(sheetName) {
      return metadataEffect('Failed to list range protection metadata', () =>
        [...metadata.rangeProtections.values()]
          .filter((record) => record.range.sheetName === sheetName)
          .toSorted((left, right) => left.id.localeCompare(right.id))
          .map(cloneRangeProtectionRecord),
      )
    },
    ...createWorkbookMetadataCellRecordService(metadata),
    ...createWorkbookMetadataDrawingService(metadata),
  }
}

function refreshLegacyCommentVmlCommentSignature(metadata: WorkbookMetadataRecord, sheetName: string): void {
  const legacyCommentVml = metadata.sheetLegacyCommentVml.get(sheetName)
  if (!legacyCommentVml) {
    return
  }
  const commentThreads = [...metadata.commentThreads.values()].filter((record) => record.sheetName === sheetName)
  metadata.sheetLegacyCommentVml.set(sheetName, {
    ...cloneSheetLegacyCommentVmlRecord(legacyCommentVml),
    commentSignature: legacyCommentThreadSignature(commentThreads),
  })
}

function legacyCommentThreadSignature(commentThreads: readonly WorkbookCommentThreadRecord[]): string {
  const normalized = commentThreads
    .map((thread) => ({
      sheetName: thread.sheetName,
      address: normalizeCommentAddress(thread.sheetName, thread.address),
      comments: thread.comments.map((comment) => ({
        body: comment.body,
        authorDisplayName: comment.authorDisplayName ?? '',
      })),
    }))
    .toSorted((left, right) => `${left.sheetName}:${left.address}`.localeCompare(`${right.sheetName}:${right.address}`))
  return JSON.stringify(normalized)
}

function normalizeCommentAddress(sheetName: string, address: string): string {
  try {
    return canonicalWorkbookAddress(sheetName, address)
  } catch {
    return address.trim().toUpperCase()
  }
}

export function runWorkbookMetadataEffect<Success, Failure>(effect: Effect.Effect<Success, Failure>): Success {
  const exit = Effect.runSyncExit(effect)
  if (Exit.isSuccess(exit)) {
    return exit.value
  }
  throw Cause.squash(exit.cause)
}
