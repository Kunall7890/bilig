import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from './workbook-range-records.js'
import {
  cloneChartRecord,
  cloneDrawingArtifactsRecord,
  cloneExternalLinkArtifactsRecord,
  cloneImageRecord,
  cloneSheetDrawingArtifactsRecord,
  cloneSheetLegacyCommentVmlRecord,
  cloneSheetThreadedCommentArtifactsRecord,
  cloneShapeRecord,
  cloneThreadedCommentArtifactsRecord,
} from './workbook-metadata-records.js'
import {
  chartKey,
  imageKey,
  shapeKey,
  type WorkbookChartRecord,
  type WorkbookDrawingArtifactsRecord,
  type WorkbookExternalLinkArtifactsRecord,
  type WorkbookImageRecord,
  type WorkbookMetadataRecord,
  type WorkbookSheetDrawingArtifactsRecord,
  type WorkbookSheetLegacyCommentVmlRecord,
  type WorkbookSheetThreadedCommentArtifactsRecord,
  type WorkbookShapeRecord,
  type WorkbookThreadedCommentArtifactsRecord,
} from './workbook-metadata-types.js'
import type { WorkbookMetadataService } from './workbook-metadata-service-contract.js'
import { metadataEffect } from './workbook-metadata-service-helpers.js'

type WorkbookMetadataDrawingService = Pick<
  WorkbookMetadataService,
  | 'setDrawingArtifacts'
  | 'getDrawingArtifacts'
  | 'clearDrawingArtifacts'
  | 'setExternalLinkArtifacts'
  | 'getExternalLinkArtifacts'
  | 'clearExternalLinkArtifacts'
  | 'setThreadedCommentArtifacts'
  | 'getThreadedCommentArtifacts'
  | 'clearThreadedCommentArtifacts'
  | 'setSheetDrawingArtifacts'
  | 'getSheetDrawingArtifacts'
  | 'deleteSheetDrawingArtifacts'
  | 'setSheetThreadedCommentArtifacts'
  | 'getSheetThreadedCommentArtifacts'
  | 'deleteSheetThreadedCommentArtifacts'
  | 'setSheetLegacyCommentVml'
  | 'getSheetLegacyCommentVml'
  | 'deleteSheetLegacyCommentVml'
  | 'setChart'
  | 'getChart'
  | 'deleteChart'
  | 'listCharts'
  | 'setImage'
  | 'getImage'
  | 'deleteImage'
  | 'listImages'
  | 'setShape'
  | 'getShape'
  | 'deleteShape'
  | 'listShapes'
>

export function createWorkbookMetadataDrawingService(metadata: WorkbookMetadataRecord): WorkbookMetadataDrawingService {
  return {
    setDrawingArtifacts(artifacts) {
      return metadataEffect('Failed to set workbook drawing artifact metadata', () => {
        const stored: WorkbookDrawingArtifactsRecord = cloneDrawingArtifactsRecord(artifacts)
        metadata.drawingArtifacts = stored
        return cloneDrawingArtifactsRecord(stored)
      })
    },
    getDrawingArtifacts() {
      return metadataEffect('Failed to get workbook drawing artifact metadata', () =>
        metadata.drawingArtifacts ? cloneDrawingArtifactsRecord(metadata.drawingArtifacts) : undefined,
      )
    },
    clearDrawingArtifacts() {
      return metadataEffect('Failed to clear workbook drawing artifact metadata', () => {
        const hadArtifacts = metadata.drawingArtifacts !== undefined
        metadata.drawingArtifacts = undefined
        return hadArtifacts
      })
    },
    setExternalLinkArtifacts(artifacts) {
      return metadataEffect('Failed to set workbook external link artifact metadata', () => {
        const stored: WorkbookExternalLinkArtifactsRecord = cloneExternalLinkArtifactsRecord(artifacts)
        metadata.externalLinkArtifacts = stored
        return cloneExternalLinkArtifactsRecord(stored)
      })
    },
    getExternalLinkArtifacts() {
      return metadataEffect('Failed to get workbook external link artifact metadata', () =>
        metadata.externalLinkArtifacts ? cloneExternalLinkArtifactsRecord(metadata.externalLinkArtifacts) : undefined,
      )
    },
    clearExternalLinkArtifacts() {
      return metadataEffect('Failed to clear workbook external link artifact metadata', () => {
        const hadArtifacts = metadata.externalLinkArtifacts !== undefined
        metadata.externalLinkArtifacts = undefined
        return hadArtifacts
      })
    },
    setThreadedCommentArtifacts(artifacts) {
      return metadataEffect('Failed to set workbook threaded comment artifact metadata', () => {
        const stored: WorkbookThreadedCommentArtifactsRecord = cloneThreadedCommentArtifactsRecord(artifacts)
        metadata.threadedCommentArtifacts = stored
        return cloneThreadedCommentArtifactsRecord(stored)
      })
    },
    getThreadedCommentArtifacts() {
      return metadataEffect('Failed to get workbook threaded comment artifact metadata', () =>
        metadata.threadedCommentArtifacts ? cloneThreadedCommentArtifactsRecord(metadata.threadedCommentArtifacts) : undefined,
      )
    },
    clearThreadedCommentArtifacts() {
      return metadataEffect('Failed to clear workbook threaded comment artifact metadata', () => {
        const hadArtifacts = metadata.threadedCommentArtifacts !== undefined
        metadata.threadedCommentArtifacts = undefined
        return hadArtifacts
      })
    },
    setSheetDrawingArtifacts(sheetName, artifacts) {
      return metadataEffect('Failed to set sheet drawing artifact metadata', () => {
        const stored: WorkbookSheetDrawingArtifactsRecord = {
          sheetName,
          relationshipTarget: artifacts.relationshipTarget,
          ...(artifacts.preservedChartRelationshipIds !== undefined
            ? { preservedChartRelationshipIds: [...artifacts.preservedChartRelationshipIds] }
            : {}),
        }
        metadata.sheetDrawingArtifacts.set(sheetName, stored)
        return cloneSheetDrawingArtifactsRecord(stored)
      })
    },
    getSheetDrawingArtifacts(sheetName) {
      return metadataEffect('Failed to get sheet drawing artifact metadata', () => {
        const record = metadata.sheetDrawingArtifacts.get(sheetName)
        return record ? cloneSheetDrawingArtifactsRecord(record) : undefined
      })
    },
    deleteSheetDrawingArtifacts(sheetName) {
      return metadataEffect('Failed to delete sheet drawing artifact metadata', () => metadata.sheetDrawingArtifacts.delete(sheetName))
    },
    setSheetThreadedCommentArtifacts(sheetName, artifacts) {
      return metadataEffect('Failed to set sheet threaded comment artifact metadata', () => {
        const stored: WorkbookSheetThreadedCommentArtifactsRecord = {
          sheetName,
          relationships: structuredClone(artifacts.relationships),
        }
        metadata.sheetThreadedCommentArtifacts.set(sheetName, stored)
        return cloneSheetThreadedCommentArtifactsRecord(stored)
      })
    },
    getSheetThreadedCommentArtifacts(sheetName) {
      return metadataEffect('Failed to get sheet threaded comment artifact metadata', () => {
        const record = metadata.sheetThreadedCommentArtifacts.get(sheetName)
        return record ? cloneSheetThreadedCommentArtifactsRecord(record) : undefined
      })
    },
    deleteSheetThreadedCommentArtifacts(sheetName) {
      return metadataEffect('Failed to delete sheet threaded comment artifact metadata', () =>
        metadata.sheetThreadedCommentArtifacts.delete(sheetName),
      )
    },
    setSheetLegacyCommentVml(sheetName, legacyCommentVml) {
      return metadataEffect('Failed to set legacy comment VML metadata', () => {
        const stored: WorkbookSheetLegacyCommentVmlRecord = {
          sheetName,
          ...structuredClone(legacyCommentVml),
        }
        metadata.sheetLegacyCommentVml.set(sheetName, stored)
        return cloneSheetLegacyCommentVmlRecord(stored)
      })
    },
    getSheetLegacyCommentVml(sheetName) {
      return metadataEffect('Failed to get legacy comment VML metadata', () => {
        const record = metadata.sheetLegacyCommentVml.get(sheetName)
        return record ? cloneSheetLegacyCommentVmlRecord(record) : undefined
      })
    },
    deleteSheetLegacyCommentVml(sheetName) {
      return metadataEffect('Failed to delete legacy comment VML metadata', () => metadata.sheetLegacyCommentVml.delete(sheetName))
    },
    setChart(record) {
      return metadataEffect('Failed to set chart metadata', () => {
        const stored: WorkbookChartRecord = {
          id: record.id.trim(),
          sheetName: record.sheetName,
          address: canonicalWorkbookAddress(record.sheetName, record.address),
          source: canonicalWorkbookRangeRef(record.source),
          chartType: record.chartType,
          rows: record.rows,
          cols: record.cols,
          ...(record.anchor !== undefined ? { anchor: structuredClone(record.anchor) } : {}),
          ...(record.seriesOrientation !== undefined ? { seriesOrientation: record.seriesOrientation } : {}),
          ...(record.firstRowAsHeaders !== undefined ? { firstRowAsHeaders: record.firstRowAsHeaders } : {}),
          ...(record.firstColumnAsLabels !== undefined ? { firstColumnAsLabels: record.firstColumnAsLabels } : {}),
          ...(record.title !== undefined ? { title: record.title } : {}),
          ...(record.legendPosition !== undefined ? { legendPosition: record.legendPosition } : {}),
        }
        metadata.charts.set(chartKey(stored.id), stored)
        return cloneChartRecord(stored)
      })
    },
    getChart(id) {
      return metadataEffect('Failed to get chart metadata', () => {
        const record = metadata.charts.get(chartKey(id))
        return record ? cloneChartRecord(record) : undefined
      })
    },
    deleteChart(id) {
      return metadataEffect('Failed to delete chart metadata', () => metadata.charts.delete(chartKey(id)))
    },
    listCharts() {
      return metadataEffect('Failed to list chart metadata', () =>
        [...metadata.charts.values()].toSorted((left, right) => left.id.localeCompare(right.id)).map(cloneChartRecord),
      )
    },
    setImage(record) {
      return metadataEffect('Failed to set image metadata', () => {
        const stored: WorkbookImageRecord = {
          id: record.id.trim(),
          sheetName: record.sheetName,
          address: canonicalWorkbookAddress(record.sheetName, record.address),
          sourceUrl: record.sourceUrl,
          rows: record.rows,
          cols: record.cols,
          ...(record.altText !== undefined ? { altText: record.altText } : {}),
        }
        metadata.images.set(imageKey(stored.id), stored)
        return cloneImageRecord(stored)
      })
    },
    getImage(id) {
      return metadataEffect('Failed to get image metadata', () => {
        const record = metadata.images.get(imageKey(id))
        return record ? cloneImageRecord(record) : undefined
      })
    },
    deleteImage(id) {
      return metadataEffect('Failed to delete image metadata', () => metadata.images.delete(imageKey(id)))
    },
    listImages() {
      return metadataEffect('Failed to list image metadata', () =>
        [...metadata.images.values()].toSorted((left, right) => left.id.localeCompare(right.id)).map(cloneImageRecord),
      )
    },
    setShape(record) {
      return metadataEffect('Failed to set shape metadata', () => {
        const stored: WorkbookShapeRecord = {
          id: record.id.trim(),
          sheetName: record.sheetName,
          address: canonicalWorkbookAddress(record.sheetName, record.address),
          shapeType: record.shapeType,
          rows: record.rows,
          cols: record.cols,
          ...(record.text !== undefined ? { text: record.text } : {}),
          ...(record.fillColor !== undefined ? { fillColor: record.fillColor } : {}),
          ...(record.strokeColor !== undefined ? { strokeColor: record.strokeColor } : {}),
        }
        metadata.shapes.set(shapeKey(stored.id), stored)
        return cloneShapeRecord(stored)
      })
    },
    getShape(id) {
      return metadataEffect('Failed to get shape metadata', () => {
        const record = metadata.shapes.get(shapeKey(id))
        return record ? cloneShapeRecord(record) : undefined
      })
    },
    deleteShape(id) {
      return metadataEffect('Failed to delete shape metadata', () => metadata.shapes.delete(shapeKey(id)))
    },
    listShapes() {
      return metadataEffect('Failed to list shape metadata', () =>
        [...metadata.shapes.values()].toSorted((left, right) => left.id.localeCompare(right.id)).map(cloneShapeRecord),
      )
    },
  }
}
