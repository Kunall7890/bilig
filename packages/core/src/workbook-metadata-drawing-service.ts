import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from './workbook-range-records.js'
import { cloneChartRecord, cloneImageRecord, cloneShapeRecord } from './workbook-metadata-records.js'
import {
  chartKey,
  imageKey,
  shapeKey,
  type WorkbookChartRecord,
  type WorkbookImageRecord,
  type WorkbookMetadataRecord,
  type WorkbookShapeRecord,
} from './workbook-metadata-types.js'
import type { WorkbookMetadataService } from './workbook-metadata-service-contract.js'
import { metadataEffect } from './workbook-metadata-service-helpers.js'

type WorkbookMetadataDrawingService = Pick<
  WorkbookMetadataService,
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
