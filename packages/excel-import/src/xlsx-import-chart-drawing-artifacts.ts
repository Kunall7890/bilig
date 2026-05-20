import type { WorkbookChartSnapshot } from '@bilig/protocol'
import { readImportedWorkbookChartArtifacts } from './xlsx-chart-artifacts.js'
import { readImportedWorkbookChartReadResult } from './xlsx-charts.js'
import { readImportedWorkbookDrawingArtifacts } from './xlsx-drawing-artifacts.js'
import type { XlsxZipSource } from './xlsx-zip.js'

type ImportedWorkbookChartArtifacts = ReturnType<typeof readImportedWorkbookChartArtifacts>
type ImportedWorkbookDrawingArtifacts = ReturnType<typeof readImportedWorkbookDrawingArtifacts>

export interface ImportedWorkbookChartDrawingArtifacts {
  readonly chartArtifacts: ImportedWorkbookChartArtifacts
  readonly charts: WorkbookChartSnapshot[] | undefined
  readonly drawingArtifacts: ImportedWorkbookDrawingArtifacts
}

export function readImportedWorkbookChartDrawingArtifacts(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): ImportedWorkbookChartDrawingArtifacts {
  const chartArtifacts = readImportedWorkbookChartArtifacts(source)
  const chartReadResult = readImportedWorkbookChartReadResult(source, sheetNames)
  const drawingArtifacts = readImportedWorkbookDrawingArtifacts(source, sheetNames, {
    supportedChartRelationshipIdsBySheet: chartReadResult.supportedChartRelationshipIdsBySheet,
  })
  return {
    chartArtifacts,
    charts: chartReadResult.charts,
    drawingArtifacts,
  }
}
