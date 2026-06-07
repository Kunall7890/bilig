import { readFileSync, statSync } from 'node:fs'
import { resolve } from 'node:path'

import { readXlsxFormulaCacheCellsFromFile } from '@bilig/xlsx/formula-cache-reader'
import { buildWorkbookCompatibilityReportFromFile } from '@bilig/xlsx/workbook-compatibility-report'
import { tryInspectLargeSimpleXlsxHeadless } from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import type { LargeSimpleXlsxImportStats } from '../packages/excel-import/src/xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazyFromByteSource } from '../packages/excel-import/src/xlsx-zip.js'
import { startSelfRssGuard } from './public-workbook-corpus-process.ts'
import {
  fingerprintLargeSimpleDataOnlyWorkbookSource,
  fingerprintFormulaFreeWorkbookFootprint,
  fingerprintWorkbookBytes,
  inspectWorkbookFootprintForWorker,
  sha256HexSync,
  type WorkbookFootprint,
} from './public-workbook-corpus-workbook.ts'
import type { PublicWorkbookFeatureCounts } from './public-workbook-corpus-types.ts'
import { FileBackedXlsxZipByteSource, isZipWorkbookSource } from './public-workbook-corpus-xlsx-byte-source.ts'
import { inspectXlsxWorkbookFootprintLowMemoryFromByteSource } from './public-workbook-corpus-xlsx-footprint.ts'

const publicWorkbookCorpusWorkerMaterializedBytesFallbackLimit = 1_000_000

export async function writeFingerprintArtifactResult(args: {
  readonly filePath: string
  readonly fileName: string
  readonly fingerprintTimeoutMs: number
  readonly fingerprintMaxRssBytes: number
}): Promise<void> {
  if (!args.filePath) {
    throw new Error('Expected --file for fingerprint-artifact')
  }
  const { fingerprintWorkbookFileIsolated } = await import('./public-workbook-corpus-fetch.ts')
  const workbookFingerprint = await fingerprintWorkbookFileIsolated(resolve(args.filePath), args.fileName, args.fingerprintTimeoutMs, {
    maxRssBytes: args.fingerprintMaxRssBytes,
    rssCheckIntervalMs: 250,
  })
  process.stdout.write(`${JSON.stringify({ workbookFingerprint })}\n`)
}

export function writeFingerprintArtifactWorkerResult(args: {
  readonly filePath: string
  readonly fileName: string
  readonly fingerprintMaxRssBytes: number
}): void {
  const stopSelfRssGuard = startSelfRssGuard(args.fingerprintMaxRssBytes, 'Workbook fingerprinting worker')
  try {
    if (!args.filePath) {
      throw new Error('Expected --file for fingerprint-artifact-worker')
    }
    const filePath = resolve(args.filePath)
    const workbookFingerprint = tryFingerprintWorkbookFromFile(filePath, args.fileName)
    if (workbookFingerprint) {
      process.stdout.write(`${JSON.stringify({ workbookFingerprint })}\n`)
      return
    }
    assertMaterializedWorkbookFallbackWithinLimit(filePath, 'Workbook fingerprinting worker')
    const fallbackFingerprint = fingerprintWorkbookBytes(readFileSync(filePath), args.fileName)
    process.stdout.write(`${JSON.stringify({ workbookFingerprint: fallbackFingerprint })}\n`)
  } catch (error) {
    process.stderr.write(`${formatWorkerError(error)}\n`)
    process.exitCode = 1
  } finally {
    stopSelfRssGuard()
  }
}

function tryFingerprintWorkbookFromFile(filePath: string, fileName: string): string | null {
  return (
    tryFingerprintLargeSimpleWorkbookFromFile(filePath, fileName) ??
    tryFingerprintFormulaFreeWorkbookFromFile(filePath, fileName) ??
    tryFingerprintFormulaWorkbookFromFile(filePath, fileName)
  )
}

function tryFingerprintLargeSimpleWorkbookFromFile(filePath: string, fileName: string): string | null {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    return isZipWorkbookSource(source) ? fingerprintLargeSimpleDataOnlyWorkbookSource(source, fileName) : null
  } catch {
    return null
  } finally {
    source.release()
  }
}

function tryFingerprintFormulaFreeWorkbookFromFile(filePath: string, fileName: string): string | null {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    if (!isZipWorkbookSource(source)) {
      return null
    }
    const footprint = inspectXlsxWorkbookFootprintLowMemoryFromByteSource(source, fileName)
    return footprint ? fingerprintFormulaFreeWorkbookFootprint(footprint) : null
  } catch {
    return null
  } finally {
    source.release()
  }
}

function tryFingerprintFormulaWorkbookFromFile(filePath: string, fileName: string): string | null {
  try {
    const report = buildWorkbookCompatibilityReportFromFile(filePath, {
      fileName,
      inspectLimit: 1,
    })
    if (report.workbook.formulaCellCount === 0) {
      return null
    }
    const formulaScan = readXlsxFormulaCacheCellsFromFile(filePath, {
      inspectLimit: 'all',
    })
    if (formulaScan.formulaCellCount === 0) {
      return null
    }
    const counts = featureCountsFromCompatibilityReport(report.workbook)
    const metadata = {
      workbookName: fileName.replace(/\.(xlsx|xlsm|csv)$/iu, '') || fileName,
      sheetNames: report.workbook.sheetNames,
      dimensions: [],
    }
    const formulaShapes = formulaScan.cells.map((cell) => `${cell.target}:${cell.formula}`).toSorted()
    return sha256HexSync(Buffer.from(JSON.stringify({ counts, metadata, formulaShapes })))
  } catch {
    return null
  }
}

function formatWorkerError(error: unknown): string {
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`
  }
  return String(error)
}

export async function writeFootprintWorkerResult(args: {
  readonly filePath: string
  readonly fileName: string
  readonly verifyMaxRssBytes: number
}): Promise<void> {
  const stopSelfRssGuard = startSelfRssGuard(args.verifyMaxRssBytes, 'Workbook footprint worker')
  try {
    const filePath = args.filePath ? resolve(args.filePath) : null
    const footprintFromFile = filePath ? tryInspectWorkbookFootprintFromFile(filePath, args.fileName) : null
    if (footprintFromFile) {
      process.stdout.write(`${JSON.stringify({ footprint: footprintFromFile })}\n`)
      return
    }
    if (filePath) {
      assertMaterializedWorkbookFallbackWithinLimit(filePath, 'Workbook footprint worker')
    }
    const bytes = filePath ? readFileSync(filePath) : readFileSync(0)
    const footprint = await inspectWorkbookFootprintForWorker(bytes, args.fileName)
    process.stdout.write(`${JSON.stringify({ footprint })}\n`)
  } finally {
    stopSelfRssGuard()
  }
}

function tryInspectWorkbookFootprintFromFile(filePath: string, fileName: string): WorkbookFootprint | null {
  return (
    tryInspectLargeSimpleWorkbookFootprintFromFile(filePath, fileName) ?? tryInspectNativeCompatibilityFootprintFromFile(filePath, fileName)
  )
}

function tryInspectLargeSimpleWorkbookFootprintFromFile(filePath: string, fileName: string): WorkbookFootprint | null {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    if (!isZipWorkbookSource(source)) {
      return null
    }
    const sourceFootprint = inspectXlsxWorkbookFootprintLowMemoryFromByteSource(source, fileName)
    if (sourceFootprint) {
      return sourceFootprint
    }
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      return null
    }
    const inspected = tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })
    return inspected ? footprintFromLargeSimpleInspect(inspected) : null
  } finally {
    source.release()
  }
}

function tryInspectNativeCompatibilityFootprintFromFile(filePath: string, fileName: string): WorkbookFootprint | null {
  try {
    const report = buildWorkbookCompatibilityReportFromFile(filePath, {
      fileName,
      inspectLimit: 1,
    })
    return {
      featureCounts: featureCountsFromCompatibilityReport(report.workbook),
      workbookMetadata: {
        workbookName: fileName.replace(/\.(xlsx|xlsm|csv)$/iu, '') || fileName,
        sheetNames: report.workbook.sheetNames,
        dimensions: [],
      },
      externalWorkbookReferences: [],
    }
  } catch {
    return null
  }
}

function footprintFromLargeSimpleInspect(inspected: NonNullable<ReturnType<typeof tryInspectLargeSimpleXlsxHeadless>>): WorkbookFootprint {
  const featureCounts = featureCountsFromLargeSimpleStats(inspected.stats)
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: inspected.workbookName,
      sheetNames: inspected.sheetNames,
      dimensions: inspected.stats.dimensions,
    },
    externalWorkbookReferences: [],
    largeSimpleXlsxImport: { eligible: true, blockers: [] },
  }
}

function featureCountsFromCompatibilityReport(workbook: {
  readonly sheetCount: number
  readonly nonEmptyCellCount: number
  readonly formulaCellCount: number
  readonly definedNameCount: number
  readonly tableCount: number
  readonly pivotTableCount: number
  readonly chartCount: number
  readonly macroModuleCount: number
}): PublicWorkbookFeatureCounts {
  return {
    sheetCount: workbook.sheetCount,
    cellCount: workbook.nonEmptyCellCount,
    formulaCellCount: workbook.formulaCellCount,
    valueCellCount: Math.max(0, workbook.nonEmptyCellCount - workbook.formulaCellCount),
    definedNameCount: workbook.definedNameCount,
    tableCount: workbook.tableCount,
    chartCount: workbook.chartCount,
    pivotCount: workbook.pivotTableCount,
    mergeCount: 0,
    styleRangeCount: 0,
    conditionalFormatCount: 0,
    dataValidationCount: 0,
    macroPayloadCount: workbook.macroModuleCount,
    warningCount: 0,
  }
}

function featureCountsFromLargeSimpleStats(stats: LargeSimpleXlsxImportStats): WorkbookFootprint['featureCounts'] {
  return {
    sheetCount: stats.sheetCount,
    cellCount: stats.cellCount,
    formulaCellCount: stats.formulaCellCount,
    valueCellCount: stats.valueCellCount,
    definedNameCount: stats.definedNameCount,
    tableCount: stats.tableCount,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: stats.mergeCount,
    styleRangeCount: 0,
    conditionalFormatCount: stats.conditionalFormatCount,
    dataValidationCount: stats.dataValidationCount ?? 0,
    macroPayloadCount: 0,
    warningCount: stats.warningCount,
  }
}

function assertMaterializedWorkbookFallbackWithinLimit(filePath: string, phase: string): void {
  const byteLength = statSync(filePath).size
  if (byteLength <= publicWorkbookCorpusWorkerMaterializedBytesFallbackLimit) {
    return
  }
  throw new Error(
    `${phase} materialized bytes fallback is small-workbook only (${byteLength.toLocaleString('en-US')} bytes > ` +
      `${publicWorkbookCorpusWorkerMaterializedBytesFallbackLimit.toLocaleString('en-US')} bytes). Use native file-backed XLSX scanners for large corpus workbooks.`,
  )
}
