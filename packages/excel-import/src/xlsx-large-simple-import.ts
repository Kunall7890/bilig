import type {
  CellStyleRecord,
  WorkbookConditionalFormatSnapshot,
  SheetMetadataSnapshot,
  WorkbookSnapshot,
  WorkbookTableSnapshot,
} from '@bilig/protocol'
import { createSheetPreview, normalizeWorkbookName } from './workbook-import-helpers.js'
import { XLSX_CONTENT_TYPE } from './workbook-import-content-types.js'
import { createWorkbookPreview, type ImportedWorkbookPreview } from './workbook-import-preview.js'
import {
  readImportedSheetConditionalFormatArtifactsFromElementXml,
  readImportedSheetConditionalFormatArtifactsFromWorksheetXml,
  readImportedSheetConditionalFormatsFromElementXml,
  readImportedSheetConditionalFormatsFromWorksheetXml,
} from './xlsx-conditional-formats.js'
import { readImportedWorkbookDrawingArtifactsFromWorksheetRelationships } from './xlsx-drawing-artifacts.js'
import { readImportedSheetAutoFilters } from './xlsx-filters.js'
import { readLargeSimpleSheetHyperlinks, resolveLargeSimpleSheetHyperlinks } from './xlsx-large-simple-hyperlinks.js'
import { LargeSimpleXlsxImportPhaseRecorder, type LargeSimpleXlsxImportPhaseTelemetry } from './xlsx-large-simple-import-telemetry.js'
import { readLargeSimpleSheetPrintMetadata, readLargeSimpleSheetPrintPageSetup } from './xlsx-large-simple-printer-settings.js'
import { readAllLargeSimpleSharedStrings, readReferencedLargeSimpleSharedStrings } from './xlsx-large-simple-referenced-shared-strings.js'
import type { LargeSimpleSharedStringEntry } from './xlsx-large-simple-shared-strings.js'
import { shouldUseSharedStringlessFastPathBytes } from './xlsx-large-simple-shared-stringless-fast-path.js'
import { buildLargeSimpleStyleRanges } from './xlsx-large-simple-style-ranges.js'
import { readLargeSimpleWorkbookStylesFromChunks } from './xlsx-large-simple-styles.js'
import { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import { readWorkbookDefinedNames, readWorkbookSheets, readWorksheetPathsByRelationshipId } from './xlsx-large-simple-workbook-metadata.js'
import {
  attachImportedRuntimeCoordinates,
  createImportedRuntimeSheetCells,
  type ImportedRuntimeSheetCells,
} from './xlsx-import-runtime-coordinates.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import { importedWorksheetCellScanFromHeadless } from './xlsx-large-simple-headless-cell-scan.js'
import { parseHeadlessLargeSimpleWorksheetFromChunks } from './xlsx-large-simple-headless-worksheet-scanner.js'
import {
  hasUnsupportedLargeSimpleWorksheetTags,
  needsLargeSimpleWorksheetMetadataXml,
  readLargeSimpleWorksheetMetadataXml,
  parseLargeSimpleWorksheetCells,
} from './xlsx-large-simple-worksheet-scanner.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from './xlsx-large-simple-worksheet-stream-scanner.js'
import {
  readLargeSimpleColumnMetadata,
  readLargeSimpleDrawingRelationshipId,
  readLargeSimpleMergeRanges,
  readLargeSimpleRowMetadata,
  readLargeSimpleSheetFormatPr,
  type LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'
import { readImportedSheetTablesFromRelationshipIds, readImportedSheetTablesFromWorksheetXml } from './xlsx-tables.js'
import {
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  normalizeZipPath,
  readLazyXlsxZipSourceByteLength,
  releaseLazyXlsxZipSource,
  type XlsxZipEntries,
} from './xlsx-zip.js'

export interface LargeSimpleXlsxImportResult {
  snapshot: WorkbookSnapshot
  workbookName: string
  sheetNames: string[]
  warnings: string[]
  preview: ImportedWorkbookPreview
  stats: LargeSimpleXlsxImportStats
}

export interface LargeSimpleXlsxImportOptions {
  minByteLength?: number
  materializeCells?: boolean
  materializeMetadata?: boolean
  releaseArenaAfterMaterialization?: boolean
  releaseZipSource?: boolean
  releaseOwnedSourceBytes?: () => LargeSimpleXlsxOwnedSourceReleaseEvidence | undefined
}

export interface LargeSimpleXlsxImportSource {
  readonly byteLength: number
}

export interface LargeSimpleXlsxOwnedSourceReleaseEvidence {
  readonly ownedSourceBytesBeforeRelease?: number
  readonly ownedSourceBytesAfterRelease?: number
}

export interface LargeSimpleXlsxImportStats {
  readonly sheetCount: number
  readonly cellCount: number
  readonly formulaCellCount: number
  readonly valueCellCount: number
  readonly definedNameCount: number
  readonly tableCount: number
  readonly mergeCount: number
  readonly conditionalFormatCount: number
  readonly dataValidationCount?: number
  readonly warningCount: number
  readonly dimensions: readonly LargeSimpleXlsxSheetDimension[]
  readonly phaseTelemetry: readonly LargeSimpleXlsxImportPhaseTelemetry[]
}

export interface LargeSimpleXlsxSheetDimension {
  readonly sheetName: string
  readonly rowCount: number
  readonly columnCount: number
  readonly nonEmptyCellCount: number
  readonly usedRange: ImportedWorksheetCellScan['usedRange']
}

interface ParsedWorksheet {
  readonly sheet: WorkbookSnapshot['sheets'][number]
  readonly runtimeSheetCells: ImportedRuntimeSheetCells
  readonly preview: ReturnType<typeof createSheetPreview>
  readonly stats: {
    readonly cellCount: number
    readonly formulaCellCount: number
    readonly valueCellCount: number
    readonly tableCount: number
    readonly mergeCount: number
    readonly conditionalFormatCount: number
    readonly dataValidationCount: number
    readonly dimension: LargeSimpleXlsxSheetDimension
  }
}

interface ScannedWorksheet {
  readonly name: string
  readonly order: number
  readonly cellScan: ImportedWorksheetCellScan
  readonly worksheetXml: string | undefined
  readonly metadataScan: LargeSimpleWorksheetScannedMetadata | undefined
  readonly metadataInput: LargeSimpleSheetMetadataInput
}

type LargeSimpleSheetMetadataInput = Pick<
  SheetMetadataSnapshot,
  'conditionalFormatArtifacts' | 'conditionalFormats' | 'drawingArtifacts' | 'filters' | 'hyperlinks' | 'printerSettings' | 'printPageSetup'
> &
  Pick<SheetMetadataSnapshot, 'validations'>

const defaultLargeSimpleXlsxByteThreshold = 1_000_000
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const sharedStringsPath = 'xl/sharedStrings.xml'
const stylesPath = 'xl/styles.xml'
const unsupportedPackagePathPattern =
  /^xl\/(?:charts|chartSheets|comments|ctrlProps|externalLinks|model|pivotCache|pivotTables|threadedComments|vbaProject\.bin)/u
const maxPreservedBlankStyleCellCount = 100_000

export function tryImportLargeSimpleXlsx(
  source: LargeSimpleXlsxImportSource,
  fileName: string,
  zip: XlsxZipEntries,
  options: LargeSimpleXlsxImportOptions = {},
): LargeSimpleXlsxImportResult | null {
  if (source.byteLength < (options.minByteLength ?? defaultLargeSimpleXlsxByteThreshold)) {
    return null
  }
  const phaseRecorder = new LargeSimpleXlsxImportPhaseRecorder()
  const zipSetupStart = phaseRecorder.start()
  const packagePaths = Object.keys(zip).map(normalizeZipPath)
  if (packagePaths.some((path) => unsupportedPackagePathPattern.test(path))) {
    return null
  }

  const workbookXml = getZipText(zip, workbookPath)
  const workbookRelationshipsXml = getZipText(zip, workbookRelationshipsPath)
  if (!workbookXml || !workbookRelationshipsXml) {
    return null
  }

  const stringPool = new ImportedWorkbookStringPool()
  const workbookSheets = readWorkbookSheets(workbookXml, stringPool)
  const worksheetPathsByRelationshipId = readWorksheetPathsByRelationshipId(workbookRelationshipsXml)
  if (workbookSheets.length === 0 || worksheetPathsByRelationshipId.size === 0) {
    return null
  }
  const workbookDefinedNames = readWorkbookDefinedNames(
    workbookXml,
    workbookSheets.map((entry) => entry.name),
  )
  if (workbookDefinedNames.externalWorkbookReferenceSeen) {
    return null
  }

  const worksheetEntries = workbookSheets.flatMap((entry) => {
    const path = worksheetPathsByRelationshipId.get(entry.relationshipId)
    return path ? [{ name: entry.name, relationshipId: entry.relationshipId, path }] : []
  })
  if (worksheetEntries.length !== workbookSheets.length) {
    return null
  }
  const materializeCells = options.materializeCells !== false
  const materializeMetadata = options.materializeMetadata !== false
  const hasSharedStrings = packagePaths.includes(sharedStringsPath)
  const hasStyles = packagePaths.includes(stylesPath)
  let fallbackSharedStrings: readonly LargeSimpleSharedStringEntry[] | null | undefined = hasSharedStrings ? undefined : []
  delete zip[workbookPath]
  delete zip[workbookRelationshipsPath]
  const workbookName = stringPool.intern(normalizeWorkbookName(fileName))
  const warnings = workbookDefinedNames.ignoredCount > 0 ? ['Some defined names were ignored during XLSX import.'] : []
  const hasDrawingParts = packagePaths.some((path) => path.startsWith('xl/drawings/') || path.startsWith('xl/media/'))
  phaseRecorder.finish('zip-setup', zipSetupStart)
  const importedTables: WorkbookTableSnapshot[] = []
  const sheets: WorkbookSnapshot['sheets'] = []
  const runtimeSheetCells: ImportedRuntimeSheetCells[] = []
  const previewSheets: ParsedWorksheet['preview'][] = []
  const sheetStats: ParsedWorksheet['stats'][] = []
  const styleCatalog = new Map<string, CellStyleRecord>()
  const scannedWorksheets: (ScannedWorksheet | undefined)[] = []
  const referencedSharedStringIndexes = new Set<number>()
  const materializeSheetsImmediately =
    materializeCells && options.releaseZipSource !== true && !hasSharedStrings && !hasStyles && !hasDrawingParts
  const emptyStylesByIndex = new Map<number, Omit<CellStyleRecord, 'id'>>()
  const appendParsedWorksheet = (parsed: ParsedWorksheet): void => {
    sheets.push(parsed.sheet)
    runtimeSheetCells.push(parsed.runtimeSheetCells)
    previewSheets.push(parsed.preview)
    sheetStats.push(parsed.stats)
  }

  for (const [order, entry] of worksheetEntries.entries()) {
    const worksheetScanStart = phaseRecorder.start()
    let streamedWorksheetXml: string | undefined
    let streamedMetadataScan: LargeSimpleWorksheetScannedMetadata | undefined
    let retainedMetadataScan: LargeSimpleWorksheetScannedMetadata | undefined
    let cellScan: ImportedWorksheetCellScan | null = null
    if (!materializeCells && !materializeMetadata) {
      const headless = parseHeadlessLargeSimpleWorksheetFromChunks(
        (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
        order,
        { hasSharedStrings, sheetName: entry.name },
      )
      if (headless && (hasSharedStrings || headless.valueCellCount > 0)) {
        cellScan = importedWorksheetCellScanFromHeadless(headless)
        delete zip[entry.path]
      }
    } else {
      const streamed = parseLargeSimpleWorksheetCellsFromChunks(
        (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
        order,
        {
          hasSharedStrings,
          retainCells: materializeCells,
          sharedStrings: fallbackSharedStrings ?? [],
          deferSharedStrings: materializeCells && hasSharedStrings,
          retainMetadataXml: materializeMetadata,
          sheetName: entry.name,
          stringPool,
        },
      )
      if (streamed && (hasSharedStrings || streamed.cellScan.valueCellCount > 0)) {
        cellScan = streamed.cellScan
        streamedWorksheetXml = streamed.metadataXml
        streamedMetadataScan = streamed.metadata
        delete zip[entry.path]
      }
    }
    let worksheetBytes: Uint8Array | undefined
    if (!cellScan) {
      worksheetBytes = zip[entry.path]
      if (!worksheetBytes) {
        return null
      }
      delete zip[entry.path]
      if (!hasSharedStrings && !shouldUseSharedStringlessFastPathBytes(worksheetBytes)) {
        return null
      }
      if (hasUnsupportedLargeSimpleWorksheetTags(worksheetBytes)) {
        return null
      }
      if (hasSharedStrings && fallbackSharedStrings === undefined) {
        fallbackSharedStrings = readAllLargeSimpleSharedStrings(zip)
        if (fallbackSharedStrings === null) {
          return null
        }
      }
      cellScan = parseLargeSimpleWorksheetCells(worksheetBytes, fallbackSharedStrings ?? [], order, {
        retainCells: materializeCells,
        stringPool,
      })
    }
    if (!cellScan) {
      return null
    }
    retainedMetadataScan = streamedMetadataScan
    if (materializeCells && cellScan.blankStyleCellCount > 0 && cellScan.blankStyleCellCount <= maxPreservedBlankStyleCellCount) {
      return null
    }
    cellScan.arena.collectSharedStringIndexes(referencedSharedStringIndexes)
    phaseRecorder.finish('worksheet-scan', worksheetScanStart)
    const metadataParsingStart = phaseRecorder.start()
    let worksheetXml: string | undefined
    let metadataInput: LargeSimpleSheetMetadataInput = {}
    const needsWorksheetXml =
      materializeMetadata &&
      (streamedWorksheetXml !== undefined || (worksheetBytes ? needsLargeSimpleWorksheetMetadataXml(worksheetBytes) : false))
    if (needsWorksheetXml) {
      worksheetXml = streamedWorksheetXml ?? (worksheetBytes ? readLargeSimpleWorksheetMetadataXml(worksheetBytes) : undefined)
      if (!worksheetXml) {
        return null
      }
      const sheetTables = streamedMetadataScan?.tableRelationshipIds
        ? undefined
        : /<(?:[A-Za-z_][\w.-]*:)?tableParts\b/u.test(worksheetXml)
          ? readImportedSheetTablesFromWorksheetXml(zip, entry.name, entry.path, worksheetXml)
          : undefined
      if (sheetTables) {
        importedTables.push(...sheetTables)
      }
      const hasConditionalFormats = /<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b/u.test(worksheetXml)
      const conditionalFormats = hasConditionalFormats
        ? readImportedSheetConditionalFormatsFromWorksheetXml(zip, entry.name, worksheetXml)
        : undefined
      if (materializeCells) {
        const hyperlinks = streamedMetadataScan?.hyperlinks
          ? undefined
          : readLargeSimpleSheetHyperlinks(zip, entry.name, entry.path, worksheetXml)
        if (hyperlinks === null) {
          return null
        }
        const filters = streamedMetadataScan?.filters ? [] : readImportedSheetAutoFilters(entry.name, worksheetXml)
        const conditionalFormatArtifacts = hasConditionalFormats
          ? readImportedSheetConditionalFormatArtifactsFromWorksheetXml(worksheetXml)
          : undefined
        metadataInput = appendConditionalFormats(
          {
            ...(hyperlinks ? { hyperlinks } : {}),
            ...(filters.length > 0 ? { filters } : {}),
            ...(conditionalFormatArtifacts ? { conditionalFormatArtifacts } : {}),
          },
          conditionalFormats,
        )
      } else {
        metadataInput = appendConditionalFormats(metadataInput, conditionalFormats)
      }
    }
    if (streamedMetadataScan?.conditionalFormats && streamedMetadataScan.conditionalFormats.length > 0) {
      metadataInput = appendConditionalFormats(metadataInput, streamedMetadataScan.conditionalFormats)
    }
    if (materializeMetadata && streamedMetadataScan?.conditionalFormattingXml && streamedMetadataScan.conditionalFormattingXml.length > 0) {
      const conditionalFormats = readImportedSheetConditionalFormatsFromElementXml(
        zip,
        entry.name,
        streamedMetadataScan.conditionalFormattingXml,
      )
      const conditionalFormatArtifacts = materializeCells
        ? readImportedSheetConditionalFormatArtifactsFromElementXml(streamedMetadataScan.conditionalFormattingXml)
        : undefined
      metadataInput = appendConditionalFormats(
        {
          ...metadataInput,
          ...(conditionalFormatArtifacts ? { conditionalFormatArtifacts } : {}),
        },
        conditionalFormats,
      )
      retainedMetadataScan = withoutConditionalFormattingXml(streamedMetadataScan)
    }
    if (materializeMetadata && streamedMetadataScan?.tableRelationshipIds && streamedMetadataScan.tableRelationshipIds.length > 0) {
      const sheetTables = readImportedSheetTablesFromRelationshipIds(zip, entry.name, entry.path, streamedMetadataScan.tableRelationshipIds)
      if (sheetTables) {
        importedTables.push(...sheetTables)
      }
    }
    if (materializeCells && streamedMetadataScan?.dataValidations && streamedMetadataScan.dataValidations.length > 0) {
      metadataInput = { ...metadataInput, validations: [...streamedMetadataScan.dataValidations] }
    }
    if (materializeCells) {
      const printPageSetup =
        streamedMetadataScan?.printPageSetup ?? (worksheetXml ? readLargeSimpleSheetPrintPageSetup(worksheetXml) : undefined)
      const printMetadata = readLargeSimpleSheetPrintMetadata(zip, entry.path, printPageSetup)
      if (printMetadata === null) {
        return null
      }
      metadataInput = { ...metadataInput, ...printMetadata }
    }
    const streamedHyperlinks =
      materializeCells && streamedMetadataScan?.hyperlinks
        ? resolveLargeSimpleSheetHyperlinks(zip, entry.name, entry.path, streamedMetadataScan.hyperlinks)
        : undefined
    if (streamedHyperlinks === null) {
      return null
    }
    if (streamedHyperlinks) {
      metadataInput = { ...metadataInput, hyperlinks: streamedHyperlinks }
    }
    if (materializeCells && streamedMetadataScan?.filters && streamedMetadataScan.filters.length > 0) {
      metadataInput = { ...metadataInput, filters: [...streamedMetadataScan.filters] }
    }
    worksheetBytes = undefined
    if (materializeSheetsImmediately) {
      phaseRecorder.finish('metadata-parsing', metadataParsingStart)
      const snapshotMaterializationStart = phaseRecorder.start()
      appendParsedWorksheet(
        buildParsedWorksheet(entry.name, order, cellScan, worksheetXml, retainedMetadataScan, metadataInput, {
          materializeCells,
          releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
          styleCatalog,
          stylesByIndex: emptyStylesByIndex,
        }),
      )
      phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
      continue
    }
    scannedWorksheets.push({
      name: entry.name,
      order,
      cellScan,
      worksheetXml,
      metadataScan: retainedMetadataScan,
      metadataInput,
    })
    phaseRecorder.finish('metadata-parsing', metadataParsingStart)
  }
  const sharedStringResolutionStart = phaseRecorder.start()
  let sharedStrings: readonly LargeSimpleSharedStringEntry[] = fallbackSharedStrings ?? []
  if (materializeCells && hasSharedStrings && referencedSharedStringIndexes.size > 0) {
    const referencedSharedStrings = fallbackSharedStrings ?? readReferencedLargeSimpleSharedStrings(zip, referencedSharedStringIndexes)
    if (referencedSharedStrings === null) {
      return null
    }
    sharedStrings = referencedSharedStrings
  }
  delete zip[sharedStringsPath]
  referencedSharedStringIndexes.clear()
  fallbackSharedStrings = null
  phaseRecorder.finish('shared-string-resolution', sharedStringResolutionStart)
  const styleParsingStart = phaseRecorder.start()
  const requiredStyleIndexes = new Set<number>()
  for (const scanned of scannedWorksheets) {
    if (!scanned) {
      continue
    }
    scanned.cellScan.styleIndexes.collectRequiredStyleIndexes(requiredStyleIndexes)
  }
  const stylesByIndex =
    materializeCells && hasStyles
      ? readLargeSimpleWorkbookStylesFromChunks(
          (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, stylesPath, onChunk),
          requiredStyleIndexes,
        )
      : new Map()
  if (stylesByIndex === null) {
    return null
  }
  requiredStyleIndexes.clear()
  delete zip[stylesPath]
  phaseRecorder.finish('style-parsing', styleParsingStart)
  const importedDrawingArtifacts =
    materializeCells && hasDrawingParts
      ? readImportedWorkbookDrawingArtifactsFromWorksheetRelationships(
          zip,
          scannedWorksheets.flatMap((scanned) => {
            if (!scanned) {
              return []
            }
            const drawingRelationshipId = drawingRelationshipIdForScannedWorksheet(scanned)
            return [
              {
                name: scanned.name,
                path: worksheetEntries[scanned.order]?.path ?? '',
                ...(drawingRelationshipId ? { drawingRelationshipId } : {}),
              },
            ]
          }),
        )
      : null
  if (options.releaseZipSource === true) {
    const zipSourceReleaseStart = phaseRecorder.start()
    const zipSourceBytesBeforeRelease = readLazyXlsxZipSourceByteLength(zip)
    releaseLazyXlsxZipSource(zip)
    const ownedSourceReleaseEvidence = options.releaseOwnedSourceBytes?.()
    phaseRecorder.finish('zip-source-release', zipSourceReleaseStart, {
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesBeforeRelease } : {}),
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesAfterRelease: readLazyXlsxZipSourceByteLength(zip) ?? 0 } : {}),
      ...ownedSourceReleaseEvidence,
    })
  }
  let remainingSharedStringSheets =
    materializeCells && hasSharedStrings
      ? scannedWorksheets.reduce((sum, scanned) => sum + ((scanned?.cellScan.arena.sharedStringRefCount ?? 0) > 0 ? 1 : 0), 0)
      : 0
  for (const [index, scanned] of scannedWorksheets.entries()) {
    if (!scanned) {
      continue
    }
    const snapshotMaterializationStart = phaseRecorder.start()
    const sheetHasSharedStrings = materializeCells && hasSharedStrings && scanned.cellScan.arena.sharedStringRefCount > 0
    const resolvedRichTextCells = sheetHasSharedStrings ? scanned.cellScan.arena.resolveSharedStrings(sharedStrings) : []
    if (resolvedRichTextCells === null) {
      return null
    }
    if (resolvedRichTextCells.length > 0) {
      scanned.cellScan.richTextCells.push(...resolvedRichTextCells)
    }
    if (sheetHasSharedStrings) {
      remainingSharedStringSheets -= 1
      if (remainingSharedStringSheets === 0) {
        sharedStrings = []
      }
    }
    const drawingArtifacts = importedDrawingArtifacts?.sheetArtifactsByName.get(scanned.name)
    const parsed = buildParsedWorksheet(
      scanned.name,
      scanned.order,
      scanned.cellScan,
      scanned.worksheetXml,
      scanned.metadataScan,
      {
        ...scanned.metadataInput,
        ...(drawingArtifacts ? { drawingArtifacts } : {}),
      },
      {
        materializeCells,
        releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
        styleCatalog,
        stylesByIndex,
      },
    )
    appendParsedWorksheet(parsed)
    scannedWorksheets[index] = undefined
    phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
  }
  sharedStrings = []
  stringPool.release()
  const sortedImportedTables =
    importedTables.length > 0 ? importedTables.toSorted((left, right) => left.name.localeCompare(right.name)) : undefined
  const stats: LargeSimpleXlsxImportStats = {
    sheetCount: sheets.length,
    cellCount: sheetStats.reduce((sum, entry) => sum + entry.cellCount, 0),
    formulaCellCount: sheetStats.reduce((sum, entry) => sum + entry.formulaCellCount, 0),
    valueCellCount: sheetStats.reduce((sum, entry) => sum + entry.valueCellCount, 0),
    definedNameCount: workbookDefinedNames.definedNames?.length ?? 0,
    tableCount: sortedImportedTables?.length ?? sheetStats.reduce((sum, entry) => sum + entry.tableCount, 0),
    mergeCount: sheetStats.reduce((sum, entry) => sum + entry.mergeCount, 0),
    conditionalFormatCount: sheetStats.reduce((sum, entry) => sum + entry.conditionalFormatCount, 0),
    dataValidationCount: sheetStats.reduce((sum, entry) => sum + entry.dataValidationCount, 0),
    warningCount: warnings.length,
    dimensions: sheetStats.map((entry) => entry.dimension),
    phaseTelemetry: phaseRecorder.entries(),
  }

  return {
    snapshot: attachImportedRuntimeCoordinates(
      {
        version: 1,
        workbook: {
          name: workbookName,
          ...(workbookDefinedNames.definedNames || importedDrawingArtifacts?.artifacts || sortedImportedTables || styleCatalog.size > 0
            ? {
                metadata: {
                  ...(workbookDefinedNames.definedNames ? { definedNames: workbookDefinedNames.definedNames } : {}),
                  ...(importedDrawingArtifacts?.artifacts ? { drawingArtifacts: importedDrawingArtifacts.artifacts } : {}),
                  ...(sortedImportedTables ? { tables: sortedImportedTables } : {}),
                  ...(styleCatalog.size > 0 ? { styles: [...styleCatalog.values()] } : {}),
                },
              }
            : {}),
        },
        sheets,
      },
      runtimeSheetCells,
    ),
    workbookName,
    sheetNames: workbookSheets.map((entry) => entry.name),
    warnings,
    preview: createWorkbookPreview({
      contentType: XLSX_CONTENT_TYPE,
      fileName,
      fileSizeBytes: source.byteLength,
      workbookName,
      sheets: previewSheets,
      warnings,
    }),
    stats,
  }
}

function withoutConditionalFormattingXml(
  metadata: LargeSimpleWorksheetScannedMetadata | undefined,
): LargeSimpleWorksheetScannedMetadata | undefined {
  if (!metadata?.conditionalFormattingXml) {
    return metadata
  }
  const { conditionalFormattingXml: _released, ...retained } = metadata
  return Object.keys(retained).length > 0 ? retained : undefined
}

function appendConditionalFormats(
  input: LargeSimpleSheetMetadataInput,
  conditionalFormats: readonly WorkbookConditionalFormatSnapshot[] | undefined,
): LargeSimpleSheetMetadataInput {
  if (!conditionalFormats || conditionalFormats.length === 0) {
    return input
  }
  return {
    ...input,
    conditionalFormats: [...(input.conditionalFormats ?? []), ...conditionalFormats],
  }
}

function normalizeConditionalFormatIds(
  sheetName: string,
  conditionalFormats: readonly WorkbookConditionalFormatSnapshot[] | undefined,
): SheetMetadataSnapshot['conditionalFormats'] | undefined {
  if (!conditionalFormats || conditionalFormats.length === 0) {
    return undefined
  }
  return conditionalFormats.map((format, index) => ({
    ...format,
    id: `xlsx-cf:${sheetName}:${format.range.startAddress}:${format.range.endAddress}:${String(index + 1)}`,
  }))
}

function buildParsedWorksheet(
  sheetName: string,
  order: number,
  cellScan: ImportedWorksheetCellScan,
  worksheetXml: string | undefined,
  metadataScan: LargeSimpleWorksheetScannedMetadata | undefined,
  input: LargeSimpleSheetMetadataInput = {},
  options: {
    readonly materializeCells: boolean
    readonly releaseArenaAfterMaterialization?: boolean
    readonly styleCatalog?: Map<string, CellStyleRecord>
    readonly stylesByIndex?: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>
  } = { materializeCells: true },
): ParsedWorksheet {
  const merges =
    metadataScan?.merges?.map((range) => ({ sheetName, ...range })) ??
    (worksheetXml ? readLargeSimpleMergeRanges(sheetName, worksheetXml) : [])
  const mergeCount = worksheetXml ? merges.length : (cellScan.mergeCount ?? 0)
  const columns = metadataScan?.columns ?? (worksheetXml ? readLargeSimpleColumnMetadata(worksheetXml) : { entries: [], metadata: [] })
  const rows = metadataScan?.rows ?? (worksheetXml ? readLargeSimpleRowMetadata(worksheetXml) : { entries: [], metadata: [] })
  const sheetFormatPr = metadataScan?.sheetFormatPr ?? (worksheetXml ? readLargeSimpleSheetFormatPr(worksheetXml) : undefined)
  const conditionalFormatCount =
    input.conditionalFormats?.length ??
    (worksheetXml ? readConditionalFormattingBlockCount(worksheetXml) : (cellScan.conditionalFormatCount ?? 0))
  const conditionalFormats = normalizeConditionalFormatIds(sheetName, input.conditionalFormats)
  const styleRanges =
    options.materializeCells && options.styleCatalog && options.stylesByIndex
      ? buildLargeSimpleStyleRanges(sheetName, cellScan, options.stylesByIndex, options.styleCatalog)
      : []
  const metadata: SheetMetadataSnapshot = {
    ...(columns.entries.length > 0 ? { columns: columns.entries } : {}),
    ...(rows.entries.length > 0 ? { rows: rows.entries } : {}),
    ...(columns.metadata.length > 0 ? { columnMetadata: columns.metadata } : {}),
    ...(rows.metadata.length > 0 ? { rowMetadata: rows.metadata } : {}),
    ...(sheetFormatPr ? { sheetFormatPr } : {}),
    ...(styleRanges.length > 0 ? { styleRanges } : {}),
    ...(merges.length > 0 ? { merges } : {}),
    ...(input.drawingArtifacts ? { drawingArtifacts: input.drawingArtifacts } : {}),
    ...(input.filters ? { filters: input.filters } : {}),
    ...(input.hyperlinks ? { hyperlinks: input.hyperlinks } : {}),
    ...(input.validations ? { validations: input.validations } : {}),
    ...(conditionalFormats ? { conditionalFormats } : {}),
    ...(input.conditionalFormatArtifacts ? { conditionalFormatArtifacts: input.conditionalFormatArtifacts } : {}),
    ...(input.printerSettings ? { printerSettings: input.printerSettings } : {}),
    ...(input.printPageSetup ? { printPageSetup: input.printPageSetup } : {}),
    ...(cellScan.richTextCells.length > 0 ? { richTextArtifacts: { cells: cellScan.richTextCells } } : {}),
  }
  const preview = createSheetPreview({
    name: sheetName,
    rowCount: cellScan.rowCount,
    columnCount: cellScan.columnCount,
    nonEmptyCellCount: cellScan.cellCount,
    readCellText: (row, column) => cellScan.arena.readPreviewText(row, column),
  })
  const cells = options.materializeCells ? cellScan.arena.materializeSheetCells(cellScan.sheetIndex) : []
  const runtimeSheetCells = createImportedRuntimeSheetCells({
    sheetName,
    coords: cells.flatMap((cell) =>
      typeof cell.row === 'number' && typeof cell.col === 'number' ? [{ row: cell.row, col: cell.col }] : [],
    ),
    width: cellScan.columnCount,
    height: cellScan.rowCount,
  })
  if (options.releaseArenaAfterMaterialization === true) {
    cellScan.arena.release()
    cellScan.styleIndexes.release()
  }
  const sheet: WorkbookSnapshot['sheets'][number] = {
    id: order + 1,
    name: sheetName,
    order,
    ...(Object.keys(metadata).length > 0 ? { metadata } : {}),
    cells,
  }
  return {
    sheet,
    runtimeSheetCells,
    preview,
    stats: {
      cellCount: cellScan.cellCount,
      formulaCellCount: cellScan.formulaCellCount,
      valueCellCount: cellScan.valueCellCount,
      tableCount: cellScan.tableCount ?? 0,
      mergeCount,
      conditionalFormatCount,
      dataValidationCount: input.validations?.length ?? cellScan.dataValidationCount ?? 0,
      dimension: {
        sheetName,
        rowCount: cellScan.rowCount,
        columnCount: cellScan.columnCount,
        nonEmptyCellCount: cellScan.cellCount,
        usedRange: cellScan.usedRange,
      },
    },
  }
}

function drawingRelationshipIdForScannedWorksheet(scanned: ScannedWorksheet): string | undefined {
  return (
    scanned.metadataScan?.drawingRelationshipId ??
    (scanned.worksheetXml ? readLargeSimpleDrawingRelationshipId(scanned.worksheetXml) : undefined)
  )
}

function readConditionalFormattingBlockCount(worksheetXml: string): number {
  return [...worksheetXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b/gu)].length
}
