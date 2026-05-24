import type { CellStyleRecord, WorkbookRichTextCellSnapshot, WorkbookSnapshot, WorkbookTableSnapshot } from '@bilig/protocol'
import { attachImportedRuntimeImage } from './import-runtime-image.js'
import { normalizeWorkbookName } from './workbook-import-helpers.js'
import { XLSX_CONTENT_TYPE } from './workbook-import-content-types.js'
import { createWorkbookPreview } from './workbook-import-preview.js'
import { applyImportedAutoFilterRowVisibility } from './xlsx-autofilter-row-visibility.js'
import {
  readImportedSheetConditionalFormatArtifactsFromElementXml,
  readImportedSheetConditionalFormatArtifactsFromWorksheetXml,
  readImportedSheetConditionalFormatsFromElementXml,
  readImportedSheetConditionalFormatsFromWorksheetXml,
} from './xlsx-conditional-formats.js'
import { readImportedWorkbookCellMetadataPart } from './xlsx-cell-metadata.js'
import { legacyCommentThreadSignature, readImportedWorkbookLegacyCommentVmlFromSheetSources } from './xlsx-comment-vml.js'
import { readImportedWorkbookControlArtifactsFromSheetSources } from './xlsx-control-artifacts.js'
import { readImportedWorkbookDataModelArtifacts } from './xlsx-data-model-artifacts.js'
import { readImportedWorkbookDrawingArtifactsFromWorksheetRelationships } from './xlsx-drawing-artifacts.js'
import { readImportedWorkbookExternalLinkArtifacts } from './xlsx-external-link-artifacts.js'
import { readImportedSheetAutoFilters } from './xlsx-filters.js'
import { readImportedWorkbookChartDrawingArtifacts } from './xlsx-import-chart-drawing-artifacts.js'
import { externalPivotCachesWarning, unsupportedCellStylesWarning } from './xlsx-import-warnings.js'
import { readWorkbookDefinedNames } from './xlsx-large-simple-defined-names.js'
import { readLargeSimpleSheetHyperlinks, resolveLargeSimpleSheetHyperlinks } from './xlsx-large-simple-hyperlinks.js'
import { LargeSimpleXlsxImportPhaseRecorder } from './xlsx-large-simple-import-telemetry.js'
import { appendLargeSimpleConditionalFormats } from './xlsx-large-simple-conditional-format-helpers.js'
import { internLargeSimpleWorksheetMetadata } from './xlsx-large-simple-metadata-interning.js'
import { prepareLargeSimplePackageArtifactsForZipRelease } from './xlsx-large-simple-package-artifact-release.js'
import { readLargeSimpleSheetPrintMetadata, readLargeSimpleSheetPrintPageSetup } from './xlsx-large-simple-printer-settings.js'
import {
  readAllLargeSimpleSharedStrings,
  readAllLargeSimpleSharedStringsStreamed,
  readReferencedLargeSimpleSharedStrings,
} from './xlsx-large-simple-referenced-shared-strings.js'
import {
  emptyLargeSimpleSharedStringIndexes,
  LargeSimpleSharedStringIndexCollector,
  type LargeSimpleSharedStringIndexSet,
} from './xlsx-large-simple-shared-string-indexes.js'
import {
  collectReferencedLargeSimpleRichSharedStringIndexes,
  createLargeSimpleSharedStringSubset,
  type LargeSimpleSharedStrings,
} from './xlsx-large-simple-shared-strings.js'
import { shouldUseSharedStringlessFastPathBytes } from './xlsx-large-simple-shared-stringless-fast-path.js'
import { readLargeSimpleWorkbookStyleArtifactsFromChunks } from './xlsx-large-simple-styles.js'
import {
  maxPreallocatedWorksheetCells,
  prepareLargeSimpleStyleIndexForWorksheet,
  releaseLargeSimpleStyleIndexes,
  shouldDeferLargeSimpleStyleCoordinates,
} from './xlsx-large-simple-style-coordinate-rescan.js'
import { collectLargeSimpleImportGarbage } from './xlsx-large-simple-garbage.js'
import { forEachLargeSimpleInflatedZipEntryChunk } from './xlsx-large-simple-stream-garbage.js'
import {
  drawingRelationshipIdForScannedWorksheet,
  sheetPivotArtifactsWithStreamedDefinitions,
} from './xlsx-large-simple-materialization-helpers.js'
import {
  buildParsedWorksheet,
  lazySheetCellMaterializationNumberFormatThreshold,
  lazySheetCellMaterializationThreshold,
} from './xlsx-large-simple-build-parsed-worksheet.js'
import { mergeWorkbookRichTextCells } from './xlsx-large-simple-lazy-rich-text-cells.js'
import { buildLargeSimpleRuntimeSheetCells } from './xlsx-large-simple-runtime-sheet-cells.js'
import { hasExternalLargeSimplePivotCaches } from './xlsx-large-simple-pivot-warnings.js'
import { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import { readWorkbookSheets, readWorksheetPathsByRelationshipId } from './xlsx-large-simple-workbook-metadata.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import {
  largeSimpleControlArtifactSheetSources,
  largeSimpleLegacyCommentVmlSheetSources,
  largeSimpleSlicerConnectionRelationshipSheetNames,
  largeSimpleSlicerConnectionSheetSources,
} from './xlsx-large-simple-package-artifact-sources.js'
import { parseHeadlessLargeSimpleWorksheetFromChunks } from './xlsx-large-simple-headless-worksheet-scanner.js'
import { importedWorksheetCellScanFromHeadless } from './xlsx-large-simple-headless-cell-scan.js'
import {
  hasUnsupportedLargeSimpleWorksheetTags,
  needsLargeSimpleWorksheetMetadataXml,
  readLargeSimpleWorksheetMetadataXml,
  parseLargeSimpleWorksheetCells,
} from './xlsx-large-simple-worksheet-scanner.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from './xlsx-large-simple-worksheet-stream-scanner.js'
import {
  withoutLargeSimpleConditionalFormattingXml,
  type LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'
import { readImportedPivotArtifacts } from './xlsx-pivot-artifacts.js'
import { readImportedWorkbookSlicerConnectionArtifactsFromSheets } from './xlsx-slicer-connection-artifacts.js'
import { readImportedSheetTablesFromRelationshipIds, readImportedSheetTablesFromWorksheetXml } from './xlsx-tables.js'
import { readImportedWorkbookDocumentPropertiesArtifacts, readImportedWorkbookProperties } from './xlsx-workbook-properties.js'
import {
  getZipText,
  normalizeZipPath,
  readLazyXlsxZipSourceByteLength,
  readXlsxZipEntryUncompressedSize,
  releaseLazyXlsxZipSource,
  replaceLazyXlsxZipSource,
  type XlsxZipEntries,
} from './xlsx-zip.js'
import type {
  LargeSimpleSheetMetadataInput,
  LargeSimpleXlsxImportOptions,
  LargeSimpleXlsxImportResult,
  LargeSimpleXlsxImportSource,
  LargeSimpleXlsxImportStats,
  ParsedWorksheet,
  ScannedWorksheet,
} from './xlsx-large-simple-import-types.js'

export type {
  LargeSimpleXlsxImportOptions,
  LargeSimpleXlsxImportResult,
  LargeSimpleXlsxImportSource,
  LargeSimpleXlsxImportStats,
  LargeSimpleXlsxOwnedSourceReleaseEvidence,
  LargeSimpleXlsxSheetDimension,
} from './xlsx-large-simple-import-types.js'

const defaultLargeSimpleXlsxByteThreshold = 1_000_000
const maxMultiSheetDimensionCellPreallocation = 1_000_000
const eagerSharedStringsXmlByteThreshold = 1_000_000
const repeatedInlineStringDedupeMaxEntries = 65_536
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const sharedStringsPath = 'xl/sharedStrings.xml'
const stylesPath = 'xl/styles.xml'
const unsupportedPackagePathPattern = /^xl\/(?:ctrlProps|threadedComments|vbaProject\.bin)/u
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
  const hasCalcChain = packagePaths.includes('xl/calcChain.xml')
  const hasDrawingParts = packagePaths.some((path) => path.startsWith('xl/drawings/') || path.startsWith('xl/media/'))
  const hasChartParts = packagePaths.some((path) => path.startsWith('xl/charts/') || path.startsWith('xl/chartSheets/'))
  const hasPivotParts = packagePaths.some((path) => path.startsWith('xl/pivotTables/') || path.startsWith('xl/pivotCache/'))
  const hasExternalLinkParts = packagePaths.some((path) => path.startsWith('xl/externalLinks/'))
  const hasLegacyCommentParts = packagePaths.some((path) => path.startsWith('xl/comments') || path.endsWith('.vml'))
  const hasDataModelParts = packagePaths.some(
    (path) => path.startsWith('xl/model/') || path.startsWith('xl/customData/') || path.startsWith('customXml/'),
  )
  const hasSlicerConnectionParts = packagePaths.some(
    (path) =>
      path === 'xl/connections.xml' ||
      path.startsWith('xl/queryTables/') ||
      path.startsWith('xl/slicerCaches/') ||
      path.startsWith('xl/slicers/') ||
      /^xl\/tables\/_rels\/table[1-9][0-9]*\.xml\.rels$/u.test(path),
  )
  phaseRecorder.finish('zip-setup', zipSetupStart)
  let ownedSourceReleasedForReplacement = false
  if (options.releaseZipSource === true && options.replacementZipSource) {
    const zipSourceReleaseStart = phaseRecorder.start()
    const zipSourceBytesBeforeRelease = readLazyXlsxZipSourceByteLength(zip)
    const zipSourceReplaced = replaceLazyXlsxZipSource(zip, options.replacementZipSource)
    const ownedSourceReleaseEvidence = zipSourceReplaced ? options.releaseOwnedSourceBytes?.() : undefined
    ownedSourceReleasedForReplacement = Boolean(ownedSourceReleaseEvidence)
    phaseRecorder.finish('zip-source-release', zipSourceReleaseStart, {
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesBeforeRelease } : {}),
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesAfterRelease: readLazyXlsxZipSourceByteLength(zip) ?? 0 } : {}),
      ...ownedSourceReleaseEvidence,
    })
  }
  const importedExternalLinkArtifacts =
    materializeCells && hasExternalLinkParts ? readImportedWorkbookExternalLinkArtifacts(zip) : undefined
  const importedDataModelArtifacts = materializeCells && hasDataModelParts ? readImportedWorkbookDataModelArtifacts(zip) : undefined
  const importedWorkbookProperties = materializeCells ? readImportedWorkbookProperties(zip) : undefined
  const importedWorkbookDocumentProperties = materializeCells ? readImportedWorkbookDocumentPropertiesArtifacts(zip) : undefined
  const importedWorkbookCellMetadata = materializeCells ? readImportedWorkbookCellMetadataPart(zip) : undefined
  const importedPivotArtifacts =
    materializeCells && hasPivotParts
      ? readImportedPivotArtifacts(
          zip,
          workbookSheets.map((entry) => entry.name),
          { readWorksheetPivotTableDefinitionsXml: false },
        )
      : null
  const importedChartDrawingArtifacts =
    materializeCells && hasChartParts
      ? readImportedWorkbookChartDrawingArtifacts(
          zip,
          workbookSheets.map((entry) => entry.name),
        )
      : null
  const deduplicateInlineStrings = hasSharedStrings ? true : 'bounded'
  const inlineStringDedupeMaxEntries = hasSharedStrings ? undefined : repeatedInlineStringDedupeMaxEntries
  const deduplicateFormulaStrings = 'bounded'
  let fallbackSharedStrings: LargeSimpleSharedStrings | null | undefined = hasSharedStrings ? undefined : []
  if (
    materializeCells &&
    hasSharedStrings &&
    !hasCalcChain &&
    !hasPivotParts &&
    (readXlsxZipEntryUncompressedSize(zip, sharedStringsPath) ?? Number.POSITIVE_INFINITY) <= eagerSharedStringsXmlByteThreshold
  ) {
    const sharedStringResolutionStart = phaseRecorder.start()
    fallbackSharedStrings = readAllLargeSimpleSharedStringsStreamed(zip, {
      deduplicateText: 'bounded',
      stringPool,
    })
    if (fallbackSharedStrings === null) {
      return null
    }
    phaseRecorder.finish('shared-string-resolution', sharedStringResolutionStart)
  }
  delete zip[workbookPath]
  delete zip[workbookRelationshipsPath]
  const workbookName = stringPool.intern(normalizeWorkbookName(fileName))
  const warnings = workbookDefinedNames.ignoredCount > 0 ? ['Some defined names were ignored during XLSX import.'] : []
  if (hasExternalLargeSimplePivotCaches(zip)) {
    warnings.push(externalPivotCachesWarning)
  }
  const importedTables: WorkbookTableSnapshot[] = []
  const sheets: WorkbookSnapshot['sheets'] = []
  const previewSheets: ParsedWorksheet['preview'][] = []
  const sheetStats: ParsedWorksheet['stats'][] = []
  const styleCatalog = new Map<string, CellStyleRecord>()
  let materializationStringPool: ImportedWorkbookStringPool | undefined = stringPool
  const scannedWorksheets: (ScannedWorksheet | undefined)[] = []
  const slicerConnectionRelationshipSheetNames = hasSlicerConnectionParts
    ? largeSimpleSlicerConnectionRelationshipSheetNames(zip, worksheetEntries)
    : new Set<string>()
  const referencedSharedStringIndexes = new LargeSimpleSharedStringIndexCollector()
  const allowPreReleaseSheetFinalization =
    materializeCells &&
    worksheetEntries.length > 1 &&
    options.allowPreReleaseSheetFinalization === true &&
    (options.releaseOwnedSourceBytes === undefined || options.allowPreReleaseSheetFinalizationWithOwnedSourceRelease === true)
  const sheetHasRelationshipBackedArtifacts = (
    sheetName: string,
    metadataScan: LargeSimpleWorksheetScannedMetadata | undefined,
    worksheetXml: string | undefined,
  ): boolean =>
    Boolean(
      importedChartDrawingArtifacts?.drawingArtifacts.sheetArtifactsByName.has(sheetName) ||
      importedPivotArtifacts?.sheetArtifactsByName.has(sheetName) ||
      drawingRelationshipIdForScannedWorksheet({ metadataScan, worksheetXml }) ||
      metadataScan?.controlArtifacts ||
      metadataScan?.legacyDrawingRelationshipId ||
      metadataScan?.pivotTableDefinitionsXml ||
      metadataScan?.sheetSlicerListExtXml ||
      slicerConnectionRelationshipSheetNames.has(sheetName),
    )
  const canFinalizeSheetBeforeStyleParsing = (
    sheetName: string,
    cellScan: ImportedWorksheetCellScan,
    metadataScan: LargeSimpleWorksheetScannedMetadata | undefined,
    worksheetXml: string | undefined,
  ): boolean =>
    materializeCells &&
    !sheetHasRelationshipBackedArtifacts(sheetName, metadataScan, worksheetXml) &&
    cellScan.styleIndexes.count === 0 &&
    !cellScan.styleIndexes.hasRequiredStyleIndexes &&
    (options.releaseZipSource !== true || allowPreReleaseSheetFinalization)
  const emptyStylesByIndex = new Map<number, Omit<CellStyleRecord, 'id'>>()
  const appendParsedWorksheet = (parsed: ParsedWorksheet): void => {
    sheets[parsed.sheet.order] = parsed.sheet
    previewSheets[parsed.sheet.order] = parsed.preview
    sheetStats[parsed.sheet.order] = parsed.stats
  }
  const retainUnresolvedSharedStringReferences = (
    scanned: ScannedWorksheet,
    fallback: LargeSimpleSharedStrings,
  ): WorkbookRichTextCellSnapshot[] | null => {
    if (!materializeCells || !hasSharedStrings || scanned.hasUnresolvedSharedStringReferences !== true) {
      return []
    }
    const retainedSharedStrings = scanned.sharedStrings ?? fallback
    if (scanned.hasRichSharedStringReferences === false) {
      return scanned.cellScan.arena.retainPlainSharedStringReferences(retainedSharedStrings) ? [] : null
    }
    return scanned.cellScan.arena.retainSharedStringReferences(retainedSharedStrings)
  }

  for (const [order, entry] of worksheetEntries.entries()) {
    const worksheetScanStart = phaseRecorder.start()
    let streamedWorksheetXml: string | undefined
    let streamedMetadataScan: LargeSimpleWorksheetScannedMetadata | undefined
    let retainedMetadataScan: LargeSimpleWorksheetScannedMetadata | undefined
    let cellScan: ImportedWorksheetCellScan | null = null
    if (!materializeCells && !materializeMetadata) {
      const headless = parseHeadlessLargeSimpleWorksheetFromChunks(
        (onChunk) => forEachLargeSimpleInflatedZipEntryChunk(zip, entry.path, onChunk),
        order,
        { hasSharedStrings },
      )
      if (!headless) {
        return null
      }
      if (hasSharedStrings || headless.valueCellCount > 0) {
        cellScan = importedWorksheetCellScanFromHeadless(headless)
        delete zip[entry.path]
      }
    } else {
      const deferStyleCoordinates = shouldDeferLargeSimpleStyleCoordinates(zip, entry.path, { materializeCells, hasStyles })
      const streamed = parseLargeSimpleWorksheetCellsFromChunks(
        (onChunk) => forEachLargeSimpleInflatedZipEntryChunk(zip, entry.path, onChunk),
        order,
        {
          hasSharedStrings,
          retainCells: materializeCells,
          retainStyleIndexes: materializeCells && hasStyles,
          retainStyleCoordinates: materializeCells && hasStyles && !deferStyleCoordinates,
          sharedStrings: fallbackSharedStrings ?? [],
          deferSharedStrings: materializeCells && hasSharedStrings && fallbackSharedStrings === undefined,
          retainMetadataXml: materializeMetadata,
          sheetName: entry.name,
          stringPool,
          deduplicateStrings: deduplicateInlineStrings,
          deduplicateFormulas: deduplicateFormulaStrings,
          ...(inlineStringDedupeMaxEntries === undefined ? {} : { dedupeMaxEntries: inlineStringDedupeMaxEntries }),
          ...(options.allowUnsupportedFormulaText === undefined
            ? {}
            : { allowUnsupportedFormulaText: options.allowUnsupportedFormulaText }),
          ...(options.allowUnsupportedCellMetadata === undefined
            ? {}
            : { allowUnsupportedCellMetadata: options.allowUnsupportedCellMetadata }),
          maxDimensionCellPreallocation:
            worksheetEntries.length === 1
              ? maxPreallocatedWorksheetCells(zip, entry.path)
              : Math.min(maxPreallocatedWorksheetCells(zip, entry.path), maxMultiSheetDimensionCellPreallocation),
        },
      )
      if (!streamed) {
        return null
      }
      if (hasSharedStrings || streamed.cellScan.valueCellCount > 0) {
        cellScan = streamed.cellScan
        streamedWorksheetXml = streamed.metadataXml
        streamedMetadataScan = internLargeSimpleWorksheetMetadata(streamed.metadata, stringPool)
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
        fallbackSharedStrings = readAllLargeSimpleSharedStrings(zip, {
          deduplicateText: 'bounded',
          stringPool,
        })
        if (fallbackSharedStrings === null) {
          return null
        }
      }
      cellScan = parseLargeSimpleWorksheetCells(worksheetBytes, fallbackSharedStrings ?? [], order, {
        retainCells: materializeCells,
        stringPool,
        deduplicateStrings: deduplicateInlineStrings,
        deduplicateFormulas: deduplicateFormulaStrings,
        ...(inlineStringDedupeMaxEntries === undefined ? {} : { dedupeMaxEntries: inlineStringDedupeMaxEntries }),
        ...(options.allowUnsupportedFormulaText === undefined ? {} : { allowUnsupportedFormulaText: options.allowUnsupportedFormulaText }),
      })
    }
    if (!cellScan) {
      return null
    }
    retainedMetadataScan = streamedMetadataScan
    let sharedStringIndexes: LargeSimpleSharedStringIndexSet = emptyLargeSimpleSharedStringIndexes
    if (hasSharedStrings) {
      const sharedStringIndexCollector = new LargeSimpleSharedStringIndexCollector()
      cellScan.arena.collectSharedStringIndexes(sharedStringIndexCollector)
      sharedStringIndexes = sharedStringIndexCollector.finalize()
      referencedSharedStringIndexes.addAll(sharedStringIndexes)
    }
    phaseRecorder.finish('worksheet-scan', worksheetScanStart)
    collectLargeSimpleImportGarbage()
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
        metadataInput = appendLargeSimpleConditionalFormats(
          {
            ...(hyperlinks ? { hyperlinks } : {}),
            ...(filters.length > 0 ? { filters } : {}),
            ...(conditionalFormatArtifacts ? { conditionalFormatArtifacts } : {}),
          },
          conditionalFormats,
        )
      } else {
        metadataInput = appendLargeSimpleConditionalFormats(metadataInput, conditionalFormats)
      }
    }
    if (streamedMetadataScan?.conditionalFormats && streamedMetadataScan.conditionalFormats.length > 0) {
      metadataInput = appendLargeSimpleConditionalFormats(metadataInput, streamedMetadataScan.conditionalFormats)
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
      metadataInput = appendLargeSimpleConditionalFormats(
        {
          ...metadataInput,
          ...(conditionalFormatArtifacts ? { conditionalFormatArtifacts } : {}),
        },
        conditionalFormats,
      )
      retainedMetadataScan = withoutLargeSimpleConditionalFormattingXml(streamedMetadataScan)
    }
    if (materializeMetadata && streamedMetadataScan?.dataValidations && streamedMetadataScan.dataValidations.length > 0) {
      metadataInput = {
        ...metadataInput,
        validations: [...(metadataInput.validations ?? []), ...streamedMetadataScan.dataValidations],
      }
    }
    if (materializeMetadata && streamedMetadataScan?.tableRelationshipIds && streamedMetadataScan.tableRelationshipIds.length > 0) {
      const sheetTables = readImportedSheetTablesFromRelationshipIds(zip, entry.name, entry.path, streamedMetadataScan.tableRelationshipIds)
      if (sheetTables) {
        importedTables.push(...sheetTables)
      }
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
    if (
      (!hasSharedStrings || fallbackSharedStrings !== undefined) &&
      canFinalizeSheetBeforeStyleParsing(entry.name, cellScan, retainedMetadataScan, worksheetXml)
    ) {
      phaseRecorder.finish('metadata-parsing', metadataParsingStart)
      const snapshotMaterializationStart = phaseRecorder.start()
      appendParsedWorksheet(
        buildParsedWorksheet(entry.name, order, cellScan, worksheetXml, retainedMetadataScan, metadataInput, {
          materializeCells,
          releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
          styleCatalog,
          stylesByIndex: emptyStylesByIndex,
          ...(materializationStringPool ? { stringPool: materializationStringPool } : {}),
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
      sharedStringIndexes,
    })
    phaseRecorder.finish('metadata-parsing', metadataParsingStart)
  }
  const sharedStringResolutionStart = phaseRecorder.start()
  let sharedStrings: LargeSimpleSharedStrings = fallbackSharedStrings ?? []
  const referencedSharedStringIndexSet = referencedSharedStringIndexes.finalize()
  if (materializeCells && hasSharedStrings && referencedSharedStringIndexSet.size > 0) {
    const referencedSharedStrings =
      fallbackSharedStrings ??
      readReferencedLargeSimpleSharedStrings(zip, referencedSharedStringIndexSet, {
        deduplicateText: 'bounded',
        stringPool,
      })
    if (referencedSharedStrings === null) {
      return null
    }
    sharedStrings = referencedSharedStrings
  }
  delete zip[sharedStringsPath]
  if (materializeCells && hasSharedStrings && referencedSharedStringIndexSet.size > 0) {
    for (const [index, scanned] of scannedWorksheets.entries()) {
      if (!scanned || scanned.sharedStringIndexes.size === 0) {
        continue
      }
      const richSharedStringIndexes = collectReferencedLargeSimpleRichSharedStringIndexes(sharedStrings, scanned.sharedStringIndexes)
      if (!richSharedStringIndexes) {
        return null
      }
      if (scanned.cellScan.cellCount > lazySheetCellMaterializationThreshold) {
        scannedWorksheets[index] = {
          ...scanned,
          sharedStrings,
          sharedStringIndexes: emptyLargeSimpleSharedStringIndexes,
          hasUnresolvedSharedStringReferences: true,
          hasRichSharedStringReferences: richSharedStringIndexes.size > 0,
        }
        continue
      }
      if (richSharedStringIndexes.size === 0) {
        if (scanned.cellScan.arena.resolveSharedStrings(sharedStrings) === null) {
          return null
        }
        scannedWorksheets[index] = {
          ...scanned,
          sharedStringIndexes: emptyLargeSimpleSharedStringIndexes,
        }
        continue
      }
      if (!scanned.cellScan.arena.resolveSharedStringsExcept(sharedStrings, richSharedStringIndexes)) {
        return null
      }
      const sheetSharedStrings = createLargeSimpleSharedStringSubset(sharedStrings, richSharedStringIndexes)
      if (sheetSharedStrings === null) {
        return null
      }
      scannedWorksheets[index] = {
        ...scanned,
        sharedStrings: sheetSharedStrings,
        sharedStringIndexes: emptyLargeSimpleSharedStringIndexes,
        hasUnresolvedSharedStringReferences: true,
        hasRichSharedStringReferences: true,
      }
    }
    sharedStrings = []
  }
  referencedSharedStringIndexes.release()
  fallbackSharedStrings = null
  stringPool.release()
  materializationStringPool = undefined
  phaseRecorder.finish('shared-string-resolution', sharedStringResolutionStart)
  collectLargeSimpleImportGarbage()
  if (options.releaseZipSource !== true || allowPreReleaseSheetFinalization) {
    for (const [index, scanned] of scannedWorksheets.entries()) {
      if (
        !scanned ||
        scanned.sharedStringIndexes.size > 0 ||
        !canFinalizeSheetBeforeStyleParsing(scanned.name, scanned.cellScan, scanned.metadataScan, scanned.worksheetXml)
      ) {
        continue
      }
      const snapshotMaterializationStart = phaseRecorder.start()
      const resolvedRichTextCells = retainUnresolvedSharedStringReferences(scanned, sharedStrings)
      if (resolvedRichTextCells === null) {
        return null
      }
      appendParsedWorksheet(
        buildParsedWorksheet(
          scanned.name,
          scanned.order,
          {
            ...scanned.cellScan,
            richTextCells: mergeWorkbookRichTextCells(scanned.cellScan.richTextCells, resolvedRichTextCells),
          },
          scanned.worksheetXml,
          scanned.metadataScan,
          scanned.metadataInput,
          {
            materializeCells,
            releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
            styleCatalog,
            stylesByIndex: emptyStylesByIndex,
            ...(materializationStringPool ? { stringPool: materializationStringPool } : {}),
          },
        ),
      )
      scannedWorksheets[index] = undefined
      phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
      collectLargeSimpleImportGarbage()
    }
  }
  const styleParsingStart = phaseRecorder.start()
  const requiredStyleIndexes = new Set<number>()
  for (const scanned of scannedWorksheets) {
    if (!scanned) {
      continue
    }
    scanned.cellScan.styleIndexes.collectRequiredStyleIndexes(requiredStyleIndexes)
  }
  const parsedStyleArtifacts =
    materializeCells && hasStyles
      ? readLargeSimpleWorkbookStyleArtifactsFromChunks(
          (onChunk) => forEachLargeSimpleInflatedZipEntryChunk(zip, stylesPath, onChunk),
          requiredStyleIndexes,
        )
      : { stylesByIndex: new Map(), numberFormatsByStyleIndex: new Map() }
  const parsedStylesByIndex = parsedStyleArtifacts.stylesByIndex
  const parsedNumberFormatsByStyleIndex = parsedStyleArtifacts.numberFormatsByStyleIndex
  if (parsedNumberFormatsByStyleIndex === null) {
    return null
  }
  const stylesByIndex = parsedStylesByIndex ?? new Map()
  const numberFormatsByStyleIndex = parsedNumberFormatsByStyleIndex
  if (parsedStylesByIndex === null) {
    warnings.push(unsupportedCellStylesWarning)
  }
  requiredStyleIndexes.clear()
  const sheetNeedsStyleCoordinateMaterialization = (cellScan: ImportedWorksheetCellScan): boolean =>
    stylesByIndex.size > 0 ||
    (numberFormatsByStyleIndex.size > 0 && cellScan.cellCount <= lazySheetCellMaterializationNumberFormatThreshold)
  if (!stylesByIndex.size && !numberFormatsByStyleIndex.size) {
    releaseLargeSimpleStyleIndexes(scannedWorksheets)
  }
  const needsDeferredStyleCoordinateMaterialization = scannedWorksheets.some(
    (scanned) =>
      scanned && sheetNeedsStyleCoordinateMaterialization(scanned.cellScan) && !scanned.cellScan.styleIndexes.hasCoordinateStorage,
  )
  delete zip[stylesPath]
  phaseRecorder.finish('style-parsing', styleParsingStart)
  if (
    materializeCells &&
    stylesByIndex.size === 0 &&
    numberFormatsByStyleIndex.size === 0 &&
    (options.releaseZipSource !== true || allowPreReleaseSheetFinalization)
  ) {
    for (const [index, scanned] of scannedWorksheets.entries()) {
      if (
        !scanned ||
        scanned.sharedStringIndexes.size > 0 ||
        sheetHasRelationshipBackedArtifacts(scanned.name, scanned.metadataScan, scanned.worksheetXml)
      ) {
        continue
      }
      const snapshotMaterializationStart = phaseRecorder.start()
      const resolvedRichTextCells = retainUnresolvedSharedStringReferences(scanned, sharedStrings)
      if (resolvedRichTextCells === null) {
        return null
      }
      appendParsedWorksheet(
        buildParsedWorksheet(
          scanned.name,
          scanned.order,
          {
            ...scanned.cellScan,
            richTextCells: mergeWorkbookRichTextCells(scanned.cellScan.richTextCells, resolvedRichTextCells),
          },
          scanned.worksheetXml,
          scanned.metadataScan,
          scanned.metadataInput,
          {
            materializeCells,
            releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
            styleCatalog,
            stylesByIndex: emptyStylesByIndex,
            ...(materializationStringPool ? { stringPool: materializationStringPool } : {}),
          },
        ),
      )
      scannedWorksheets[index] = undefined
      phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
      collectLargeSimpleImportGarbage()
    }
  }
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
  const importedSlicerConnectionArtifacts =
    materializeCells && hasSlicerConnectionParts
      ? readImportedWorkbookSlicerConnectionArtifactsFromSheets(
          zip,
          largeSimpleSlicerConnectionSheetSources(scannedWorksheets, worksheetEntries),
          {
            workbookXml,
            workbookRelationshipsXml,
          },
        )
      : undefined
  const importedControlArtifacts = materializeCells
    ? readImportedWorkbookControlArtifactsFromSheetSources(zip, largeSimpleControlArtifactSheetSources(scannedWorksheets, worksheetEntries))
    : undefined
  const importedLegacyCommentVmlBySheet =
    materializeCells && hasLegacyCommentParts
      ? readImportedWorkbookLegacyCommentVmlFromSheetSources(
          zip,
          largeSimpleLegacyCommentVmlSheetSources(scannedWorksheets, worksheetEntries),
        )
      : null
  const releaseZipSourceForMaterialization = (): void => {
    const zipSourceReleaseStart = phaseRecorder.start()
    const zipSourceBytesBeforeRelease = readLazyXlsxZipSourceByteLength(zip)
    const artifactReleasePlan = prepareLargeSimplePackageArtifactsForZipRelease({
      ...(options.maxMaterializedLazyPackageArtifactBytes !== undefined
        ? { maxMaterializedBytes: options.maxMaterializedLazyPackageArtifactBytes }
        : {}),
      preservedArtifacts: [
        importedDataModelArtifacts,
        importedSlicerConnectionArtifacts,
        importedDrawingArtifacts?.artifacts,
        importedChartDrawingArtifacts?.drawingArtifacts.artifacts,
        importedChartDrawingArtifacts?.chartArtifacts.artifacts,
      ],
      opaqueArtifacts: [importedPivotArtifacts?.artifacts],
    })
    const retainZipSourceForLazyPackageArtifacts = zipSourceBytesBeforeRelease !== undefined && artifactReleasePlan.retainZipSource
    const releaseZipSource = !retainZipSourceForLazyPackageArtifacts && !options.replacementZipSource
    if (releaseZipSource) {
      releaseLazyXlsxZipSource(zip)
    }
    const ownedSourceReleaseEvidence =
      ownedSourceReleasedForReplacement || (retainZipSourceForLazyPackageArtifacts && options.replacementZipSource)
        ? undefined
        : options.releaseOwnedSourceBytes?.()
    phaseRecorder.finish('zip-source-release', zipSourceReleaseStart, {
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesBeforeRelease } : {}),
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesAfterRelease: readLazyXlsxZipSourceByteLength(zip) ?? 0 } : {}),
      ...ownedSourceReleaseEvidence,
    })
  }
  if (options.releaseZipSource === true && !needsDeferredStyleCoordinateMaterialization) {
    releaseZipSourceForMaterialization()
  }
  for (const [index, scanned] of scannedWorksheets.entries()) {
    if (!scanned) {
      continue
    }
    const snapshotMaterializationStart = phaseRecorder.start()
    const needsStyleCoordinatesForSheet = sheetNeedsStyleCoordinateMaterialization(scanned.cellScan)
    const resolvedRichTextCells = retainUnresolvedSharedStringReferences(scanned, sharedStrings)
    if (resolvedRichTextCells === null) {
      return null
    }
    const styleIndexes = needsStyleCoordinatesForSheet
      ? scanned.cellScan.styleIndexes.hasCoordinateStorage
        ? scanned.cellScan.styleIndexes
        : prepareLargeSimpleStyleIndexForWorksheet(zip, worksheetEntries, scanned, {
            hasSharedStrings,
            ...(options.allowUnsupportedFormulaText === undefined
              ? {}
              : { allowUnsupportedFormulaText: options.allowUnsupportedFormulaText }),
            ...(options.allowUnsupportedCellMetadata === undefined
              ? {}
              : { allowUnsupportedCellMetadata: options.allowUnsupportedCellMetadata }),
          })
      : scanned.cellScan.styleIndexes
    if (!styleIndexes) {
      return null
    }
    const cellScan = {
      ...scanned.cellScan,
      styleIndexes,
      richTextCells: mergeWorkbookRichTextCells(scanned.cellScan.richTextCells, resolvedRichTextCells),
    }
    const drawingArtifacts =
      importedChartDrawingArtifacts?.drawingArtifacts.sheetArtifactsByName.get(scanned.name) ??
      importedDrawingArtifacts?.sheetArtifactsByName.get(scanned.name)
    const controlArtifacts = importedControlArtifacts?.sheetArtifactsByName.get(scanned.name)
    const pivotArtifacts = sheetPivotArtifactsWithStreamedDefinitions(
      importedPivotArtifacts?.sheetArtifactsByName.get(scanned.name),
      scanned.metadataScan?.pivotTableDefinitionsXml,
    )
    const legacyCommentVml = importedLegacyCommentVmlBySheet?.get(scanned.name)
    const parsed = buildParsedWorksheet(
      scanned.name,
      scanned.order,
      cellScan,
      scanned.worksheetXml,
      scanned.metadataScan,
      {
        ...scanned.metadataInput,
        ...(drawingArtifacts ? { drawingArtifacts } : {}),
        ...(controlArtifacts ? { controlArtifacts } : {}),
        ...(pivotArtifacts ? { pivotArtifacts } : {}),
        ...(legacyCommentVml
          ? {
              legacyCommentVml: {
                ...legacyCommentVml,
                commentSignature: legacyCommentThreadSignature(undefined),
              },
            }
          : {}),
      },
      {
        materializeCells,
        releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
        styleCatalog,
        stylesByIndex,
        numberFormatsByStyleIndex,
        ...(materializationStringPool ? { stringPool: materializationStringPool } : {}),
      },
    )
    appendParsedWorksheet(parsed)
    scannedWorksheets[index] = undefined
    phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
    collectLargeSimpleImportGarbage()
  }
  if (options.releaseZipSource === true && needsDeferredStyleCoordinateMaterialization) {
    releaseZipSourceForMaterialization()
  }
  sharedStrings = []
  stringPool.release()
  const sortedImportedTables =
    importedTables.length > 0 ? importedTables.toSorted((left, right) => left.name.localeCompare(right.name)) : undefined
  const importedSheets = sheets.map((sheet) => applyImportedAutoFilterRowVisibility(sheet, sortedImportedTables))
  const hasFormulaCells = sheetStats.some((entry) => entry.formulaCellCount > 0)
  const workbookMetadata =
    workbookDefinedNames.definedNames ||
    importedWorkbookProperties ||
    importedWorkbookDocumentProperties ||
    importedChartDrawingArtifacts?.drawingArtifacts.artifacts ||
    importedDrawingArtifacts?.artifacts ||
    importedChartDrawingArtifacts?.chartArtifacts.artifacts ||
    importedChartDrawingArtifacts?.chartArtifacts.chartSheetArtifacts ||
    importedChartDrawingArtifacts?.charts ||
    importedPivotArtifacts?.artifacts ||
    importedControlArtifacts?.artifacts ||
    sortedImportedTables ||
    styleCatalog.size > 0 ||
    importedDataModelArtifacts ||
    importedExternalLinkArtifacts ||
    importedSlicerConnectionArtifacts ||
    importedWorkbookCellMetadata ||
    hasFormulaCells
      ? {
          ...(importedWorkbookProperties ? { properties: importedWorkbookProperties } : {}),
          ...(importedWorkbookDocumentProperties ? { documentPropertyArtifacts: importedWorkbookDocumentProperties } : {}),
          ...(workbookDefinedNames.definedNames ? { definedNames: workbookDefinedNames.definedNames } : {}),
          ...(importedChartDrawingArtifacts?.drawingArtifacts.artifacts
            ? { drawingArtifacts: importedChartDrawingArtifacts.drawingArtifacts.artifacts }
            : importedDrawingArtifacts?.artifacts
              ? { drawingArtifacts: importedDrawingArtifacts.artifacts }
              : {}),
          ...(importedChartDrawingArtifacts?.chartArtifacts.artifacts
            ? { chartArtifacts: importedChartDrawingArtifacts.chartArtifacts.artifacts }
            : {}),
          ...(importedChartDrawingArtifacts?.chartArtifacts.chartSheetArtifacts
            ? { chartSheetArtifacts: importedChartDrawingArtifacts.chartArtifacts.chartSheetArtifacts }
            : {}),
          ...(importedChartDrawingArtifacts?.charts ? { charts: importedChartDrawingArtifacts.charts } : {}),
          ...(importedPivotArtifacts?.artifacts ? { pivotArtifacts: importedPivotArtifacts.artifacts } : {}),
          ...(importedControlArtifacts?.artifacts ? { controlArtifacts: importedControlArtifacts.artifacts } : {}),
          ...(sortedImportedTables ? { tables: sortedImportedTables } : {}),
          ...(styleCatalog.size > 0 ? { styles: [...styleCatalog.values()] } : {}),
          ...(importedDataModelArtifacts ? { dataModelArtifacts: importedDataModelArtifacts } : {}),
          ...(importedExternalLinkArtifacts ? { externalLinkArtifacts: importedExternalLinkArtifacts } : {}),
          ...(importedSlicerConnectionArtifacts ? { slicerConnectionArtifacts: importedSlicerConnectionArtifacts } : {}),
          ...(importedWorkbookCellMetadata ? { cellMetadata: importedWorkbookCellMetadata } : {}),
          ...(hasFormulaCells
            ? {
                calculationSettings: {
                  mode: 'automatic' as const,
                  compatibilityMode: 'excel-modern' as const,
                  fullCalcOnLoad: false,
                  forceFullCalc: false,
                },
              }
            : {}),
        }
      : undefined
  const runtimeSheetCells = buildLargeSimpleRuntimeSheetCells(sheetStats, importedSheets)
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: workbookName,
      ...(workbookMetadata ? { metadata: workbookMetadata } : {}),
    },
    sheets: importedSheets,
  }
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
