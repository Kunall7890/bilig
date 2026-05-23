import type {
  CellRangeRef,
  SheetMetadataSnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookSheetArrayFormulasSnapshot,
  WorkbookSheetDataTableFormulasSnapshot,
  WorkbookSheetTabColorSnapshot,
  WorkbookSparklinesSnapshot,
} from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook'
import type { WorkbookAxisMetadataRecord, WorkbookFreezePaneRecord, WorkbookSheetTabColorRecord, WorkbookStore } from './workbook-store.js'

function cloneSnapshotRangeRef(range: CellRangeRef): CellRangeRef {
  return {
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  }
}

function axisMetadataToSnapshot(records: readonly WorkbookAxisMetadataRecord[]): WorkbookAxisMetadataSnapshot[] {
  return records.map((record) => {
    const snapshot: WorkbookAxisMetadataSnapshot = {
      start: record.start,
      count: record.count,
    }
    if (record.size !== null) {
      snapshot.size = record.size
    }
    if (record.hidden !== null) {
      snapshot.hidden = record.hidden
    }
    if (record.filterHidden !== null) {
      snapshot.filterHidden = record.filterHidden
    }
    if (record.styleIndex !== undefined && record.styleIndex !== null) {
      snapshot.styleIndex = record.styleIndex
    }
    if (record.xlsxWidth !== undefined && record.xlsxWidth !== null) {
      snapshot.xlsxWidth = record.xlsxWidth
    }
    if (record.xlsxHeight !== undefined && record.xlsxHeight !== null) {
      snapshot.xlsxHeight = record.xlsxHeight
    }
    if (record.customFormat !== undefined && record.customFormat !== null) {
      snapshot.customFormat = record.customFormat
    }
    if (record.customWidth !== undefined && record.customWidth !== null) {
      snapshot.customWidth = record.customWidth
    }
    if (record.bestFit !== undefined && record.bestFit !== null) {
      snapshot.bestFit = record.bestFit
    }
    if (record.outlineLevel !== undefined && record.outlineLevel !== null) {
      snapshot.outlineLevel = record.outlineLevel
    }
    if (record.collapsed !== undefined && record.collapsed !== null) {
      snapshot.collapsed = record.collapsed
    }
    if (record.customHeight !== undefined && record.customHeight !== null) {
      snapshot.customHeight = record.customHeight
    }
    if (record.thickTop !== undefined && record.thickTop !== null) {
      snapshot.thickTop = record.thickTop
    }
    if (record.thickBottom !== undefined && record.thickBottom !== null) {
      snapshot.thickBottom = record.thickBottom
    }
    return snapshot
  })
}

function freezePaneToSnapshot(record: WorkbookFreezePaneRecord | undefined): WorkbookFreezePaneSnapshot | undefined {
  if (!record) {
    return undefined
  }
  const snapshot: WorkbookFreezePaneSnapshot = { rows: record.rows, cols: record.cols }
  if (record.topLeftCell !== undefined) {
    snapshot.topLeftCell = record.topLeftCell
  }
  if (record.activePane !== undefined) {
    snapshot.activePane = record.activePane
  }
  return snapshot
}

function mergeRangeToSnapshot(record: WorkbookMergeRangeSnapshot): WorkbookMergeRangeSnapshot {
  return {
    sheetName: record.sheetName,
    startAddress: record.startAddress,
    endAddress: record.endAddress,
  }
}

function sheetTabColorToSnapshot(record: WorkbookSheetTabColorRecord | undefined): WorkbookSheetTabColorSnapshot | undefined {
  if (!record) {
    return undefined
  }
  const snapshot: WorkbookSheetTabColorSnapshot = {}
  if (record.rgb !== undefined) {
    snapshot.rgb = record.rgb
  }
  if (record.theme !== undefined) {
    snapshot.theme = record.theme
  }
  if (record.tint !== undefined) {
    snapshot.tint = record.tint
  }
  if (record.indexed !== undefined) {
    snapshot.indexed = record.indexed
  }
  if (record.auto !== undefined) {
    snapshot.auto = record.auto
  }
  return snapshot
}

function dataTableFormulasToSnapshot(
  formulas: WorkbookSheetDataTableFormulasSnapshot | undefined,
): WorkbookSheetDataTableFormulasSnapshot | undefined {
  if (!formulas || formulas.formulas.length === 0) {
    return undefined
  }
  return { formulas: formulas.formulas.map((formula) => ({ ...formula })) }
}

function arrayFormulasToSnapshot(formulas: WorkbookSheetArrayFormulasSnapshot | undefined): WorkbookSheetArrayFormulasSnapshot | undefined {
  if (!formulas || formulas.formulas.length === 0) {
    return undefined
  }
  return { formulas: formulas.formulas.map((formula) => ({ ...formula })) }
}

function sparklinesToSnapshot(sparklines: WorkbookSparklinesSnapshot | undefined): WorkbookSparklinesSnapshot | undefined {
  return sparklines ? { xml: sparklines.xml } : undefined
}

export function exportSheetMetadata(workbook: WorkbookStore, sheetName: string): SheetMetadataSnapshot | undefined {
  const sheet = workbook.getSheet(sheetName)
  const rows = workbook.listRowAxisEntries(sheetName)
  const columns = workbook.listColumnAxisEntries(sheetName)
  const rowMetadata = axisMetadataToSnapshot(workbook.listRowMetadata(sheetName))
  const columnMetadata = axisMetadataToSnapshot(workbook.listColumnMetadata(sheetName))
  const sheetFormatPr = workbook.getSheetFormatPr(sheetName)
  const visibility = workbook.getSheetVisibility(sheetName)
  const styleRanges = workbook.listStyleRanges(sheetName).map((record) => ({
    range: cloneSnapshotRangeRef(record.range),
    styleId: record.styleId,
  }))
  const formatRanges = workbook.listFormatRanges(sheetName).map((record) => ({
    range: cloneSnapshotRangeRef(record.range),
    formatId: record.formatId,
  }))
  const freezePane = freezePaneToSnapshot(workbook.getFreezePane(sheetName))
  const tabColor = sheetTabColorToSnapshot(workbook.getSheetTabColor(sheetName))
  const merges = workbook.listMergeRanges(sheetName).map(mergeRangeToSnapshot)
  const sheetProtection = workbook.getSheetProtection(sheetName)
  const filters = workbook.listFilters(sheetName).map((filter) => structuredClone(filter.range))
  const sorts = workbook.listSorts(sheetName).map((sort) => ({
    range: { ...sort.range },
    keys: sort.keys.map((key) => ({ ...key })),
  }))
  const validations = workbook.listDataValidations(sheetName).map((validation) => structuredClone(validation))
  const conditionalFormats = workbook.listConditionalFormats(sheetName).map((format) => structuredClone(format))
  const conditionalFormatArtifacts = workbook.getConditionalFormatArtifacts(sheetName)
  const drawingArtifacts = workbook.getSheetDrawingArtifacts(sheetName)
  const threadedCommentArtifacts = workbook.getSheetThreadedCommentArtifacts(sheetName)
  const legacyCommentVml = workbook.getSheetLegacyCommentVml(sheetName)
  const protectedRanges = workbook.listRangeProtections(sheetName).map((protection) => structuredClone(protection))
  const commentThreads = workbook.listCommentThreads(sheetName).map((thread) => structuredClone(thread))
  const notes = workbook.listNotes(sheetName).map((note) => structuredClone(note))
  const hyperlinks = workbook.listHyperlinks(sheetName).map((hyperlink) => structuredClone(hyperlink))
  const arrayFormulas = arrayFormulasToSnapshot(sheet?.arrayFormulas)
  const dataTableFormulas = dataTableFormulasToSnapshot(sheet?.dataTableFormulas)
  const ignoredErrors = sheet?.ignoredErrors ? { xml: sheet.ignoredErrors.xml } : undefined
  const printPageSetup = sheet?.printPageSetup ? structuredClone(sheet.printPageSetup) : undefined
  const printerSettings = sheet?.printerSettings ? structuredClone(sheet.printerSettings) : undefined
  const sparklines = sparklinesToSnapshot(sheet?.sparklines)
  const richTextArtifacts = sheet?.richTextArtifacts ? structuredClone(sheet.richTextArtifacts) : undefined
  const cellMetadataRefs = sheet?.cellMetadataRefs ? structuredClone(sheet.cellMetadataRefs) : undefined

  if (
    rows.length === 0 &&
    columns.length === 0 &&
    rowMetadata.length === 0 &&
    columnMetadata.length === 0 &&
    sheetFormatPr === undefined &&
    visibility === undefined &&
    styleRanges.length === 0 &&
    formatRanges.length === 0 &&
    freezePane === undefined &&
    tabColor === undefined &&
    merges.length === 0 &&
    sheetProtection === undefined &&
    filters.length === 0 &&
    sorts.length === 0 &&
    validations.length === 0 &&
    conditionalFormats.length === 0 &&
    conditionalFormatArtifacts === undefined &&
    drawingArtifacts === undefined &&
    threadedCommentArtifacts === undefined &&
    legacyCommentVml === undefined &&
    protectedRanges.length === 0 &&
    commentThreads.length === 0 &&
    notes.length === 0 &&
    hyperlinks.length === 0 &&
    arrayFormulas === undefined &&
    dataTableFormulas === undefined &&
    ignoredErrors === undefined &&
    printPageSetup === undefined &&
    printerSettings === undefined &&
    sparklines === undefined &&
    richTextArtifacts === undefined &&
    cellMetadataRefs === undefined
  ) {
    return undefined
  }

  const metadata: SheetMetadataSnapshot = {}
  if (rows.length > 0) {
    metadata.rows = rows
  }
  if (columns.length > 0) {
    metadata.columns = columns
  }
  if (rowMetadata.length > 0) {
    metadata.rowMetadata = rowMetadata
  }
  if (columnMetadata.length > 0) {
    metadata.columnMetadata = columnMetadata
  }
  if (sheetFormatPr) {
    metadata.sheetFormatPr = sheetFormatPr
  }
  if (visibility) {
    metadata.visibility = visibility
  }
  if (styleRanges.length > 0) {
    metadata.styleRanges = styleRanges
  }
  if (formatRanges.length > 0) {
    metadata.formatRanges = formatRanges
  }
  if (freezePane) {
    metadata.freezePane = freezePane
  }
  if (tabColor) {
    metadata.tabColor = tabColor
  }
  if (merges.length > 0) {
    metadata.merges = merges
  }
  if (sheetProtection) {
    metadata.sheetProtection = structuredClone(sheetProtection)
  }
  if (filters.length > 0) {
    metadata.filters = filters
  }
  if (sorts.length > 0) {
    metadata.sorts = sorts
  }
  if (validations.length > 0) {
    metadata.validations = validations
  }
  if (conditionalFormats.length > 0) {
    metadata.conditionalFormats = conditionalFormats
  }
  if (conditionalFormatArtifacts) {
    metadata.conditionalFormatArtifacts = { xml: conditionalFormatArtifacts.xml }
  }
  if (drawingArtifacts) {
    metadata.drawingArtifacts = {
      relationshipTarget: drawingArtifacts.relationshipTarget,
      ...(drawingArtifacts.preservedChartRelationshipIds !== undefined
        ? { preservedChartRelationshipIds: [...drawingArtifacts.preservedChartRelationshipIds] }
        : {}),
    }
  }
  if (threadedCommentArtifacts) {
    metadata.threadedCommentArtifacts = {
      relationships: structuredClone(threadedCommentArtifacts.relationships),
    }
  }
  if (legacyCommentVml) {
    metadata.legacyCommentVml = {
      relationshipTarget: legacyCommentVml.relationshipTarget,
      vmlXml: legacyCommentVml.vmlXml,
      ...(legacyCommentVml.commentsRelationshipTarget !== undefined
        ? { commentsRelationshipTarget: legacyCommentVml.commentsRelationshipTarget }
        : {}),
      ...(legacyCommentVml.commentsXml !== undefined ? { commentsXml: legacyCommentVml.commentsXml } : {}),
      commentSignature: legacyCommentVml.commentSignature,
    }
  }
  if (protectedRanges.length > 0) {
    metadata.protectedRanges = protectedRanges
  }
  if (commentThreads.length > 0) {
    metadata.commentThreads = commentThreads
  }
  if (notes.length > 0) {
    metadata.notes = notes
  }
  if (hyperlinks.length > 0) {
    metadata.hyperlinks = hyperlinks
  }
  if (arrayFormulas) {
    metadata.arrayFormulas = arrayFormulas
  }
  if (dataTableFormulas) {
    metadata.dataTableFormulas = dataTableFormulas
  }
  if (ignoredErrors) {
    metadata.ignoredErrors = ignoredErrors
  }
  if (printPageSetup) {
    metadata.printPageSetup = printPageSetup
  }
  if (printerSettings) {
    metadata.printerSettings = printerSettings
  }
  if (sparklines) {
    metadata.sparklines = sparklines
  }
  if (richTextArtifacts) {
    metadata.richTextArtifacts = richTextArtifacts
  }
  if (cellMetadataRefs) {
    metadata.cellMetadataRefs = cellMetadataRefs
  }
  return metadata
}

export function sheetMetadataToOps(
  workbook: WorkbookStore,
  sheetName: string,
  options: {
    includeAxisEntries?: boolean
  } = {},
): EngineOp[] {
  const includeAxisEntries = options.includeAxisEntries ?? true
  const ops: EngineOp[] = []
  if (includeAxisEntries) {
    workbook.listRowAxisEntries(sheetName).forEach((entry) => {
      ops.push({ kind: 'insertRows', sheetName, start: entry.index, count: 1, entries: [entry] })
    })
    workbook.listColumnAxisEntries(sheetName).forEach((entry) => {
      ops.push({
        kind: 'insertColumns',
        sheetName,
        start: entry.index,
        count: 1,
        entries: [entry],
      })
    })
  }
  workbook.listRowMetadata(sheetName).forEach((record) => {
    ops.push({
      kind: 'updateRowMetadata',
      sheetName,
      start: record.start,
      count: record.count,
      size: record.size,
      hidden: record.hidden,
      filterHidden: record.filterHidden,
    })
  })
  workbook.listColumnMetadata(sheetName).forEach((record) => {
    ops.push({
      kind: 'updateColumnMetadata',
      sheetName,
      start: record.start,
      count: record.count,
      size: record.size,
      hidden: record.hidden,
    })
  })
  workbook.listStyleRanges(sheetName).forEach((record) => {
    ops.push({ kind: 'setStyleRange', range: { ...record.range }, styleId: record.styleId })
  })
  workbook.listFormatRanges(sheetName).forEach((record) => {
    ops.push({ kind: 'setFormatRange', range: { ...record.range }, formatId: record.formatId })
  })
  const freezePane = workbook.getFreezePane(sheetName)
  if (freezePane) {
    ops.push({ kind: 'setFreezePane', sheetName, rows: freezePane.rows, cols: freezePane.cols })
  }
  workbook.listMergeRanges(sheetName).forEach((record) => {
    ops.push({ kind: 'mergeCells', range: mergeRangeToSnapshot(record) })
  })
  const sheetProtection = workbook.getSheetProtection(sheetName)
  if (sheetProtection) {
    ops.push({ kind: 'setSheetProtection', protection: structuredClone(sheetProtection) })
  }
  workbook.listFilters(sheetName).forEach((record) => {
    ops.push({ kind: 'setFilter', sheetName, range: structuredClone(record.range) })
  })
  workbook.listSorts(sheetName).forEach((record) => {
    ops.push({
      kind: 'setSort',
      sheetName,
      range: { ...record.range },
      keys: record.keys.map((key) => ({ ...key })),
    })
  })
  workbook.listDataValidations(sheetName).forEach((record) => {
    ops.push({
      kind: 'setDataValidation',
      validation: structuredClone(record),
    })
  })
  workbook.listConditionalFormats(sheetName).forEach((record) => {
    ops.push({
      kind: 'upsertConditionalFormat',
      format: structuredClone(record),
    })
  })
  const conditionalFormatArtifacts = workbook.getConditionalFormatArtifacts(sheetName)
  if (conditionalFormatArtifacts) {
    ops.push({
      kind: 'setConditionalFormatArtifacts',
      sheetName,
      artifacts: { xml: conditionalFormatArtifacts.xml },
    })
  }
  workbook.listRangeProtections(sheetName).forEach((record) => {
    ops.push({
      kind: 'upsertRangeProtection',
      protection: structuredClone(record),
    })
  })
  workbook.listCommentThreads(sheetName).forEach((record) => {
    ops.push({
      kind: 'upsertCommentThread',
      thread: structuredClone(record),
    })
  })
  workbook.listNotes(sheetName).forEach((record) => {
    ops.push({
      kind: 'upsertNote',
      note: structuredClone(record),
    })
  })
  workbook.listHyperlinks(sheetName).forEach((record) => {
    ops.push({
      kind: 'upsertHyperlink',
      hyperlink: structuredClone(record),
    })
  })
  return ops
}
