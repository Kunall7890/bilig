import {
  MAX_COLS,
  MAX_ROWS,
  type CellRangeRef,
  type SheetFormatRangeSnapshot,
  type SheetStyleRangeSnapshot,
  type WorkbookSheetDataTableFormulaSnapshot,
  type WorkbookSheetDataTableFormulasSnapshot,
} from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  parseFormula,
  type FormulaNode,
  type StructuralAxisTransform,
} from '@bilig/formula'
import { mapStructuralBoundary } from '../../engine-structural-utils.js'
import { normalizeDefinedName, type WorkbookTableRecord } from '../../workbook-store.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'
import { rewriteArrayFormulasForStructuralTransform } from './structure-array-formula-metadata-rewrite.js'
import { rewriteLegacyCommentVmlForStructuralTransform } from './structure-legacy-comment-vml-rewrite.js'
import { rewriteFormulaSourceForDeletedStructuredReferences } from './structure-structured-ref-rewrite.js'
import { chartGeometryFromAnchor, rewriteChartAnchorForStructuralTransform } from './structure-chart-anchor-metadata-rewrite.js'
import { rewriteThreadedCommentArtifactsForStructuralTransform } from './structure-threaded-comment-artifact-rewrite.js'
import { nextGeneratedTableColumnName, normalizeTableColumnName } from './table-column-name-helpers.js'

type StructureMetadataRewriteArgs = Pick<CreateEngineStructureServiceArgs, 'state' | 'clearOwnedPivot'>

type MetadataRangeLike = {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
}

interface StructuralTableHeaderCellWrite {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly value: string
}

export interface DeletedTableColumnReference {
  readonly tableName: string
  readonly columnName: string
}

type WorkbookTableColumnRecord = NonNullable<WorkbookTableRecord['columns']>[number]
const METADATA_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i

interface DataTableStructuralGeometry {
  readonly address: string
  readonly ref: string
  readonly rowInput: string
  readonly columnInput?: string
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function decodeXmlAttribute(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function escapeXmlAttribute(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/"/gu, '&quot;').replace(/'/gu, '&apos;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;')
}

function quoteFormulaSheetName(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/u.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function invalidDefinedNameReferenceFormula(sheetName: string): { readonly kind: 'formula'; readonly formula: string } {
  return { kind: 'formula', formula: `=${quoteFormulaSheetName(sheetName)}!#REF!` }
}

function rewriteDataTableFormulaXmlAttribute(formulaXml: string, attributeName: string, nextValue: string): string | undefined {
  if (readXmlAttribute(formulaXml, attributeName) === null) {
    return undefined
  }
  let replaced = false
  const nextXml = formulaXml.replace(new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u'), (_source: string, quote: string) => {
    replaced = true
    return `${attributeName}=${quote}${escapeXmlAttribute(nextValue)}${quote}`
  })
  return replaced ? nextXml : undefined
}

function readDataTableFormulaXmlAttribute(formulaXml: string, attributeName: string): string | undefined {
  const value = readXmlAttribute(formulaXml, attributeName)
  return value === null ? undefined : decodeXmlAttribute(value)
}

function normalizeDataTableCellAddress(address: string): string | undefined {
  const parsed = parseUnboundedMetadataCellAddress(address)
  return parsed ? formatAddress(parsed[0], parsed[1]) : undefined
}

function topLeftAddressOfDataTableRef(value: string): string | undefined {
  const [start, end, extra] = value.split(':')
  if (!start || extra !== undefined) {
    return undefined
  }
  if (end !== undefined && !normalizeDataTableCellAddress(end)) {
    return undefined
  }
  return normalizeDataTableCellAddress(start)
}

function dataTableAddressAxisIndex(address: string, axis: StructuralAxisTransform['axis']): number | undefined {
  const parsed = parseUnboundedMetadataCellAddress(address)
  return parsed ? parsed[axis === 'row' ? 0 : 1] : undefined
}

function dataTableRefAxisBounds(
  ref: string,
  axis: StructuralAxisTransform['axis'],
): { readonly min: number; readonly max: number } | undefined {
  const [start, end, extra] = ref.split(':')
  if (!start || extra !== undefined) {
    return undefined
  }
  const startIndex = dataTableAddressAxisIndex(start, axis)
  const endIndex = dataTableAddressAxisIndex(end ?? start, axis)
  if (startIndex === undefined || endIndex === undefined) {
    return undefined
  }
  return { min: Math.min(startIndex, endIndex), max: Math.max(startIndex, endIndex) }
}

function dataTableFormulaAxisBounds(
  geometry: DataTableStructuralGeometry,
  axis: StructuralAxisTransform['axis'],
): { readonly min: number; readonly max: number } | undefined {
  const refBounds = dataTableRefAxisBounds(geometry.ref, axis)
  const addressIndex = dataTableAddressAxisIndex(geometry.address, axis)
  const rowInputIndex = dataTableAddressAxisIndex(geometry.rowInput, axis)
  const columnInputIndex = geometry.columnInput ? dataTableAddressAxisIndex(geometry.columnInput, axis) : undefined
  if (!refBounds || addressIndex === undefined || rowInputIndex === undefined || (geometry.columnInput && columnInputIndex === undefined)) {
    return undefined
  }
  const indices = [refBounds.min, refBounds.max, addressIndex, rowInputIndex, ...(columnInputIndex === undefined ? [] : [columnInputIndex])]
  return { min: Math.min(...indices), max: Math.max(...indices) }
}

function isProvenDataTableStructuralTransform(geometry: DataTableStructuralGeometry, transform: StructuralAxisTransform): boolean {
  const bounds = dataTableFormulaAxisBounds(geometry, transform.axis)
  if (!bounds) {
    return false
  }
  switch (transform.kind) {
    case 'insert':
      return transform.start <= bounds.min || transform.start > bounds.max
    case 'delete':
      return transform.start > bounds.max
    case 'move':
      return transform.start > bounds.max && transform.target > bounds.max
  }
}

function rewriteDataTableFormulaAddressAttribute(value: string, transform: StructuralAxisTransform): string | undefined {
  try {
    return rewriteAddressForStructuralTransform(value, transform)
  } catch {
    return undefined
  }
}

function rewriteDataTableFormulaRefAttribute(value: string, transform: StructuralAxisTransform): string | undefined {
  const [start, end, extra] = value.split(':')
  if (!start) {
    return undefined
  }
  if (extra !== undefined) {
    return undefined
  }
  if (!end) {
    return rewriteDataTableFormulaAddressAttribute(start, transform)
  }
  try {
    const rewritten = rewriteRangeForStructuralTransform(start, end, transform)
    return rewritten ? `${rewritten.startAddress}:${rewritten.endAddress}` : undefined
  } catch {
    return undefined
  }
}

function rewriteDataTableFormulaForStructuralTransform(
  formula: WorkbookSheetDataTableFormulaSnapshot,
  transform: StructuralAxisTransform,
): WorkbookSheetDataTableFormulaSnapshot | undefined {
  const address = rewriteDataTableFormulaAddressAttribute(formula.address, transform)
  if (!address) {
    return undefined
  }
  const ref = readDataTableFormulaXmlAttribute(formula.formulaXml, 'ref')
  const rowInput = readDataTableFormulaXmlAttribute(formula.formulaXml, 'r1')
  if (!ref || !rowInput) {
    return undefined
  }

  const nextRef = rewriteDataTableFormulaRefAttribute(ref, transform)
  const nextRowInput = rewriteDataTableFormulaAddressAttribute(rowInput, transform)
  if (!nextRef || !nextRowInput) {
    return undefined
  }
  if (normalizeDataTableCellAddress(address) !== topLeftAddressOfDataTableRef(nextRef)) {
    return undefined
  }

  const columnInput = readDataTableFormulaXmlAttribute(formula.formulaXml, 'r2')
  if (
    !isProvenDataTableStructuralTransform({ address: formula.address, ref, rowInput, ...(columnInput ? { columnInput } : {}) }, transform)
  ) {
    return undefined
  }

  const nextColumnInput = columnInput ? rewriteDataTableFormulaAddressAttribute(columnInput, transform) : undefined
  if (columnInput && !nextColumnInput) {
    return undefined
  }

  const withRef = rewriteDataTableFormulaXmlAttribute(formula.formulaXml, 'ref', nextRef)
  const withRowInput = withRef ? rewriteDataTableFormulaXmlAttribute(withRef, 'r1', nextRowInput) : undefined
  const formulaXml =
    withRowInput && nextColumnInput ? rewriteDataTableFormulaXmlAttribute(withRowInput, 'r2', nextColumnInput) : withRowInput
  return formulaXml ? { address, formulaXml } : undefined
}

function rewriteDataTableFormulasForStructuralTransform(
  formulas: WorkbookSheetDataTableFormulasSnapshot | undefined,
  transform: StructuralAxisTransform,
): WorkbookSheetDataTableFormulasSnapshot | undefined {
  const rewritten = formulas?.formulas.flatMap((formula) => {
    const nextFormula = rewriteDataTableFormulaForStructuralTransform(formula, transform)
    return nextFormula ? [nextFormula] : []
  })
  return rewritten && rewritten.length > 0 ? { formulas: rewritten } : undefined
}

function rewriteConditionalFormatArtifactReference(value: string, transform: StructuralAxisTransform): string | undefined {
  const [start, end, extra] = value.split(':')
  if (!start || extra !== undefined) {
    return undefined
  }
  try {
    if (!end) {
      return rewriteAddressForStructuralTransform(start, transform)
    }
    const rewritten = rewriteRangeForStructuralTransform(start, end, transform)
    return rewritten ? `${rewritten.startAddress}:${rewritten.endAddress}` : undefined
  } catch {
    return undefined
  }
}

function rewriteConditionalFormatArtifactSqref(value: string, transform: StructuralAxisTransform): string | undefined {
  const references = value.trim().split(/\s+/u).filter(Boolean)
  if (references.length === 0) {
    return undefined
  }
  const rewritten = references.flatMap((reference) => {
    const nextReference = rewriteConditionalFormatArtifactReference(reference, transform)
    return nextReference ? [nextReference] : []
  })
  return rewritten.length > 0 ? rewritten.join(' ') : undefined
}

function rewriteConditionalFormatArtifactBlock(block: string, transform: StructuralAxisTransform): string | undefined {
  const openingTag = /^<conditionalFormatting\b[^>]*>/u.exec(block)?.[0]
  if (!openingTag) {
    return block
  }
  const sqref = readXmlAttribute(openingTag, 'sqref')
  if (sqref === null) {
    return block
  }
  const nextSqref = rewriteConditionalFormatArtifactSqref(decodeXmlAttribute(sqref), transform)
  if (!nextSqref) {
    return undefined
  }

  const nextOpeningTag = openingTag.replace(/\bsqref=("|')([\s\S]*?)\1/u, (_source: string, quote: string) => {
    return `sqref=${quote}${escapeXmlAttribute(nextSqref)}${quote}`
  })
  return `${nextOpeningTag}${block.slice(openingTag.length)}`
}

function rewriteConditionalFormatArtifactXmlForStructuralTransform(xml: string, transform: StructuralAxisTransform): string | undefined {
  let matchedBlock = false
  const rewrittenXml = xml.replace(/<conditionalFormatting\b[^>]*>[\s\S]*?<\/conditionalFormatting>/gu, (block) => {
    matchedBlock = true
    return rewriteConditionalFormatArtifactBlock(block, transform) ?? ''
  })
  if (!matchedBlock) {
    return xml
  }
  return rewrittenXml.trim().length > 0 ? rewrittenXml : undefined
}

export function rewriteDefinedNamesForStructuralTransform(
  args: StructureMetadataRewriteArgs,
  sheetName: string,
  transform: StructuralAxisTransform,
  deletedTableColumns: readonly DeletedTableColumnReference[],
  changedTableNames: ReadonlySet<string>,
): Set<string> {
  const workbook = args.state.workbook
  const changedNames = new Set<string>()
  workbook.listDefinedNames().forEach((record) => {
    if (typeof record.value === 'string' && record.value.startsWith('=')) {
      const nextFormula = rewriteDefinedNameFormulaOrNull(record.value.slice(1), sheetName, transform, deletedTableColumns)
      if (nextFormula === null) {
        return
      }
      if (`=${nextFormula}` !== record.value) {
        workbook.setDefinedName(record.name, `=${nextFormula}`, record.scopeSheetName)
        changedNames.add(normalizeDefinedName(record.name))
      } else if (formulaReferencesChangedTable(nextFormula, changedTableNames)) {
        changedNames.add(normalizeDefinedName(record.name))
      }
      return
    }
    if (typeof record.value !== 'object' || !record.value) {
      return
    }
    switch (record.value.kind) {
      case 'formula': {
        const nextFormula = rewriteDefinedNameFormulaOrNull(
          record.value.formula.startsWith('=') ? record.value.formula.slice(1) : record.value.formula,
          sheetName,
          transform,
          deletedTableColumns,
        )
        if (nextFormula === null) {
          return
        }
        const nextValue = {
          ...record.value,
          formula: record.value.formula.startsWith('=') ? `=${nextFormula}` : nextFormula,
        }
        if (nextValue.formula !== record.value.formula) {
          workbook.setDefinedName(record.name, nextValue, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
        } else if (formulaReferencesChangedTable(nextFormula, changedTableNames)) {
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'structured-ref': {
        const value = record.value
        if (deletedTableColumns.some((deleted) => tableColumnReferenceMatches(deleted, value.tableName, value.columnName))) {
          workbook.setDefinedName(record.name, { kind: 'formula', formula: '=#REF!' }, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (structuredReferenceTouchesChangedTable(value.tableName, changedTableNames)) {
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'cell-ref': {
        if (record.value.sheetName !== sheetName) {
          return
        }
        const nextAddress = rewriteAddressForStructuralTransform(record.value.address, transform)
        if (!nextAddress) {
          workbook.setDefinedName(record.name, invalidDefinedNameReferenceFormula(record.value.sheetName), record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextAddress !== record.value.address) {
          workbook.setDefinedName(
            record.name,
            {
              ...record.value,
              address: nextAddress,
            },
            record.scopeSheetName,
          )
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'range-ref': {
        if (record.value.sheetName !== sheetName) {
          return
        }
        const nextRange = rewriteMetadataRangeForStructuralTransform(record.value, transform)
        if (!nextRange) {
          workbook.setDefinedName(record.name, invalidDefinedNameReferenceFormula(record.value.sheetName), record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
          return
        }
        if (nextRange.startAddress !== record.value.startAddress || nextRange.endAddress !== record.value.endAddress) {
          workbook.setDefinedName(record.name, nextRange, record.scopeSheetName)
          changedNames.add(normalizeDefinedName(record.name))
        }
        return
      }
      case 'scalar':
        return
    }
  })
  return changedNames
}

function rewriteDefinedNameFormulaOrNull(
  formula: string,
  sheetName: string,
  transform: StructuralAxisTransform,
  deletedTableColumns: readonly DeletedTableColumnReference[],
): string | null {
  try {
    const structuralFormula = rewriteFormulaForStructuralTransform(formula, sheetName, sheetName, transform)
    return rewriteFormulaSourceForDeletedStructuredReferences(structuralFormula, deletedTableColumns) ?? structuralFormula
  } catch {
    return null
  }
}

function tableColumnReferenceMatches(deleted: DeletedTableColumnReference, tableName: string, columnName: string): boolean {
  return (
    normalizeTableColumnName(deleted.tableName) === normalizeTableColumnName(tableName) &&
    normalizeTableColumnName(deleted.columnName) === normalizeTableColumnName(columnName)
  )
}

function formulaReferencesChangedTable(formula: string, changedTableNames: ReadonlySet<string>): boolean {
  if (changedTableNames.size === 0) {
    return false
  }
  try {
    return formulaNodeReferencesChangedTable(parseFormula(formula), changedTableNames)
  } catch {
    return false
  }
}

function formulaNodeReferencesChangedTable(node: FormulaNode, changedTableNames: ReadonlySet<string>): boolean {
  switch (node.kind) {
    case 'StructuredRef':
      return structuredReferenceTouchesChangedTable(node.tableName, changedTableNames)
    case 'ArrayConstant':
      return node.rows.some((row) => row.some((entry) => formulaNodeReferencesChangedTable(entry, changedTableNames)))
    case 'UnaryExpr':
      return formulaNodeReferencesChangedTable(node.argument, changedTableNames)
    case 'BinaryExpr':
      return (
        formulaNodeReferencesChangedTable(node.left, changedTableNames) || formulaNodeReferencesChangedTable(node.right, changedTableNames)
      )
    case 'CallExpr':
      return node.args.some((arg) => formulaNodeReferencesChangedTable(arg, changedTableNames))
    case 'InvokeExpr':
      return (
        formulaNodeReferencesChangedTable(node.callee, changedTableNames) ||
        node.args.some((arg) => formulaNodeReferencesChangedTable(arg, changedTableNames))
      )
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'CellRef':
    case 'SpillRef':
    case 'ColumnRef':
    case 'RowRef':
    case 'RangeRef':
      return false
  }
}

function structuredReferenceTouchesChangedTable(tableName: string, changedTableNames: ReadonlySet<string>): boolean {
  if (changedTableNames.has(tableName)) {
    return true
  }
  const normalizedTableName = normalizeTableColumnName(tableName)
  for (const changedTableName of changedTableNames) {
    if (normalizeTableColumnName(changedTableName) === normalizedTableName) {
      return true
    }
  }
  return false
}

export function rewriteWorkbookMetadataForStructuralTransform(
  args: StructureMetadataRewriteArgs,
  sheetName: string,
  transform: StructuralAxisTransform,
): {
  changedTableNames: Set<string>
  tableHeaderCellWrites: StructuralTableHeaderCellWrite[]
  deletedTableColumns: DeletedTableColumnReference[]
} {
  const workbook = args.state.workbook
  const changedTableNames = new Set<string>()
  const tableHeaderCellWrites: StructuralTableHeaderCellWrite[] = []
  const deletedTableColumns: DeletedTableColumnReference[] = []
  workbook.listTables().forEach((table) => {
    if (table.sheetName !== sheetName) {
      return
    }
    const rewrite = rewriteTableForStructuralTransform(table, transform)
    if (!rewrite) {
      changedTableNames.add(table.name)
      deletedTableColumns.push(...table.columnNames.map((columnName) => ({ tableName: table.name, columnName })))
      workbook.deleteTable(table.name)
      return
    }
    changedTableNames.add(table.name)
    tableHeaderCellWrites.push(...rewrite.headerCellWrites)
    deletedTableColumns.push(...rewrite.deletedColumnNames.map((columnName) => ({ tableName: table.name, columnName })))
    workbook.setTable(rewrite.table)
  })
  const rewrittenMergeRanges: CellRangeRef[] = []
  workbook.listMergeRanges(sheetName).forEach((merge) => {
    const range = rewriteMetadataRangeForStructuralTransform(merge, transform)
    if (!range) {
      return
    }
    rewrittenMergeRanges.push(range)
  })
  workbook.setMergeRanges(sheetName, rewrittenMergeRanges)
  workbook.listFilters(sheetName).forEach((filter) => {
    const range = rewriteMetadataRangeForStructuralTransform(filter.range, transform)
    workbook.deleteFilter(sheetName, filter.range)
    if (range) {
      workbook.setFilter(sheetName, range)
    }
  })
  workbook.listSorts(sheetName).forEach((sort) => {
    const range = rewriteMetadataRangeForStructuralTransform(sort.range, transform)
    workbook.deleteSort(sheetName, sort.range)
    if (!range) {
      return
    }
    workbook.setSort(
      sheetName,
      range,
      sort.keys.map((key) => ({
        ...key,
        keyAddress: rewriteMetadataAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
      })),
    )
  })
  workbook.listDataValidations(sheetName).forEach((validation) => {
    const range = rewriteMetadataRangeForStructuralTransform(validation.range, transform)
    workbook.deleteDataValidation(sheetName, validation.range)
    if (!range) {
      return
    }
    const nextValidation = structuredClone(validation)
    nextValidation.range = range
    if (nextValidation.rule.kind === 'list' && nextValidation.rule.source) {
      switch (nextValidation.rule.source.kind) {
        case 'cell-ref': {
          if (nextValidation.rule.source.sheetName !== sheetName) {
            break
          }
          const nextAddress = rewriteMetadataAddressForStructuralTransform(nextValidation.rule.source.address, transform)
          if (!nextAddress) {
            return
          }
          nextValidation.rule.source.address = nextAddress
          break
        }
        case 'range-ref': {
          if (nextValidation.rule.source.sheetName !== sheetName) {
            break
          }
          const nextSourceRange = rewriteMetadataRangeForStructuralTransform(nextValidation.rule.source, transform)
          if (!nextSourceRange) {
            return
          }
          nextValidation.rule.source = nextSourceRange
          break
        }
        case 'named-range':
        case 'structured-ref':
          break
      }
    }
    workbook.setDataValidation(nextValidation)
  })
  workbook.listConditionalFormats(sheetName).forEach((format) => {
    const nextFormat = rewriteMetadataRangeRecord(format, transform)
    workbook.deleteConditionalFormat(format.id)
    if (!nextFormat) {
      return
    }
    workbook.setConditionalFormat(nextFormat)
  })
  const conditionalFormatArtifacts = workbook.getConditionalFormatArtifacts(sheetName)
  if (conditionalFormatArtifacts) {
    const nextArtifactsXml = rewriteConditionalFormatArtifactXmlForStructuralTransform(conditionalFormatArtifacts.xml, transform)
    workbook.deleteConditionalFormatArtifacts(sheetName)
    if (nextArtifactsXml) {
      workbook.setConditionalFormatArtifacts(sheetName, { xml: nextArtifactsXml })
    }
  }
  workbook.listRangeProtections(sheetName).forEach((protection) => {
    const nextProtection = rewriteMetadataRangeRecord(protection, transform)
    workbook.deleteRangeProtection(protection.id)
    if (!nextProtection) {
      return
    }
    workbook.setRangeProtection(nextProtection)
  })
  workbook.listCommentThreads(sheetName).forEach((thread) => {
    const nextAddress = rewriteMetadataAddressForStructuralTransform(thread.address, transform)
    workbook.deleteCommentThread(sheetName, thread.address)
    if (!nextAddress) {
      return
    }
    workbook.setCommentThread({
      ...thread,
      address: nextAddress,
    })
  })
  rewriteLegacyCommentVmlForStructuralTransform({ workbook, sheetName, transform })
  rewriteThreadedCommentArtifactsForStructuralTransform({ workbook, sheetName, transform })
  workbook.listNotes(sheetName).forEach((note) => {
    const nextAddress = rewriteMetadataAddressForStructuralTransform(note.address, transform)
    workbook.deleteNote(sheetName, note.address)
    if (!nextAddress) {
      return
    }
    workbook.setNote({
      ...note,
      address: nextAddress,
    })
  })
  workbook.listHyperlinks(sheetName).forEach((hyperlink) => {
    const nextAddress = rewriteMetadataAddressForStructuralTransform(hyperlink.address, transform)
    workbook.deleteHyperlink(sheetName, hyperlink.address)
    if (!nextAddress) {
      return
    }
    workbook.setHyperlink({
      ...hyperlink,
      address: nextAddress,
    })
  })
  const sheet = workbook.getSheet(sheetName)
  if (sheet?.arrayFormulas) {
    const arrayFormulas = rewriteArrayFormulasForStructuralTransform(sheetName, sheet.arrayFormulas, transform)
    if (arrayFormulas) {
      sheet.arrayFormulas = arrayFormulas
    } else {
      delete sheet.arrayFormulas
    }
  }
  if (sheet?.dataTableFormulas) {
    const dataTableFormulas = rewriteDataTableFormulasForStructuralTransform(sheet.dataTableFormulas, transform)
    if (dataTableFormulas) {
      sheet.dataTableFormulas = dataTableFormulas
    } else {
      delete sheet.dataTableFormulas
    }
  }
  const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = []
  const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = []
  workbook.listStyleRanges(sheetName).forEach((record) => {
    const nextRecord = rewriteMetadataRangeRecord(record, transform)
    if (nextRecord) {
      rewrittenStyleRanges.push(nextRecord)
    }
  })
  workbook.setStyleRanges(sheetName, rewrittenStyleRanges)
  workbook.listFormatRanges(sheetName).forEach((record) => {
    const nextRecord = rewriteMetadataRangeRecord(record, transform)
    if (nextRecord) {
      rewrittenFormatRanges.push(nextRecord)
    }
  })
  workbook.setFormatRanges(sheetName, rewrittenFormatRanges)
  const freezePane = workbook.getFreezePane(sheetName)
  if (freezePane) {
    const nextRows = transform.axis === 'row' ? mapStructuralBoundary(freezePane.rows, transform) : freezePane.rows
    const nextCols = transform.axis === 'column' ? mapStructuralBoundary(freezePane.cols, transform) : freezePane.cols
    if (nextRows <= 0 && nextCols <= 0) {
      workbook.clearFreezePane(sheetName)
    } else {
      workbook.setFreezePane(sheetName, nextRows, nextCols)
    }
  }
  workbook.listPivots().forEach((pivot) => {
    const nextAddress =
      pivot.sheetName === sheetName ? rewriteMetadataAddressForStructuralTransform(pivot.address, transform) : pivot.address
    const nextSource =
      pivot.source?.sheetName === sheetName ? rewriteMetadataRangeForStructuralTransform(pivot.source, transform) : pivot.source
    if (!nextAddress || (pivot.source && !nextSource)) {
      args.clearOwnedPivot(pivot)
      workbook.deletePivot(pivot.sheetName, pivot.address)
      return
    }
    if (nextAddress !== pivot.address) {
      args.clearOwnedPivot(pivot)
      workbook.deletePivot(pivot.sheetName, pivot.address)
    }
    workbook.setPivot({
      ...pivot,
      address: nextAddress,
      ...(nextSource ? { source: nextSource } : {}),
    })
  })
  workbook.listCharts().forEach((chart) => {
    const nextAddress =
      chart.sheetName === sheetName ? rewriteMetadataAddressForStructuralTransform(chart.address, transform) : chart.address
    const nextSource =
      chart.source.sheetName === sheetName ? rewriteMetadataRangeForStructuralTransform(chart.source, transform) : chart.source
    const nextAnchor =
      chart.sheetName === sheetName && chart.anchor ? rewriteChartAnchorForStructuralTransform(chart.anchor, transform) : chart.anchor
    if (!nextAddress || !nextSource || (chart.anchor && !nextAnchor)) {
      workbook.deleteChart(chart.id)
      return
    }
    const nextGeometry = chartGeometryFromAnchor(nextAnchor)
    workbook.setChart({
      ...chart,
      address: nextGeometry?.address ?? nextAddress,
      rows: nextGeometry?.rows ?? chart.rows,
      cols: nextGeometry?.cols ?? chart.cols,
      ...(nextAnchor !== undefined ? { anchor: nextAnchor } : {}),
      source: nextSource,
    })
  })
  workbook.listImages().forEach((image) => {
    if (image.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteMetadataAddressForStructuralTransform(image.address, transform)
    if (!nextAddress) {
      workbook.deleteImage(image.id)
      return
    }
    workbook.setImage({
      ...image,
      address: nextAddress,
    })
  })
  workbook.listShapes().forEach((shape) => {
    if (shape.sheetName !== sheetName) {
      return
    }
    const nextAddress = rewriteMetadataAddressForStructuralTransform(shape.address, transform)
    if (!nextAddress) {
      workbook.deleteShape(shape.id)
      return
    }
    workbook.setShape({
      ...shape,
      address: nextAddress,
    })
  })
  return { changedTableNames, tableHeaderCellWrites, deletedTableColumns }
}

function rewriteMetadataRangeForStructuralTransform<T extends MetadataRangeLike>(
  range: T,
  transform: StructuralAxisTransform,
): T | undefined {
  const rewritten = rewriteRangeForStructuralTransform(range.startAddress, range.endAddress, transform)
  if (!rewritten) {
    return undefined
  }
  const clipped = clipMetadataRangeToSheetGrid(range.sheetName, rewritten.startAddress, rewritten.endAddress)
  return clipped ? withRewrittenMetadataRange(range, clipped) : undefined
}

function withRewrittenMetadataRange<T extends MetadataRangeLike>(range: T, rewritten: CellRangeRef): T {
  return {
    ...range,
    startAddress: rewritten.startAddress,
    endAddress: rewritten.endAddress,
  }
}

function rewriteTableForStructuralTransform(
  table: WorkbookTableRecord,
  transform: StructuralAxisTransform,
): { table: WorkbookTableRecord; headerCellWrites: StructuralTableHeaderCellWrite[]; deletedColumnNames: string[] } | undefined {
  const rewrittenRange = rewriteMetadataRangeForStructuralTransform(table, transform)
  const range = rewrittenRange ? ensureMinimumTableDataBodyRange(rewrittenRange, table) : undefined
  if (!range) {
    return undefined
  }
  if (transform.axis !== 'column') {
    return { table: range, headerCellWrites: [], deletedColumnNames: [] }
  }
  const previousStart = parseUnboundedMetadataCellAddress(table.startAddress)
  const previousEnd = parseUnboundedMetadataCellAddress(table.endAddress)
  const nextStart = parseUnboundedMetadataCellAddress(range.startAddress)
  const nextEnd = parseUnboundedMetadataCellAddress(range.endAddress)
  if (!previousStart || !previousEnd || !nextStart || !nextEnd) {
    throw new Error('Invalid table metadata reference')
  }
  const previousStartCol = Math.min(previousStart[1], previousEnd[1])
  const previousEndCol = Math.max(previousStart[1], previousEnd[1])
  const nextStartCol = Math.min(nextStart[1], nextEnd[1])
  const nextEndCol = Math.max(nextStart[1], nextEnd[1])
  const previousWidth = previousEndCol - previousStartCol + 1
  const nextWidth = nextEndCol - nextStartCol + 1
  const nextColumnNames: Array<string | undefined> = Array.from({ length: nextWidth }, () => undefined)
  const nextColumns: Array<WorkbookTableColumnRecord | undefined> | undefined = table.columns
    ? Array.from({ length: nextWidth }, () => undefined)
    : undefined
  const usedColumnNames = new Set<string>()
  const deletedColumnNames: string[] = []

  for (let previousIndex = 0; previousIndex < previousWidth; previousIndex += 1) {
    const mappedCol = mapTableColumnPointForStructuralTransform(previousStartCol + previousIndex, transform)
    if (mappedCol === undefined || mappedCol < nextStartCol || mappedCol > nextEndCol) {
      deletedColumnNames.push(table.columnNames[previousIndex] ?? nextGeneratedTableColumnName(usedColumnNames))
      continue
    }
    const nextIndex = mappedCol - nextStartCol
    const name = table.columnNames[previousIndex] ?? nextGeneratedTableColumnName(usedColumnNames)
    usedColumnNames.add(normalizeTableColumnName(name))
    nextColumnNames[nextIndex] = name
    if (nextColumns) {
      const previousColumn = table.columns?.[previousIndex]
      nextColumns[nextIndex] = previousColumn ? { ...previousColumn, name } : { name }
    }
  }

  const headerCellWrites: StructuralTableHeaderCellWrite[] = []
  const headerRow = Math.min(previousStart[0], previousEnd[0])
  for (let nextIndex = 0; nextIndex < nextWidth; nextIndex += 1) {
    if (nextColumnNames[nextIndex] !== undefined) {
      continue
    }
    const name = nextGeneratedTableColumnName(usedColumnNames)
    usedColumnNames.add(normalizeTableColumnName(name))
    nextColumnNames[nextIndex] = name
    if (nextColumns) {
      nextColumns[nextIndex] = { name }
    }
    const col = nextStartCol + nextIndex
    if (table.headerRow && transform.kind === 'insert' && col >= transform.start && col < transform.start + transform.count) {
      headerCellWrites.push({ sheetName: range.sheetName, row: headerRow, col, value: name })
    }
  }
  const finalizedColumnNames = nextColumnNames.map((name) => {
    if (name === undefined) {
      throw new Error('Missing table column name after structural rewrite')
    }
    return name
  })
  const finalizedColumns = nextColumns?.map((column) => {
    if (column === undefined) {
      throw new Error('Missing table column record after structural rewrite')
    }
    return column
  })

  return {
    table: {
      ...range,
      columnNames: finalizedColumnNames,
      ...(finalizedColumns ? { columns: finalizedColumns } : {}),
    },
    headerCellWrites,
    deletedColumnNames,
  }
}

function ensureMinimumTableDataBodyRange<T extends MetadataRangeLike>(range: T, table: WorkbookTableRecord): T | undefined {
  const start = parseUnboundedMetadataCellAddress(range.startAddress)
  const end = parseUnboundedMetadataCellAddress(range.endAddress)
  if (!start || !end) {
    throw new Error('Invalid table metadata reference')
  }
  const startRow = Math.min(start[0], end[0])
  const endRow = Math.max(start[0], end[0])
  const endCol = Math.max(start[1], end[1])
  const minimumRowCount = (table.headerRow ? 1 : 0) + (table.totalsRow ? 1 : 0) + 1
  if (endRow - startRow + 1 >= minimumRowCount) {
    return range
  }
  const nextEndRow = startRow + minimumRowCount - 1
  if (nextEndRow >= MAX_ROWS) {
    return range
  }
  return withRewrittenMetadataRange(range, {
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: formatAddress(nextEndRow, endCol),
  })
}

function mapTableColumnPointForStructuralTransform(index: number, transform: StructuralAxisTransform): number | undefined {
  switch (transform.kind) {
    case 'insert':
      return index >= transform.start ? index + transform.count : index
    case 'delete':
      if (index < transform.start) {
        return index
      }
      if (index >= transform.start + transform.count) {
        return index - transform.count
      }
      return undefined
    case 'move':
      if (transform.target < transform.start) {
        if (index >= transform.target && index < transform.start) {
          return index + transform.count
        }
      } else if (transform.target > transform.start) {
        if (index >= transform.start + transform.count && index < transform.target + transform.count) {
          return index - transform.count
        }
      }
      if (index >= transform.start && index < transform.start + transform.count) {
        return transform.target + (index - transform.start)
      }
      return index
  }
}

function rewriteMetadataRangeRecord<T extends { readonly range: MetadataRangeLike }>(
  record: T,
  transform: StructuralAxisTransform,
): T | undefined {
  const range = rewriteMetadataRangeForStructuralTransform(record.range, transform)
  return range ? { ...record, range } : undefined
}

function rewriteMetadataAddressForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseUnboundedMetadataCellAddress(rewritten)
  if (!parsed) {
    throw new Error('Invalid metadata reference')
  }
  if (parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function clipMetadataRangeToSheetGrid(sheetName: string, startAddress: string, endAddress: string): CellRangeRef | undefined {
  const start = parseUnboundedMetadataCellAddress(startAddress)
  const end = parseUnboundedMetadataCellAddress(endAddress)
  if (!start || !end) {
    throw new Error('Invalid metadata reference')
  }
  const startRow = Math.min(start[0], end[0])
  const endRow = Math.min(MAX_ROWS - 1, Math.max(start[0], end[0]))
  const startCol = Math.min(start[1], end[1])
  const endCol = Math.min(MAX_COLS - 1, Math.max(start[1], end[1]))
  if (startRow > endRow || startCol > endCol) {
    return undefined
  }
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  }
}

function parseUnboundedMetadataCellAddress(address: string): [number, number] | undefined {
  const match = METADATA_CELL_REF_RE.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
