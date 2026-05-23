import type { WorkbookSheetDataTableFormulaSnapshot, WorkbookSheetDataTableFormulasSnapshot } from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'

interface DataTableStructuralGeometry {
  readonly address: string
  readonly ref: string
  readonly rowInput: string
  readonly columnInput?: string
}

const METADATA_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i

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

export function rewriteDataTableFormulasForStructuralTransform(
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

const conditionalFormattingElementName = '(?:[A-Za-z_][\\w.-]*:)?conditionalFormatting'
const conditionalFormattingBlockRegex = new RegExp(
  `<${conditionalFormattingElementName}\\b[^>]*>[\\s\\S]*?<\\/${conditionalFormattingElementName}>|<${conditionalFormattingElementName}\\b[^>]*/>`,
  'gu',
)
const conditionalFormattingOpeningTagRegex = new RegExp(`^<${conditionalFormattingElementName}\\b[^>]*\\/?>`, 'u')

function rewriteConditionalFormatArtifactBlock(block: string, transform: StructuralAxisTransform): string | undefined {
  const openingTag = conditionalFormattingOpeningTagRegex.exec(block)?.[0]
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

export function rewriteConditionalFormatArtifactXmlForStructuralTransform(
  xml: string,
  transform: StructuralAxisTransform,
): string | undefined {
  let matchedBlock = false
  const rewrittenXml = xml.replace(conditionalFormattingBlockRegex, (block) => {
    matchedBlock = true
    return rewriteConditionalFormatArtifactBlock(block, transform) ?? ''
  })
  if (!matchedBlock) {
    return xml
  }
  return rewrittenXml.trim().length > 0 ? rewrittenXml : undefined
}

function parseUnboundedMetadataCellAddress(address: string): [number, number] | undefined {
  const match = METADATA_CELL_REF_RE.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
