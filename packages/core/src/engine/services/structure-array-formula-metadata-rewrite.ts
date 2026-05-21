import type { WorkbookSheetArrayFormulaSnapshot, WorkbookSheetArrayFormulasSnapshot } from '@bilig/protocol'
import {
  columnToIndex,
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'

const METADATA_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i

function readXmlAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function escapeXmlText(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;')
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/gu, '&quot;').replace(/'/gu, '&apos;')
}

function rewriteFormulaXmlAttribute(formulaXml: string, attributeName: string, nextValue: string): string | undefined {
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

function readFormulaXmlAttribute(formulaXml: string, attributeName: string): string | undefined {
  const value = readXmlAttribute(formulaXml, attributeName)
  return value === null ? undefined : decodeXmlText(value)
}

function readArrayFormulaSource(formulaXml: string): string | undefined {
  const match = /<f\b[^>]*>([\s\S]*?)<\/f>/u.exec(formulaXml)
  return match ? decodeXmlText(match[1] ?? '') : undefined
}

function parseMetadataCellAddress(address: string): [number, number] | undefined {
  const match = METADATA_CELL_REF_RE.exec(address)
  return match ? [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())] : undefined
}

function arrayFormulaAddressAxisIndex(address: string, axis: StructuralAxisTransform['axis']): number | undefined {
  const parsed = parseMetadataCellAddress(address)
  return parsed ? parsed[axis === 'row' ? 0 : 1] : undefined
}

function arrayFormulaRefAxisBounds(
  ref: string,
  axis: StructuralAxisTransform['axis'],
): { readonly min: number; readonly max: number } | undefined {
  const [start, end, extra] = ref.split(':')
  if (!start || extra !== undefined) {
    return undefined
  }
  const startIndex = arrayFormulaAddressAxisIndex(start, axis)
  const endIndex = arrayFormulaAddressAxisIndex(end ?? start, axis)
  if (startIndex === undefined || endIndex === undefined) {
    return undefined
  }
  return { min: Math.min(startIndex, endIndex), max: Math.max(startIndex, endIndex) }
}

function isProvenArrayFormulaStructuralTransform(
  formula: WorkbookSheetArrayFormulaSnapshot,
  ref: string,
  transform: StructuralAxisTransform,
): boolean {
  const bounds = arrayFormulaStructuralAxisBounds(formula, ref, transform.axis)
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

function arrayFormulaStructuralAxisBounds(
  formula: WorkbookSheetArrayFormulaSnapshot,
  ref: string,
  axis: StructuralAxisTransform['axis'],
): { readonly min: number; readonly max: number } | undefined {
  const refBounds = arrayFormulaRefAxisBounds(ref, axis)
  const addressIndex = arrayFormulaAddressAxisIndex(formula.address, axis)
  if (!refBounds || addressIndex === undefined) {
    return undefined
  }
  return { min: Math.min(refBounds.min, addressIndex), max: Math.max(refBounds.max, addressIndex) }
}

function structuralTransformIsAfterArrayFormulaMetadata(
  formula: WorkbookSheetArrayFormulaSnapshot,
  ref: string,
  transform: StructuralAxisTransform,
): boolean {
  const bounds = arrayFormulaStructuralAxisBounds(formula, ref, transform.axis)
  if (!bounds) {
    return false
  }
  switch (transform.kind) {
    case 'insert':
    case 'delete':
      return transform.start > bounds.max
    case 'move':
      return transform.start > bounds.max && transform.target > bounds.max
  }
}

function structuralTransformMovesArrayFormulaMetadata(
  formula: WorkbookSheetArrayFormulaSnapshot,
  ref: string,
  transform: StructuralAxisTransform,
): boolean {
  const bounds = arrayFormulaStructuralAxisBounds(formula, ref, transform.axis)
  if (!bounds) {
    return false
  }
  switch (transform.kind) {
    case 'insert':
      return transform.start <= bounds.min
    case 'delete':
    case 'move':
      return false
  }
}

function rewriteArrayFormulaSource(formulaXml: string, nextFormula: string): string | undefined {
  if (!/<f\b[^>]*>[\s\S]*?<\/f>/u.test(formulaXml)) {
    return undefined
  }
  return formulaXml.replace(/(<f\b[^>]*>)[\s\S]*?(<\/f>)/u, (_source: string, openTag: string, closeTag: string) => {
    return `${openTag}${escapeXmlText(nextFormula)}${closeTag}`
  })
}

function rewriteArrayFormulaAddress(value: string, transform: StructuralAxisTransform): string | undefined {
  try {
    return rewriteAddressForStructuralTransform(value, transform)
  } catch {
    return undefined
  }
}

function rewriteArrayFormulaRef(value: string, transform: StructuralAxisTransform): string | undefined {
  const [start, end, extra] = value.split(':')
  if (!start || extra !== undefined) {
    return undefined
  }
  if (!end) {
    return rewriteArrayFormulaAddress(start, transform)
  }
  try {
    const rewritten = rewriteRangeForStructuralTransform(start, end, transform)
    return rewritten ? `${rewritten.startAddress}:${rewritten.endAddress}` : undefined
  } catch {
    return undefined
  }
}

function rewriteArrayFormulaForStructuralTransform(
  sheetName: string,
  formula: WorkbookSheetArrayFormulaSnapshot,
  transform: StructuralAxisTransform,
): WorkbookSheetArrayFormulaSnapshot | undefined {
  const address = rewriteArrayFormulaAddress(formula.address, transform)
  const ref = readFormulaXmlAttribute(formula.formulaXml, 'ref')
  const source = readArrayFormulaSource(formula.formulaXml)
  if (!address || !ref || source === undefined) {
    return undefined
  }
  if (!isProvenArrayFormulaStructuralTransform(formula, ref, transform)) {
    return undefined
  }
  if (structuralTransformIsAfterArrayFormulaMetadata(formula, ref, transform)) {
    return { ...formula }
  }
  if (!structuralTransformMovesArrayFormulaMetadata(formula, ref, transform)) {
    return undefined
  }

  const nextRef = rewriteArrayFormulaRef(ref, transform)
  if (!nextRef) {
    return undefined
  }

  let nextSource: string
  try {
    nextSource = rewriteFormulaForStructuralTransform(source, sheetName, sheetName, transform)
  } catch {
    return undefined
  }

  const withRef = rewriteFormulaXmlAttribute(formula.formulaXml, 'ref', nextRef)
  const formulaXml = withRef ? rewriteArrayFormulaSource(withRef, nextSource) : undefined
  return formulaXml ? { address, formulaXml } : undefined
}

export function rewriteArrayFormulasForStructuralTransform(
  sheetName: string,
  formulas: WorkbookSheetArrayFormulasSnapshot | undefined,
  transform: StructuralAxisTransform,
): WorkbookSheetArrayFormulasSnapshot | undefined {
  const rewritten = formulas?.formulas.flatMap((formula) => {
    const nextFormula = rewriteArrayFormulaForStructuralTransform(sheetName, formula, transform)
    return nextFormula ? [nextFormula] : []
  })
  return rewritten && rewritten.length > 0 ? { formulas: rewritten } : undefined
}
