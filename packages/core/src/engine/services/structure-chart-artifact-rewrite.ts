import {
  parseFormula,
  renameFormulaSheetReferences,
  rewriteFormulaForStructuralTransform,
  serializeFormula,
  type FormulaNode,
  type StructuralAxisKind,
  type StructuralAxisTransform,
} from '@bilig/formula'
import { ErrorCode, type WorkbookDrawingArtifactsSnapshot, type WorkbookPreservedPackagePartSnapshot } from '@bilig/protocol'
import type { WorkbookPreservedMetadataRecord } from '../../workbook-metadata-types.js'
import type { WorkbookStore } from '../../workbook-store.js'

const binaryChunkSize = 0x8000
const chartPackageFormulaOwnerSheetName = '__bilig_chart_package_artifact__'
const chartPartPathPattern = /^xl\/charts\/chart\d+\.xml$/u

export function rewritePreservedChartPackageArtifactsForStructuralTransform(
  metadata: WorkbookPreservedMetadataRecord,
  sheetName: string,
  transform: StructuralAxisTransform,
): WorkbookPreservedMetadataRecord | undefined {
  const chartArtifacts = metadata.chartArtifacts
  if (!chartArtifacts || chartArtifacts.parts.length === 0) {
    return undefined
  }

  let changed = false
  const parts = chartArtifacts.parts.map((part) => {
    if (!chartPartPathPattern.test(normalizeZipPath(part.path))) {
      return structuredClone(part)
    }
    const nextPart = rewriteChartPackagePartForStructuralTransform(part, sheetName, transform)
    changed ||= nextPart.dataBase64 !== part.dataBase64 || nextPart.byteLength !== part.byteLength
    return nextPart
  })

  return changed
    ? {
        ...metadata,
        chartArtifacts: {
          ...chartArtifacts,
          parts,
        },
      }
    : undefined
}

export function preservedChartPackageArtifactsTouchStructuralDelete(
  metadata: WorkbookPreservedMetadataRecord,
  sheetName: string,
  axis: StructuralAxisKind,
  start: number,
  count: number,
): boolean {
  const chartArtifacts = metadata.chartArtifacts
  if (!chartArtifacts) {
    return false
  }

  return chartArtifacts.parts.some((part) => {
    if (!chartPartPathPattern.test(normalizeZipPath(part.path))) {
      return false
    }
    const bytes = decodeBase64(part.dataBase64)
    if (bytes.byteLength !== part.byteLength) {
      return false
    }
    const xml = new TextDecoder().decode(bytes)
    return chartFormulaTexts(xml).some((formula) => chartFormulaWouldRewriteForDelete(formula, sheetName, axis, start, count))
  })
}

export function rewriteDrawingChartPackageArtifactsForStructuralTransform(args: {
  readonly workbook: WorkbookStore
  readonly sheetName: string
  readonly transform: StructuralAxisTransform
}): void {
  const drawingArtifacts = args.workbook.getDrawingArtifacts()
  if (!drawingArtifacts || drawingArtifacts.parts.length === 0) {
    return
  }

  let changed = false
  const parts = drawingArtifacts.parts.map((part) => {
    if (!chartPartPathPattern.test(normalizeZipPath(part.path))) {
      return structuredClone(part)
    }
    const nextPart = rewriteChartPackagePartForStructuralTransform(part, args.sheetName, args.transform)
    changed ||= nextPart.dataBase64 !== part.dataBase64 || nextPart.byteLength !== part.byteLength
    return nextPart
  })

  if (changed) {
    args.workbook.setDrawingArtifacts({ ...drawingArtifacts, parts })
  }
}

export function drawingChartPackageArtifactsTouchStructuralDelete(
  workbook: WorkbookStore,
  sheetName: string,
  axis: StructuralAxisKind,
  start: number,
  count: number,
): boolean {
  const drawingArtifacts = workbook.getDrawingArtifacts()
  if (!drawingArtifacts) {
    return false
  }
  return drawingArtifacts.parts.some((part) => {
    if (!chartPartPathPattern.test(normalizeZipPath(part.path))) {
      return false
    }
    const bytes = decodeBase64(part.dataBase64)
    if (bytes.byteLength !== part.byteLength) {
      return false
    }
    const xml = new TextDecoder().decode(bytes)
    return chartFormulaTexts(xml).some((formula) => chartFormulaWouldRewriteForDelete(formula, sheetName, axis, start, count))
  })
}

export function renamePreservedChartPackageArtifactsSheetReferences(
  metadata: WorkbookPreservedMetadataRecord,
  oldSheetName: string,
  newSheetName: string,
): WorkbookPreservedMetadataRecord | undefined {
  const chartArtifacts = metadata.chartArtifacts
  if (!chartArtifacts || chartArtifacts.parts.length === 0) {
    return undefined
  }

  const nextChartArtifacts = renameChartPackageArtifactsSheetReferences(chartArtifacts, oldSheetName, newSheetName)
  return nextChartArtifacts
    ? {
        ...metadata,
        chartArtifacts: nextChartArtifacts,
      }
    : undefined
}

export function renameDrawingChartPackageArtifactsSheetReferences(
  drawingArtifacts: WorkbookDrawingArtifactsSnapshot | undefined,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDrawingArtifactsSnapshot | undefined {
  if (!drawingArtifacts || drawingArtifacts.parts.length === 0) {
    return undefined
  }
  return renameChartPackageArtifactsSheetReferences(drawingArtifacts, oldSheetName, newSheetName)
}

export function rewritePreservedChartPackageArtifactsForSheetDeletion(
  metadata: WorkbookPreservedMetadataRecord,
  deletedSheetName: string,
): WorkbookPreservedMetadataRecord | undefined {
  const chartArtifacts = metadata.chartArtifacts
  if (!chartArtifacts || chartArtifacts.parts.length === 0) {
    return undefined
  }

  const nextChartArtifacts = rewriteChartPackageArtifactsForSheetDeletion(chartArtifacts, deletedSheetName)
  return nextChartArtifacts
    ? {
        ...metadata,
        chartArtifacts: nextChartArtifacts,
      }
    : undefined
}

export function rewriteDrawingChartPackageArtifactsForSheetDeletion(
  drawingArtifacts: WorkbookDrawingArtifactsSnapshot | undefined,
  deletedSheetName: string,
): WorkbookDrawingArtifactsSnapshot | undefined {
  if (!drawingArtifacts || drawingArtifacts.parts.length === 0) {
    return undefined
  }
  return rewriteChartPackageArtifactsForSheetDeletion(drawingArtifacts, deletedSheetName)
}

function renameChartPackageArtifactsSheetReferences(
  artifacts: WorkbookDrawingArtifactsSnapshot,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDrawingArtifactsSnapshot | undefined {
  let changed = false
  const parts = artifacts.parts.map((part) => {
    if (!chartPartPathPattern.test(normalizeZipPath(part.path))) {
      return structuredClone(part)
    }
    const nextPart = renameChartPackagePartSheetReferences(part, oldSheetName, newSheetName)
    changed ||= nextPart.dataBase64 !== part.dataBase64 || nextPart.byteLength !== part.byteLength
    return nextPart
  })

  return changed ? { ...artifacts, parts } : undefined
}

function rewriteChartPackageArtifactsForSheetDeletion(
  artifacts: WorkbookDrawingArtifactsSnapshot,
  deletedSheetName: string,
): WorkbookDrawingArtifactsSnapshot | undefined {
  let changed = false
  const parts = artifacts.parts.map((part) => {
    if (!chartPartPathPattern.test(normalizeZipPath(part.path))) {
      return structuredClone(part)
    }
    const nextPart = rewriteChartPackagePartForSheetDeletion(part, deletedSheetName)
    changed ||= nextPart.dataBase64 !== part.dataBase64 || nextPart.byteLength !== part.byteLength
    return nextPart
  })

  return changed ? { ...artifacts, parts } : undefined
}

function rewriteChartPackagePartForStructuralTransform(
  part: WorkbookPreservedPackagePartSnapshot,
  sheetName: string,
  transform: StructuralAxisTransform,
): WorkbookPreservedPackagePartSnapshot {
  const bytes = decodeBase64(part.dataBase64)
  if (bytes.byteLength !== part.byteLength) {
    return structuredClone(part)
  }
  const xml = new TextDecoder().decode(bytes)
  const nextXml = rewriteChartFormulaXmlForStructuralTransform(xml, sheetName, transform)
  if (nextXml === xml) {
    return structuredClone(part)
  }
  const nextBytes = new TextEncoder().encode(nextXml)
  return {
    ...part,
    dataBase64: encodeBase64(nextBytes),
    byteLength: nextBytes.byteLength,
  }
}

function rewriteChartPackagePartForSheetDeletion(
  part: WorkbookPreservedPackagePartSnapshot,
  deletedSheetName: string,
): WorkbookPreservedPackagePartSnapshot {
  const bytes = decodeBase64(part.dataBase64)
  if (bytes.byteLength !== part.byteLength) {
    return structuredClone(part)
  }
  const xml = new TextDecoder().decode(bytes)
  const nextXml = rewriteChartFormulaXmlForSheetDeletion(xml, deletedSheetName)
  if (nextXml === xml) {
    return structuredClone(part)
  }
  const nextBytes = new TextEncoder().encode(nextXml)
  return {
    ...part,
    dataBase64: encodeBase64(nextBytes),
    byteLength: nextBytes.byteLength,
  }
}

function renameChartPackagePartSheetReferences(
  part: WorkbookPreservedPackagePartSnapshot,
  oldSheetName: string,
  newSheetName: string,
): WorkbookPreservedPackagePartSnapshot {
  const bytes = decodeBase64(part.dataBase64)
  if (bytes.byteLength !== part.byteLength) {
    return structuredClone(part)
  }
  const xml = new TextDecoder().decode(bytes)
  const nextXml = renameChartFormulaXmlSheetReferences(xml, oldSheetName, newSheetName)
  if (nextXml === xml) {
    return structuredClone(part)
  }
  const nextBytes = new TextEncoder().encode(nextXml)
  return {
    ...part,
    dataBase64: encodeBase64(nextBytes),
    byteLength: nextBytes.byteLength,
  }
}

function rewriteChartFormulaXmlForStructuralTransform(xml: string, sheetName: string, transform: StructuralAxisTransform): string {
  return xml.replace(
    /(<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>)([\s\S]*?)(<\/(?:[A-Za-z_][\w.-]*:)?f>)/gu,
    (source: string, open: string, text: string, close: string) => {
      const formula = decodeXmlText(text)
      const nextFormula = rewriteChartFormulaForStructuralTransform(formula, sheetName, transform)
      return nextFormula === undefined || nextFormula === formula ? source : `${open}${escapeXmlText(nextFormula)}${close}`
    },
  )
}

function rewriteChartFormulaXmlForSheetDeletion(xml: string, deletedSheetName: string): string {
  let changed = false
  const rewrittenXml = xml.replace(
    /(<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>)([\s\S]*?)(<\/(?:[A-Za-z_][\w.-]*:)?f>)/gu,
    (source: string, open: string, text: string, close: string) => {
      const formula = decodeXmlText(text)
      const nextFormula = rewriteChartFormulaForSheetDeletion(formula, deletedSheetName)
      if (nextFormula === undefined || nextFormula === formula) {
        return source
      }
      changed = true
      return `${open}${escapeXmlText(nextFormula)}${close}`
    },
  )
  return changed ? removeInvalidatedAuxiliaryChartSeriesRefs(rewrittenXml) : rewrittenXml
}

function renameChartFormulaXmlSheetReferences(xml: string, oldSheetName: string, newSheetName: string): string {
  return xml.replace(
    /(<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>)([\s\S]*?)(<\/(?:[A-Za-z_][\w.-]*:)?f>)/gu,
    (source: string, open: string, text: string, close: string) => {
      const formula = decodeXmlText(text)
      const nextFormula = renameChartFormulaSheetReferences(formula, oldSheetName, newSheetName)
      return nextFormula === undefined || nextFormula === formula ? source : `${open}${escapeXmlText(nextFormula)}${close}`
    },
  )
}

function chartFormulaTexts(xml: string): string[] {
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/gu)].map((match) =>
    decodeXmlText(match[1] ?? ''),
  )
}

function rewriteChartFormulaForStructuralTransform(
  formula: string,
  sheetName: string,
  transform: StructuralAxisTransform,
): string | undefined {
  try {
    return rewriteFormulaForStructuralTransform(formula, chartPackageFormulaOwnerSheetName, sheetName, transform)
  } catch {
    return undefined
  }
}

function renameChartFormulaSheetReferences(formula: string, oldSheetName: string, newSheetName: string): string | undefined {
  try {
    return renameFormulaSheetReferences(formula, oldSheetName, newSheetName)
  } catch {
    return undefined
  }
}

function rewriteChartFormulaForSheetDeletion(formula: string, deletedSheetName: string): string | undefined {
  try {
    return serializeFormula(rewriteFormulaNodeForSheetDeletion(parseFormula(formula), deletedSheetName))
  } catch {
    return undefined
  }
}

function rewriteFormulaNodeForSheetDeletion(node: FormulaNode, deletedSheetName: string): FormulaNode {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'StructuredRef':
      return node
    case 'NameRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
      return node.sheetName === deletedSheetName ? refErrorNode() : node
    case 'RangeRef':
      return node.sheetName === deletedSheetName || node.sheetEndName === deletedSheetName ? refErrorNode() : node
    case 'ArrayConstant':
      return { ...node, rows: node.rows.map((row) => row.map((entry) => rewriteFormulaNodeForSheetDeletion(entry, deletedSheetName))) }
    case 'UnaryExpr':
      return {
        ...node,
        argument: rewriteFormulaNodeForSheetDeletion(node.argument, deletedSheetName),
      }
    case 'BinaryExpr':
      return {
        ...node,
        left: rewriteFormulaNodeForSheetDeletion(node.left, deletedSheetName),
        right: rewriteFormulaNodeForSheetDeletion(node.right, deletedSheetName),
      }
    case 'CallExpr':
      return {
        ...node,
        args: node.args.map((arg) => rewriteFormulaNodeForSheetDeletion(arg, deletedSheetName)),
      }
    case 'InvokeExpr':
      return {
        ...node,
        callee: rewriteFormulaNodeForSheetDeletion(node.callee, deletedSheetName),
        args: node.args.map((arg) => rewriteFormulaNodeForSheetDeletion(arg, deletedSheetName)),
      }
  }
}

function refErrorNode(): FormulaNode {
  return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
}

function removeInvalidatedAuxiliaryChartSeriesRefs(xml: string): string {
  return xml.replace(/<((?:[A-Za-z_][\w.-]*:)?ser)\b[^>]*>[\s\S]*?<\/\1>/gu, (seriesXml: string) => {
    let nextSeriesXml = removeInvalidatedChartSeriesElement(seriesXml, 'tx')
    nextSeriesXml = removeInvalidatedChartSeriesElement(nextSeriesXml, 'cat')
    return nextSeriesXml
  })
}

function removeInvalidatedChartSeriesElement(seriesXml: string, localName: 'tx' | 'cat'): string {
  const elementPattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`, 'gu')
  return seriesXml.replace(elementPattern, (elementXml: string) => {
    const formulas = chartFormulaTexts(elementXml)
    return formulas.length > 0 && formulas.every((formula) => formula === '#REF!') ? '' : elementXml
  })
}

function chartFormulaWouldRewriteForDelete(
  formula: string,
  sheetName: string,
  axis: StructuralAxisKind,
  start: number,
  count: number,
): boolean {
  const nextFormula = rewriteChartFormulaForStructuralTransform(formula, sheetName, {
    kind: 'delete',
    axis,
    start,
    count,
  })
  return nextFormula !== undefined && nextFormula !== formula
}

function normalizeZipPath(path: string): string {
  return path.replace(/\\/gu, '/').replace(/^\/+/u, '')
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

function encodeBinaryString(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += binaryChunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + binaryChunkSize))
  }
  return binary
}

function decodeBinaryString(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

function encodeBase64(bytes: Uint8Array): string {
  return globalThis.btoa(encodeBinaryString(bytes))
}

function decodeBase64(dataBase64: string): Uint8Array {
  return decodeBinaryString(globalThis.atob(dataBase64))
}
