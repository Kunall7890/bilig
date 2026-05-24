import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  renameFormulaSheetReferences,
  rewriteAddressForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import type { WorkbookPreservedMetadataRecord, WorkbookPreservedSheetMetadataRecord } from '../../workbook-metadata-types.js'
import { hasPreservedSheetMetadata } from '../../workbook-preserved-metadata.js'
import type { WorkbookSheetDeletionMetadataContext, WorkbookSheetReorderMetadataContext } from '../../workbook-metadata-service-contract.js'

const cellReferencePattern = /^\$?([A-Z]+)\$?([1-9]\d*)$/iu
const pivotTableDefinitionLocationPattern = /<((?:[A-Za-z_][\w.-]*:)?location)\b([^>]*\bref=(["'])([^"']+)\3[^>]*)\/?>/gu
const pivotTablePartPathPattern = /^xl\/pivotTables\/pivotTable\d+\.xml$/u
const pivotCacheDefinitionPartPathPattern = /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/u
const worksheetSourcePattern = /<((?:[A-Za-z_][\w.-]*:)?worksheetSource)\b([^>]*\bref=(["'])([^"']+)\3[^>]*)\/?>/gu
const worksheetSourceElementPattern = /<((?:[A-Za-z_][\w.-]*:)?worksheetSource)\b[^>]*\/?>/gu
const workbookViewElementPattern = /<((?:[A-Za-z_][\w.-]*:)?workbookView)\b[^>]*\/?>/gu
const queryTableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable'
const slicerRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
type WorkbookViewIndexAttribute = 'activeTab' | 'firstSheet'
type WorkbookSlicerConnectionArtifactsRecord = NonNullable<WorkbookPreservedMetadataRecord['slicerConnectionArtifacts']>
type WorkbookSlicerConnectionSheetArtifactRecord = NonNullable<WorkbookSlicerConnectionArtifactsRecord['sheetArtifacts']>[number]

export function rewritePreservedSheetMetadataForStructuralTransform(
  metadata: WorkbookPreservedSheetMetadataRecord | undefined,
  transform: StructuralAxisTransform,
): WorkbookPreservedSheetMetadataRecord | undefined {
  if (!metadata) {
    return undefined
  }
  const next: WorkbookPreservedSheetMetadataRecord = { ...metadata }
  if (metadata.styleArtifacts) {
    const styleArtifacts = rewriteStyleArtifactsForStructuralTransform(metadata.styleArtifacts, transform)
    if (styleArtifacts) {
      next.styleArtifacts = styleArtifacts
    } else {
      delete next.styleArtifacts
    }
  }
  return hasPreservedSheetMetadata(next) ? next : undefined
}

export function renamePreservedWorkbookMetadataSheetReferences(
  metadata: WorkbookPreservedMetadataRecord,
  oldSheetName: string,
  newSheetName: string,
): WorkbookPreservedMetadataRecord | undefined {
  let changed = false
  const next: WorkbookPreservedMetadataRecord = { ...metadata }

  if (metadata.unsupportedFormulaDependencies) {
    next.unsupportedFormulaDependencies = metadata.unsupportedFormulaDependencies.map((entry) => {
      const sheetName = renameSheetName(entry.sheetName, oldSheetName, newSheetName)
      const formula = renameFormulaText(entry.formula, oldSheetName, newSheetName)
      const importedFormula = renameFormulaText(entry.importedFormula, oldSheetName, newSheetName)
      changed ||= sheetName !== entry.sheetName || formula !== entry.formula || importedFormula !== entry.importedFormula
      return { ...entry, sheetName, formula, importedFormula }
    })
  }

  if (metadata.unsupportedPivots) {
    next.unsupportedPivots = metadata.unsupportedPivots.map((entry) => {
      if (entry.sheetName === undefined) {
        return { ...entry }
      }
      const sheetName = renameSheetName(entry.sheetName, oldSheetName, newSheetName)
      changed ||= sheetName !== entry.sheetName
      return { ...entry, sheetName }
    })
  }

  if (metadata.formulaAudit) {
    const formulas = metadata.formulaAudit.formulas.map((entry) => {
      const formula = renameFormulaText(entry.formula, oldSheetName, newSheetName)
      if (entry.sheetName === undefined) {
        changed ||= formula !== entry.formula
        return { ...entry, formula }
      }
      const sheetName = renameSheetName(entry.sheetName, oldSheetName, newSheetName)
      changed ||= sheetName !== entry.sheetName || formula !== entry.formula
      return { ...entry, sheetName, formula }
    })
    const calcChain = metadata.formulaAudit.calcChain
      ? {
          ...metadata.formulaAudit.calcChain,
          cells: metadata.formulaAudit.calcChain.cells.map((entry) => {
            if (entry.sheetName === undefined) {
              return { ...entry }
            }
            const sheetName = renameSheetName(entry.sheetName, oldSheetName, newSheetName)
            changed ||= sheetName !== entry.sheetName
            return { ...entry, sheetName }
          }),
        }
      : undefined
    next.formulaAudit = {
      ...metadata.formulaAudit,
      formulas,
      ...(calcChain ? { calcChain } : {}),
    }
  }

  if (metadata.pivotArtifacts) {
    next.pivotArtifacts = {
      ...metadata.pivotArtifacts,
      parts: metadata.pivotArtifacts.parts.map((part) => {
        const normalizedPath = normalizePivotPackagePath(part.path)
        if (!pivotCacheDefinitionPartPathPattern.test(normalizedPath)) {
          return part
        }
        const xml = renamePivotCacheWorksheetSourceSheetReferences(part.xml, oldSheetName, newSheetName)
        changed ||= xml !== part.xml
        return xml === part.xml ? part : { ...part, xml }
      }),
    }
  }

  if (metadata.chartSheetArtifacts) {
    next.chartSheetArtifacts = metadata.chartSheetArtifacts.map((entry) => {
      const name = renameSheetName(entry.name, oldSheetName, newSheetName)
      changed ||= name !== entry.name
      return { ...entry, name }
    })
  }

  if (metadata.slicerConnectionArtifacts) {
    const sheetArtifacts = metadata.slicerConnectionArtifacts.sheetArtifacts?.map((entry) => {
      const sheetName = renameSheetName(entry.sheetName, oldSheetName, newSheetName)
      changed ||= sheetName !== entry.sheetName
      return { ...entry, sheetName }
    })
    next.slicerConnectionArtifacts = {
      ...metadata.slicerConnectionArtifacts,
      ...(sheetArtifacts ? { sheetArtifacts } : {}),
    }
  }

  return changed ? next : undefined
}

export function rewritePreservedWorkbookMetadataForSheetDeletion(
  metadata: WorkbookPreservedMetadataRecord,
  deletedSheetName: string,
  context: WorkbookSheetDeletionMetadataContext,
): WorkbookPreservedMetadataRecord | undefined {
  let changed = false
  const next: WorkbookPreservedMetadataRecord = { ...metadata }

  const bookViewsXml = metadata.viewState?.bookViewsXml
  if (bookViewsXml && context.sheetCountBeforeDelete > 1) {
    const rewrittenBookViewsXml = rewriteWorkbookViewIndexesForSheetDeletion(bookViewsXml, context)
    if (rewrittenBookViewsXml !== bookViewsXml) {
      changed = true
      next.viewState = {
        ...metadata.viewState,
        bookViewsXml: rewrittenBookViewsXml,
      }
    }
  }

  if (metadata.formulaAudit) {
    const formulaAudit = rewriteFormulaAuditForSheetDeletion(metadata.formulaAudit, deletedSheetName, context)
    changed ||= formulaAudit !== metadata.formulaAudit
    if (formulaAudit) {
      next.formulaAudit = formulaAudit
    } else {
      delete next.formulaAudit
    }
  }

  if (metadata.unsupportedFormulaDependencies) {
    const unsupportedFormulaDependencies = metadata.unsupportedFormulaDependencies.filter((entry) => entry.sheetName !== deletedSheetName)
    changed ||= unsupportedFormulaDependencies.length !== metadata.unsupportedFormulaDependencies.length
    if (unsupportedFormulaDependencies.length > 0) {
      next.unsupportedFormulaDependencies = unsupportedFormulaDependencies
    } else {
      delete next.unsupportedFormulaDependencies
    }
  }

  if (metadata.unsupportedPivots) {
    const unsupportedPivots = metadata.unsupportedPivots.filter((entry) => entry.sheetName !== deletedSheetName)
    changed ||= unsupportedPivots.length !== metadata.unsupportedPivots.length
    if (unsupportedPivots.length > 0) {
      next.unsupportedPivots = unsupportedPivots
    } else {
      delete next.unsupportedPivots
    }
  }

  if (metadata.slicerConnectionArtifacts) {
    const slicerConnectionArtifacts = rewriteSlicerConnectionArtifactsForSheetDeletion(metadata.slicerConnectionArtifacts, deletedSheetName)
    changed ||= slicerConnectionArtifacts !== metadata.slicerConnectionArtifacts
    if (slicerConnectionArtifacts) {
      next.slicerConnectionArtifacts = slicerConnectionArtifacts
    } else {
      delete next.slicerConnectionArtifacts
    }
  }

  return changed ? next : undefined
}

export function rewritePreservedPivotPackageArtifactsForSheetDeletion(
  workbookMetadata: WorkbookPreservedMetadataRecord | undefined,
  sheetMetadata: WorkbookPreservedSheetMetadataRecord | undefined,
  context: WorkbookSheetDeletionMetadataContext | undefined,
): WorkbookPreservedMetadataRecord | undefined {
  const pivotArtifacts = workbookMetadata?.pivotArtifacts
  const sheetPivotArtifacts = sheetMetadata?.pivotArtifacts
  if (!workbookMetadata || !pivotArtifacts || !sheetPivotArtifacts || !context) {
    return workbookMetadata
  }

  const deletedPivotPartPaths = pivotTablePartPathsForSheet(context.deletedSheetIndex, sheetPivotArtifacts)
  if (deletedPivotPartPaths.size === 0) {
    return workbookMetadata
  }

  const deletedCacheIds = new Set<string>()
  const remainingCacheIds = new Set<string>()
  for (const part of pivotArtifacts.parts) {
    const normalizedPath = normalizePivotPackagePath(part.path)
    if (!pivotTablePartPathPattern.test(normalizedPath)) {
      continue
    }
    const cacheId = readXmlAttribute(part.xml, 'cacheId')
    if (!cacheId) {
      continue
    }
    if (deletedPivotPartPaths.has(normalizedPath)) {
      deletedCacheIds.add(cacheId)
    } else {
      remainingCacheIds.add(cacheId)
    }
  }

  const deletedOnlyCacheIds = new Set([...deletedCacheIds].filter((cacheId) => !remainingCacheIds.has(cacheId)))
  const cacheDefinitionPathsToRemove = pivotCacheDefinitionPathsForCacheIds(pivotArtifacts, deletedOnlyCacheIds)
  const nextParts = pivotArtifacts.parts.filter((part) => {
    const normalizedPath = normalizePivotPackagePath(part.path)
    return !deletedPivotPartPaths.has(normalizedPath) && !cacheDefinitionPathsToRemove.has(normalizedPath)
  })
  const nextWorkbookPivotCachesXml = removeWorkbookPivotCacheEntries(pivotArtifacts.workbookPivotCachesXml, deletedOnlyCacheIds)
  const nextWorkbookRelationships = pivotArtifacts.workbookRelationships?.filter((relationship) => {
    const targetPath = normalizePivotPackagePath(resolvePivotRelationshipTarget('xl/workbook.xml', relationship.target))
    return !cacheDefinitionPathsToRemove.has(targetPath)
  })

  const nextPivotArtifacts = {
    parts: nextParts,
    ...(nextWorkbookPivotCachesXml ? { workbookPivotCachesXml: nextWorkbookPivotCachesXml } : {}),
    ...(nextWorkbookRelationships && nextWorkbookRelationships.length > 0 ? { workbookRelationships: nextWorkbookRelationships } : {}),
  }

  const next: WorkbookPreservedMetadataRecord = { ...workbookMetadata }
  if (hasPivotArtifacts(nextPivotArtifacts)) {
    next.pivotArtifacts = nextPivotArtifacts
  } else {
    delete next.pivotArtifacts
  }
  return next
}

export function rewritePreservedWorkbookMetadataForSheetReorder(
  metadata: WorkbookPreservedMetadataRecord,
  context: WorkbookSheetReorderMetadataContext,
): WorkbookPreservedMetadataRecord | undefined {
  const bookViewsXml = metadata.viewState?.bookViewsXml
  if (!bookViewsXml || context.oldSheetIndex === context.newSheetIndex || context.sheetCount <= 1) {
    return undefined
  }

  const rewrittenBookViewsXml = rewriteWorkbookViewIndexesForSheetReorder(bookViewsXml, context)
  if (rewrittenBookViewsXml === bookViewsXml) {
    return undefined
  }
  return {
    ...metadata,
    viewState: {
      ...metadata.viewState,
      bookViewsXml: rewrittenBookViewsXml,
    },
  }
}

export function rewritePreservedPivotPackageArtifactsForStructuralTransform(
  workbookMetadata: WorkbookPreservedMetadataRecord | undefined,
  sheetMetadata: WorkbookPreservedSheetMetadataRecord | undefined,
  sheetName: string,
  sheetIndex: number,
  transform: StructuralAxisTransform,
): WorkbookPreservedMetadataRecord | undefined {
  const pivotArtifacts = workbookMetadata?.pivotArtifacts
  if (!workbookMetadata || !pivotArtifacts) {
    return workbookMetadata
  }

  const pivotPartPaths = sheetMetadata?.pivotArtifacts ? pivotTablePartPathsForSheet(sheetIndex, sheetMetadata.pivotArtifacts) : new Set()

  return {
    ...workbookMetadata,
    pivotArtifacts: {
      ...pivotArtifacts,
      parts: pivotArtifacts.parts.flatMap((part) => {
        const normalizedPath = normalizePivotPackagePath(part.path)
        if (pivotPartPaths.has(normalizedPath) && pivotTablePartPathPattern.test(normalizedPath)) {
          const xml = rewritePivotTableDefinitionLocationRefsForStructuralTransform(part.xml, transform)
          return xml ? [{ ...part, xml }] : []
        }
        if (pivotCacheDefinitionPartPathPattern.test(normalizedPath)) {
          const xml = rewritePivotCacheWorksheetSourceRefsForStructuralTransform(part.xml, sheetName, transform)
          return xml ? [{ ...part, xml }] : []
        }
        if (!pivotPartPaths.has(normalizedPath)) {
          return [part]
        }
        return [part]
      }),
    },
  }
}

export function preservedSheetMetadataTouchesStructuralDelete(
  metadata: WorkbookPreservedSheetMetadataRecord | undefined,
  axis: 'row' | 'column',
  start: number,
): boolean {
  if (!metadata) {
    return false
  }
  const styleArtifacts = metadata.styleArtifacts
  if (
    styleArtifacts?.cellStyleIndexes.some((entry) => cellReferenceTouchesAxisDelete(entry.address, axis, start)) ||
    styleArtifacts?.blankCellAddresses?.some((address) => cellReferenceTouchesAxisDelete(address, axis, start))
  ) {
    return true
  }
  return metadata.pivotArtifacts !== undefined
}

function rewriteWorkbookViewIndexesForSheetDeletion(xml: string, context: WorkbookSheetDeletionMetadataContext): string {
  return xml.replace(workbookViewElementPattern, (element) => {
    let nextElement = rewriteWorkbookViewIndexAttributeForSheetDeletion(element, 'activeTab', context, 'active')
    nextElement = rewriteWorkbookViewIndexAttributeForSheetDeletion(nextElement, 'firstSheet', context, 'first')
    return nextElement
  })
}

function rewriteWorkbookViewIndexesForSheetReorder(xml: string, context: WorkbookSheetReorderMetadataContext): string {
  return xml.replace(workbookViewElementPattern, (element) => {
    return rewriteWorkbookViewIndexAttributeForSheetReorder(element, 'activeTab', context)
  })
}

function rewriteWorkbookViewIndexAttributeForSheetReorder(
  element: string,
  attributeName: WorkbookViewIndexAttribute,
  context: WorkbookSheetReorderMetadataContext,
): string {
  const attribute = readNonNegativeIntegerXmlAttribute(element, attributeName)
  if (!attribute) {
    return element
  }
  const nextIndex = sheetIndexAfterSheetReorder(attribute.value, context)
  if (nextIndex === attribute.value) {
    return element
  }
  return replaceXmlAttributeValue(element, attributeName, attribute.quote, String(nextIndex))
}

function rewriteWorkbookViewIndexAttributeForSheetDeletion(
  element: string,
  attributeName: 'activeTab' | 'firstSheet',
  context: WorkbookSheetDeletionMetadataContext,
  mode: 'active' | 'first',
): string {
  const attribute = readNonNegativeIntegerXmlAttribute(element, attributeName)
  if (!attribute) {
    return element
  }
  const nextIndex =
    mode === 'active'
      ? sheetIndexAfterActiveSheetDeletion(attribute.value, context)
      : sheetIndexAfterFirstVisibleSheetDeletion(attribute.value, context)
  if (attributeName === 'firstSheet' && nextIndex === 0) {
    return removeXmlAttribute(element, attributeName)
  }
  if (nextIndex === attribute.value) {
    return element
  }
  return replaceXmlAttributeValue(element, attributeName, attribute.quote, String(nextIndex))
}

function sheetIndexAfterActiveSheetDeletion(index: number, context: WorkbookSheetDeletionMetadataContext): number {
  const sheetCountAfterDelete = Math.max(0, context.sheetCountBeforeDelete - 1)
  if (sheetCountAfterDelete === 0) {
    return 0
  }
  if (index > context.deletedSheetIndex) {
    return index - 1
  }
  if (index === context.deletedSheetIndex) {
    return Math.min(index, sheetCountAfterDelete - 1)
  }
  return Math.min(index, sheetCountAfterDelete - 1)
}

function sheetIndexAfterFirstVisibleSheetDeletion(index: number, context: WorkbookSheetDeletionMetadataContext): number {
  const sheetCountAfterDelete = Math.max(0, context.sheetCountBeforeDelete - 1)
  if (sheetCountAfterDelete === 0) {
    return 0
  }
  const shiftedIndex = index > context.deletedSheetIndex ? index - 1 : index
  return Math.min(shiftedIndex, sheetCountAfterDelete - 1)
}

function sheetIndexAfterSheetReorder(index: number, context: WorkbookSheetReorderMetadataContext): number {
  if (index < 0 || index >= context.sheetCount) {
    return index
  }
  if (index === context.oldSheetIndex) {
    return context.newSheetIndex
  }
  if (context.oldSheetIndex < context.newSheetIndex && index > context.oldSheetIndex && index <= context.newSheetIndex) {
    return index - 1
  }
  if (context.oldSheetIndex > context.newSheetIndex && index >= context.newSheetIndex && index < context.oldSheetIndex) {
    return index + 1
  }
  return index
}

function readNonNegativeIntegerXmlAttribute(
  element: string,
  attributeName: WorkbookViewIndexAttribute,
): { readonly value: number; readonly quote: string } | undefined {
  const match = new RegExp(`\\b${attributeName}=(["'])(\\d+)\\1`, 'u').exec(element)
  if (!match) {
    return undefined
  }
  const value = Number(match[2])
  if (!Number.isSafeInteger(value)) {
    return undefined
  }
  const quote = match[1]
  if (quote !== '"' && quote !== "'") {
    return undefined
  }
  return { value, quote }
}

function replaceXmlAttributeValue(element: string, attributeName: WorkbookViewIndexAttribute, quote: string, value: string): string {
  return element.replace(new RegExp(`\\b${attributeName}=(["'])\\d+\\1`, 'u'), `${attributeName}=${quote}${value}${quote}`)
}

function removeXmlAttribute(element: string, attributeName: 'firstSheet'): string {
  return element.replace(new RegExp(`\\s+${attributeName}=(["'])\\d+\\1`, 'u'), '')
}

function rewriteFormulaAuditForSheetDeletion(
  formulaAudit: NonNullable<WorkbookPreservedMetadataRecord['formulaAudit']>,
  deletedSheetName: string,
  context: WorkbookSheetDeletionMetadataContext,
): NonNullable<WorkbookPreservedMetadataRecord['formulaAudit']> | undefined {
  const formulas = formulaAudit.formulas.filter((entry) => entry.sheetName !== deletedSheetName)
  const diagnostics = formulaAudit.diagnostics?.filter((entry) => entry.sheetName !== deletedSheetName)
  const calcChainCells = formulaAudit.calcChain?.cells.filter(
    (entry) => entry.sheetName !== deletedSheetName && entry.sheetIndex !== context.deletedSheetId,
  )
  const formulasChanged = formulas.length !== formulaAudit.formulas.length
  const diagnosticsChanged = diagnostics !== undefined && diagnostics.length !== formulaAudit.diagnostics?.length
  const calcChainChanged = calcChainCells !== undefined && calcChainCells.length !== formulaAudit.calcChain?.cells.length

  if (!formulasChanged && !diagnosticsChanged && !calcChainChanged) {
    return formulaAudit
  }
  const calcChain =
    formulaAudit.calcChain && calcChainCells && calcChainCells.length > 0 ? { ...formulaAudit.calcChain, cells: calcChainCells } : undefined
  if (formulas.length === 0 && (diagnostics?.length ?? 0) === 0 && !calcChain) {
    return undefined
  }
  return {
    formulas,
    ...(diagnostics && diagnostics.length > 0 ? { diagnostics } : {}),
    ...(calcChain ? { calcChain } : {}),
  }
}

function rewriteSlicerConnectionArtifactsForSheetDeletion(
  artifacts: WorkbookSlicerConnectionArtifactsRecord,
  deletedSheetName: string,
): WorkbookSlicerConnectionArtifactsRecord | undefined {
  const sourceSheetArtifacts = artifacts.sheetArtifacts ?? []
  const remainingSheetArtifacts = sourceSheetArtifacts.filter((entry) => entry.sheetName !== deletedSheetName)
  if (remainingSheetArtifacts.length === sourceSheetArtifacts.length) {
    return artifacts
  }

  const deletedSlicerPartPaths = slicerPartPathsReferencedBySheetArtifacts(
    sourceSheetArtifacts.filter((entry) => entry.sheetName === deletedSheetName),
  )
  const remainingSlicerPartPaths = slicerPartPathsReferencedBySheetArtifacts(remainingSheetArtifacts)
  const removedPartPaths = new Set<string>()
  const parts = artifacts.parts.filter((part) => {
    const path = normalizePackagePath(part.path)
    const slicerPartPath = packagePartPathFromRelationshipPartPath(path)
    const shouldRemove =
      (deletedSlicerPartPaths.has(path) || deletedSlicerPartPaths.has(slicerPartPath)) && !remainingSlicerPartPaths.has(slicerPartPath)
    if (shouldRemove) {
      removedPartPaths.add(path)
      return false
    }
    return true
  })

  const contentTypeOverrides = (artifacts.contentTypeOverrides ?? []).filter(
    (entry) => !removedPartPaths.has(normalizePackagePath(entry.partName)),
  )
  const next: WorkbookSlicerConnectionArtifactsRecord = {
    parts,
    ...(artifacts.workbookSlicerCachesExtXml ? { workbookSlicerCachesExtXml: artifacts.workbookSlicerCachesExtXml } : {}),
    ...(artifacts.workbookRelationships ? { workbookRelationships: artifacts.workbookRelationships } : {}),
    ...(remainingSheetArtifacts.length > 0 ? { sheetArtifacts: remainingSheetArtifacts } : {}),
    ...(artifacts.contentTypeDefaults ? { contentTypeDefaults: artifacts.contentTypeDefaults } : {}),
    ...(contentTypeOverrides.length > 0 ? { contentTypeOverrides } : {}),
  }

  return hasSlicerConnectionArtifacts(next) ? next : undefined
}

function slicerPartPathsReferencedBySheetArtifacts(sheetArtifacts: readonly WorkbookSlicerConnectionSheetArtifactRecord[]): Set<string> {
  return new Set(
    sheetArtifacts
      .flatMap((entry) => entry.relationships ?? [])
      .filter((relationship) => relationship.type === slicerRelationshipType || relationship.type === queryTableRelationshipType)
      .map((relationship) => normalizePackagePath(resolvePackageRelationshipTarget('xl/worksheets/sheet1.xml', relationship.target))),
  )
}

function hasSlicerConnectionArtifacts(artifacts: WorkbookSlicerConnectionArtifactsRecord): boolean {
  return (
    artifacts.parts.length > 0 ||
    artifacts.workbookSlicerCachesExtXml !== undefined ||
    (artifacts.workbookRelationships?.length ?? 0) > 0 ||
    (artifacts.sheetArtifacts?.length ?? 0) > 0 ||
    (artifacts.contentTypeOverrides?.length ?? 0) > 0
  )
}

function packagePartPathFromRelationshipPartPath(path: string): string {
  return path.replace(/\/_rels\/([^/]+)\.rels$/u, '/$1')
}

function resolvePackageRelationshipTarget(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1)
  }
  const segments = basePartPath.split('/')
  segments.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      segments.pop()
    } else if (segment !== '.' && segment.length > 0) {
      segments.push(segment)
    }
  }
  return segments.join('/')
}

function rewriteStyleArtifactsForStructuralTransform(
  styleArtifacts: NonNullable<WorkbookPreservedSheetMetadataRecord['styleArtifacts']>,
  transform: StructuralAxisTransform,
): WorkbookPreservedSheetMetadataRecord['styleArtifacts'] | undefined {
  const cellStyleIndexes = styleArtifacts.cellStyleIndexes.flatMap((entry) => {
    const address = rewriteCellReferenceForStructuralTransform(entry.address, transform)
    return address ? [{ ...entry, address }] : []
  })
  const blankCellAddresses = (styleArtifacts.blankCellAddresses ?? []).flatMap((address) => {
    const nextAddress = rewriteCellReferenceForStructuralTransform(address, transform)
    return nextAddress ? [nextAddress] : []
  })
  if (cellStyleIndexes.length === 0 && blankCellAddresses.length === 0) {
    return undefined
  }
  return {
    cellStyleIndexes,
    ...(blankCellAddresses.length > 0 ? { blankCellAddresses } : {}),
  }
}

function pivotTablePartPathsForSheet(
  sheetIndex: number,
  pivotArtifacts: NonNullable<WorkbookPreservedSheetMetadataRecord['pivotArtifacts']>,
): Set<string> {
  const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
  return new Set(
    pivotArtifacts.relationships
      .filter((relationship) => relationship.type.endsWith('/pivotTable'))
      .map((relationship) => normalizePivotPackagePath(resolvePivotRelationshipTarget(sheetPath, relationship.target))),
  )
}

function pivotCacheDefinitionPathsForCacheIds(
  pivotArtifacts: NonNullable<WorkbookPreservedMetadataRecord['pivotArtifacts']>,
  cacheIds: ReadonlySet<string>,
): Set<string> {
  if (cacheIds.size === 0) {
    return new Set()
  }
  const cacheRelationshipTargetsById = new Map(
    (pivotArtifacts.workbookRelationships ?? [])
      .filter((relationship) => relationship.type.endsWith('/pivotCacheDefinition'))
      .map((relationship) => [
        relationship.id,
        normalizePivotPackagePath(resolvePivotRelationshipTarget('xl/workbook.xml', relationship.target)),
      ]),
  )
  const output = new Set<string>()
  for (const match of (pivotArtifacts.workbookPivotCachesXml ?? '').matchAll(/<((?:[A-Za-z_][\w.-]*:)?pivotCache)\b([^>]*?)\/?>/gu)) {
    const attributes = match[2] ?? ''
    const cacheId = readXmlAttribute(attributes, 'cacheId')
    const relationshipId = readXmlAttribute(attributes, 'r:id') ?? readXmlAttribute(attributes, 'id')
    const target = relationshipId ? cacheRelationshipTargetsById.get(relationshipId) : undefined
    if (cacheId && target && cacheIds.has(cacheId)) {
      output.add(target)
    }
  }
  return output
}

function removeWorkbookPivotCacheEntries(xml: string | undefined, cacheIds: ReadonlySet<string>): string | undefined {
  if (!xml || cacheIds.size === 0) {
    return xml
  }
  const nextXml = xml.replace(/<((?:[A-Za-z_][\w.-]*:)?pivotCache)\b([^>]*?)\/?>/gu, (source: string, _tag: string, attributes: string) => {
    const cacheId = readXmlAttribute(attributes, 'cacheId')
    return cacheId && cacheIds.has(cacheId) ? '' : source
  })
  if (nextXml === xml) {
    return xml
  }
  return /<((?:[A-Za-z_][\w.-]*:)?pivotCache)\b/u.test(nextXml) ? nextXml : undefined
}

function hasPivotArtifacts(artifacts: NonNullable<WorkbookPreservedMetadataRecord['pivotArtifacts']>): boolean {
  return artifacts.parts.length > 0 || Boolean(artifacts.workbookPivotCachesXml) || (artifacts.workbookRelationships?.length ?? 0) > 0
}

function resolvePivotRelationshipTarget(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1)
  }
  const segments = basePartPath.split('/')
  segments.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      segments.pop()
    } else if (segment !== '.' && segment.length > 0) {
      segments.push(segment)
    }
  }
  return segments.join('/')
}

function normalizePivotPackagePath(path: string): string {
  return path.replace(/^\/+/u, '').replace(/\\/gu, '/')
}

function normalizePackagePath(path: string): string {
  return normalizePivotPackagePath(path)
}

function rewritePivotTableDefinitionLocationRefsForStructuralTransform(
  xml: string,
  transform: StructuralAxisTransform,
): string | undefined {
  let failed = false
  const nextXml = xml.replace(
    pivotTableDefinitionLocationPattern,
    (source: string, _tagName: string, _attributes: string, quote: string, ref: string) => {
      const rewrittenRef = rewritePivotLocationRefForStructuralTransform(ref, transform)
      if (!rewrittenRef) {
        failed = true
        return source
      }
      return source.replace(`ref=${quote}${ref}${quote}`, `ref=${quote}${rewrittenRef}${quote}`)
    },
  )
  return failed ? undefined : nextXml
}

function rewritePivotCacheWorksheetSourceRefsForStructuralTransform(
  xml: string,
  sheetName: string,
  transform: StructuralAxisTransform,
): string | undefined {
  let failed = false
  const nextXml = xml.replace(
    worksheetSourcePattern,
    (source: string, _tagName: string, attributes: string, quote: string, ref: string) => {
      const sourceSheetName = readXmlAttribute(attributes, 'sheet')
      if (sourceSheetName === undefined || unescapeXmlAttribute(sourceSheetName) !== sheetName) {
        return source
      }
      const rewrittenRef = rewritePivotLocationRefForStructuralTransform(ref, transform)
      if (!rewrittenRef) {
        failed = true
        return source
      }
      return source.replace(`ref=${quote}${ref}${quote}`, `ref=${quote}${rewrittenRef}${quote}`)
    },
  )
  return failed ? undefined : nextXml
}

function renamePivotCacheWorksheetSourceSheetReferences(xml: string, oldSheetName: string, newSheetName: string): string {
  return xml.replace(worksheetSourceElementPattern, (source: string) => {
    const attribute = readXmlAttributeMatch(source, 'sheet')
    if (!attribute || unescapeXmlAttribute(attribute.value) !== oldSheetName) {
      return source
    }
    return source.replace(attribute.raw, `sheet=${attribute.quote}${escapeXmlAttribute(newSheetName)}${attribute.quote}`)
  })
}

function rewritePivotLocationRefForStructuralTransform(ref: string, transform: StructuralAxisTransform): string | undefined {
  const [startAddress, endAddress = startAddress] = ref.split(':')
  if (!startAddress || !endAddress || !isCellReference(startAddress) || !isCellReference(endAddress)) {
    return ref
  }
  const rewritten = rewriteRangeForStructuralTransform(startAddress, endAddress, transform)
  if (!rewritten) {
    return undefined
  }
  const start = clipAddressToSheetGrid(rewritten.startAddress)
  const end = clipAddressToSheetGrid(rewritten.endAddress)
  if (!start || !end) {
    return undefined
  }
  return start === end ? start : `${start}:${end}`
}

function readXmlAttribute(attributes: string, name: string): string | undefined {
  return new RegExp(`\\b${name}=(["'])([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2]
}

function readXmlAttributeMatch(
  attributes: string,
  name: string,
): { readonly raw: string; readonly quote: string; readonly value: string } | undefined {
  const match = new RegExp(`\\b${name}=(["'])([\\s\\S]*?)\\1`, 'u').exec(attributes)
  if (!match || !match[0] || !match[1]) {
    return undefined
  }
  return { raw: match[0], quote: match[1], value: match[2] ?? '' }
}

function unescapeXmlAttribute(value: string): string {
  return value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&')
}

function escapeXmlAttribute(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll("'", '&apos;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function renameSheetName(value: string, oldSheetName: string, newSheetName: string): string {
  return value === oldSheetName ? newSheetName : value
}

function renameFormulaText(formula: string, oldSheetName: string, newSheetName: string): string {
  try {
    return renameFormulaSheetReferences(formula, oldSheetName, newSheetName)
  } catch {
    return formula
  }
}

function rewriteCellReferenceForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  if (!isCellReference(address)) {
    return address
  }
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  return rewritten ? clipAddressToSheetGrid(rewritten) : undefined
}

function cellReferenceTouchesAxisDelete(address: string, axis: 'row' | 'column', start: number): boolean {
  if (!isCellReference(address)) {
    return false
  }
  const parsed = parseCellAddress(address)
  if (!parsed) {
    return false
  }
  return (axis === 'row' ? parsed[0] : parsed[1]) >= start
}

function isCellReference(address: string): boolean {
  return cellReferencePattern.test(address)
}

function clipAddressToSheetGrid(address: string): string | undefined {
  const parsed = parseCellAddress(address)
  if (!parsed || parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function parseCellAddress(address: string): [number, number] | undefined {
  const match = cellReferencePattern.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
