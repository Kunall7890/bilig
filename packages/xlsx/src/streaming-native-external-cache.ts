import { ErrorCode, formatErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

import { decodeCellAddress, encodeCellAddress } from './address.js'
import type {
  XlsxExternalWorkbookHydrationDiagnostics,
  XlsxExternalWorkbookHydrationMatchKind,
  XlsxExternalWorkbookHydrationReferenceDiagnostic,
  XlsxExternalWorkbookInput,
} from './external-workbook-types.js'
import type { XlsxSourceTextPatch } from './source-preserving-literal-patches.js'
import type { NativeFormulaCell, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'
import { escapeXmlAttribute, escapeXmlText, readXmlAttribute } from './xml.js'
import { workbookSheetPathEntriesForSource } from './workbook-sheet-paths.js'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries } from './zip-reader.js'

interface ExternalReferenceScanTarget {
  readonly bookIndex: number
  readonly sheetName: string
  readonly cellsByRow: Map<number, Set<number>>
}

interface ExternalLinkReference {
  readonly bookIndex: number
  readonly path: string
  readonly target?: string
  readonly workbookName?: string
  readonly sheetNames: readonly string[]
}

interface ResolvedExternalWorkbookInput {
  readonly input?: XlsxExternalWorkbookInput
  readonly status: 'matched' | 'skipped-no-match' | 'skipped-ambiguous-match'
  readonly candidateCount: number
  readonly referenceCandidateCount?: number
  readonly matchKind?: XlsxExternalWorkbookHydrationMatchKind
}

export interface StreamingNativeExternalCachedRows {
  readonly rowsByAlias: ReadonlyMap<string, ReadonlyMap<number, Map<number, PendingCellValue>>>
  readonly textPatches: readonly XlsxSourceTextPatch[]
  readonly warnings: readonly string[]
  readonly diagnostics?: XlsxExternalWorkbookHydrationDiagnostics
}

const externalReferencePattern =
  /'\[([1-9][0-9]*)\]((?:[^']|'')+)'!((?:\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(?:\$?[A-Za-z]{1,3}\$?[0-9]+))?)|\[([1-9][0-9]*)\]([A-Za-z_][A-Za-z0-9_.]*)!((?:\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(?:\$?[A-Za-z]{1,3}\$?[0-9]+))?)/gu
const relationshipPattern = /<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu
const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const externalWorkbookCompanionNoMatchWarning =
  'Some supplied external workbook companions could not be matched; existing external-link cache values were preserved.'
const externalWorkbookCompanionAmbiguousMatchWarning =
  'Some supplied external workbook companions matched ambiguously; existing external-link cache values were preserved.'
const emptyCellValue: CellValue = Object.freeze({ tag: ValueTag.Empty })

export function normalizeExternalWorkbookReferences(formula: string): string {
  externalReferencePattern.lastIndex = 0
  return formula.replace(
    externalReferencePattern,
    (
      _match,
      quotedBookIndex: string | undefined,
      quotedSheetName: string | undefined,
      quotedRef: string | undefined,
      unquotedBookIndex: string | undefined,
      unquotedSheetName: string | undefined,
      unquotedRef: string | undefined,
    ) => {
      const bookIndex = Number(quotedBookIndex ?? unquotedBookIndex)
      const sheetName = quotedSheetName ? quotedSheetName.replaceAll("''", "'") : (unquotedSheetName ?? '')
      const ref = quotedRef ?? unquotedRef ?? ''
      return `${quoteFormulaSheetName(externalReferenceAlias(bookIndex, sheetName))}!${ref}`
    },
  )
}

export function readStreamingNativeExternalCachedRowsByAlias(
  zip: XlsxZipEntries,
  sheetScans: ReadonlyMap<string, SheetScanState>,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
  externalWorkbooks: readonly XlsxExternalWorkbookInput[] | undefined = undefined,
): StreamingNativeExternalCachedRows {
  const targetsByAlias = collectExternalReferenceScanTargets(sheetScans, resolveFormulaSource)
  if (targetsByAlias.size === 0) {
    return { rowsByAlias: new Map(), textPatches: [], warnings: [] }
  }
  const externalLinkReferencesByBookIndex = readExternalLinkReferencesByBookIndex(zip)
  const targetsByBookIndex = new Map<number, ExternalReferenceScanTarget[]>()
  for (const target of targetsByAlias.values()) {
    const targets = targetsByBookIndex.get(target.bookIndex) ?? []
    targets.push(target)
    targetsByBookIndex.set(target.bookIndex, targets)
  }
  const output = new Map<string, ReadonlyMap<number, Map<number, PendingCellValue>>>()
  for (const [bookIndex, targets] of targetsByBookIndex.entries()) {
    const reference = externalLinkReferencesByBookIndex.get(bookIndex)
    if (!reference) {
      continue
    }
    const externalLinkXml = getZipText(zip, reference.path)
    if (!externalLinkXml) {
      continue
    }
    for (const [alias, rows] of parseExternalLinkCachedRows(bookIndex, externalLinkXml, targets).entries()) {
      output.set(alias, rows)
    }
  }
  if (externalWorkbooks && externalWorkbooks.length > 0) {
    const companionRows = readCompanionExternalCachedRowsByAlias(externalWorkbooks, externalLinkReferencesByBookIndex, targetsByBookIndex)
    for (const [alias, rows] of companionRows.rowsByAlias.entries()) {
      output.set(alias, rows)
    }
    return {
      rowsByAlias: output,
      textPatches: companionRows.textPatches,
      warnings: companionRows.warnings,
      ...(companionRows.diagnostics === undefined ? {} : { diagnostics: companionRows.diagnostics }),
    }
  }
  return { rowsByAlias: output, textPatches: [], warnings: [] }
}

export function isStreamingNativeExternalReferenceAlias(sheetName: string): boolean {
  return sheetName.startsWith('__bilig_external_')
}

function collectExternalReferenceScanTargets(
  sheetScans: ReadonlyMap<string, SheetScanState>,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
): ReadonlyMap<string, ExternalReferenceScanTarget> {
  const targets = new Map<string, ExternalReferenceScanTarget>()
  for (const scan of sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      let formula
      try {
        formula = resolveFormulaSource(scan, cell)
      } catch {
        continue
      }
      for (const reference of extractExternalFormulaReferences(formula)) {
        const alias = externalReferenceAlias(reference.bookIndex, reference.sheetName)
        const target = targets.get(alias) ?? {
          bookIndex: reference.bookIndex,
          sheetName: reference.sheetName,
          cellsByRow: new Map<number, Set<number>>(),
        }
        for (let row = reference.startRow; row <= reference.endRow; row += 1) {
          const cols = target.cellsByRow.get(row) ?? new Set<number>()
          for (let col = reference.startCol; col <= reference.endCol; col += 1) {
            cols.add(col)
          }
          target.cellsByRow.set(row, cols)
        }
        targets.set(alias, target)
      }
    }
  }
  return targets
}

function extractExternalFormulaReferences(formula: string): readonly {
  readonly bookIndex: number
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}[] {
  externalReferencePattern.lastIndex = 0
  return [...formula.matchAll(externalReferencePattern)].flatMap((match) => {
    const bookIndex = Number(match[1] ?? match[4])
    const sheetName = match[2] ? match[2].replaceAll("''", "'") : (match[5] ?? '')
    const ref = match[3] ?? match[6] ?? ''
    if (!Number.isSafeInteger(bookIndex) || bookIndex < 1 || sheetName.length === 0 || ref.length === 0) {
      return []
    }
    const [startRef, endRef] = ref.split(':')
    if (!startRef) {
      return []
    }
    let start
    let end
    try {
      start = decodeCellAddress(startRef.replaceAll('$', ''))
      end = decodeCellAddress((endRef ?? startRef).replaceAll('$', ''))
    } catch {
      return []
    }
    return [
      {
        bookIndex,
        sheetName,
        startRow: Math.min(start.r, end.r),
        endRow: Math.max(start.r, end.r),
        startCol: Math.min(start.c, end.c),
        endCol: Math.max(start.c, end.c),
      },
    ]
  })
}

function readExternalLinkReferencesByBookIndex(zip: XlsxZipEntries): ReadonlyMap<number, ExternalLinkReference> {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  const workbookRelationshipsXml = getZipText(zip, 'xl/_rels/workbook.xml.rels')
  if (!workbookXml || !workbookRelationshipsXml) {
    return new Map()
  }
  const externalLinkTargetsByRelationshipId = new Map<string, string>()
  for (const match of workbookRelationshipsXml.matchAll(relationshipPattern)) {
    const tag = match[0]
    const id = readXmlAttribute(tag, 'Id')
    const target = readXmlAttribute(tag, 'Target')
    const type = readXmlAttribute(tag, 'Type')
    if (id && target && type && (type === externalLinkRelationshipType || type.endsWith('/externalLink'))) {
      externalLinkTargetsByRelationshipId.set(id, resolveTargetPath('xl/workbook.xml', target))
    }
  }
  const references = new Map<number, ExternalLinkReference>()
  let bookIndex = 1
  for (const match of workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?externalReference\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)) {
    const relationshipId = readXmlAttribute(match[0], 'r:id') ?? readXmlAttribute(match[0], 'id')
    const path = relationshipId ? externalLinkTargetsByRelationshipId.get(relationshipId) : undefined
    if (path) {
      const externalLinkXml = getZipText(zip, path)
      const externalTarget = readExternalLinkWorkbookTarget(zip, path)
      const workbookName = workbookIdentityFromName(externalTarget) ?? undefined
      references.set(bookIndex, {
        bookIndex,
        path,
        ...(externalTarget === undefined ? {} : { target: externalTarget }),
        ...(workbookName === undefined ? {} : { workbookName }),
        sheetNames: externalLinkXml ? readExternalLinkSheetNames(externalLinkXml) : [],
      })
    }
    bookIndex += 1
  }
  return references
}

function parseExternalLinkCachedRows(
  bookIndex: number,
  externalLinkXml: string,
  targets: readonly ExternalReferenceScanTarget[],
): ReadonlyMap<string, ReadonlyMap<number, Map<number, PendingCellValue>>> {
  const sheetNames = [...externalLinkXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheetName\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap(
    (match) => {
      const name = readXmlAttribute(match[0], 'val')
      return name ? [name] : []
    },
  )
  const targetsBySheetName = new Map(targets.map((target) => [target.sheetName, target]))
  const output = new Map<string, Map<number, Map<number, PendingCellValue>>>()
  for (const match of externalLinkXml.matchAll(
    /<((?:[A-Za-z_][\w.-]*:)?sheetData)\b((?:[^/>"']|"[^"]*"|'[^']*')*)(?:\/>|>([\s\S]*?)<\/\1>)/gu,
  )) {
    const sheetIdText = readXmlAttribute(match[2] ?? '', 'sheetId')
    const sheetId = sheetIdText === null ? Number.NaN : Number(sheetIdText)
    const sheetName = Number.isSafeInteger(sheetId) && sheetId >= 0 ? sheetNames[sheetId] : undefined
    const target = sheetName ? targetsBySheetName.get(sheetName) : undefined
    if (!target || sheetName === undefined) {
      continue
    }
    const rows = readExternalCachedRows(match[3] ?? '', target.cellsByRow, false)
    if (rows.size > 0) {
      output.set(externalReferenceAlias(bookIndex, sheetName), rows)
    }
  }
  return output
}

function readExternalLinkSheetNames(externalLinkXml: string): readonly string[] {
  return [...externalLinkXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheetName\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const name = readXmlAttribute(match[0], 'val')
    return name ? [name] : []
  })
}

function externalLinkRelationshipPartPath(partPath: string): string {
  const normalized = normalizeZipPath(partPath)
  const fileName = normalized.slice(normalized.lastIndexOf('/') + 1)
  return `xl/externalLinks/_rels/${fileName}.rels`
}

function readExternalLinkWorkbookTarget(zip: XlsxZipEntries, externalLinkPath: string): string | undefined {
  const relationshipsXml = getZipText(zip, externalLinkRelationshipPartPath(externalLinkPath))
  if (!relationshipsXml) {
    return undefined
  }
  for (const match of relationshipsXml.matchAll(relationshipPattern)) {
    const tag = match[0]
    const target = readXmlAttribute(tag, 'Target')
    const type = readXmlAttribute(tag, 'Type')
    const targetMode = readXmlAttribute(tag, 'TargetMode')
    if (target && (targetMode === 'External' || type?.endsWith('/externalLinkPath') === true)) {
      return target
    }
  }
  return undefined
}

function readCompanionExternalCachedRowsByAlias(
  externalWorkbooks: readonly XlsxExternalWorkbookInput[],
  referencesByBookIndex: ReadonlyMap<number, ExternalLinkReference>,
  targetsByBookIndex: ReadonlyMap<number, readonly ExternalReferenceScanTarget[]>,
): StreamingNativeExternalCachedRows {
  const output = new Map<string, ReadonlyMap<number, Map<number, PendingCellValue>>>()
  const textPatches: XlsxSourceTextPatch[] = []
  const referenceDiagnostics: XlsxExternalWorkbookHydrationReferenceDiagnostic[] = []
  const refreshedBookIndices: number[] = []
  let refreshedSheetCount = 0
  let refreshedCellCount = 0
  let skippedNoMatchCount = 0
  let skippedAmbiguousMatchCount = 0
  let skippedEmptyRefreshCount = 0
  for (const [bookIndex, targets] of targetsByBookIndex.entries()) {
    const reference = referencesByBookIndex.get(bookIndex)
    if (!reference) {
      continue
    }
    const resolved = resolveExternalWorkbookInput(externalWorkbooks, referencesByBookIndex, reference)
    if (!resolved.input) {
      const skippedStatus = resolved.status === 'skipped-no-match' ? 'skipped-no-match' : 'skipped-ambiguous-match'
      if (skippedStatus === 'skipped-no-match') {
        skippedNoMatchCount += 1
      } else {
        skippedAmbiguousMatchCount += 1
      }
      referenceDiagnostics.push({
        bookIndex,
        ...(reference.workbookName === undefined ? {} : { workbookName: reference.workbookName }),
        ...(reference.target === undefined ? {} : { target: reference.target }),
        status: skippedStatus,
        candidateCount: resolved.candidateCount,
        ...(resolved.referenceCandidateCount === undefined ? {} : { referenceCandidateCount: resolved.referenceCandidateCount }),
        ...(resolved.matchKind === undefined ? {} : { matchKind: resolved.matchKind }),
      })
      continue
    }
    const companionRows = readCompanionWorkbookCachedRowsByAlias(resolved.input, bookIndex, targets)
    let referenceSheetCount = 0
    let referenceCellCount = 0
    for (const [alias, rows] of companionRows.rowsByAlias.entries()) {
      output.set(alias, rows)
      referenceSheetCount += 1
      referenceCellCount += countExternalCachedCells(rows)
    }
    if (referenceSheetCount === 0 || referenceCellCount === 0) {
      skippedEmptyRefreshCount += 1
      referenceDiagnostics.push({
        bookIndex,
        ...(reference.workbookName === undefined ? {} : { workbookName: reference.workbookName }),
        ...(reference.target === undefined ? {} : { target: reference.target }),
        status: 'skipped-empty-refresh',
        candidateCount: resolved.candidateCount,
        ...(resolved.referenceCandidateCount === undefined ? {} : { referenceCandidateCount: resolved.referenceCandidateCount }),
        ...(resolved.matchKind === undefined ? {} : { matchKind: resolved.matchKind }),
        ...(resolved.input.fileName === undefined ? {} : { matchedFileName: resolved.input.fileName }),
        ...(resolved.input.workbookName === undefined ? {} : { matchedWorkbookName: resolved.input.workbookName }),
        ...(resolved.input.target === undefined ? {} : { matchedTarget: resolved.input.target }),
      })
      continue
    }
    const textPatch = externalLinkCacheTextPatch(reference, companionRows.rowsBySheetName)
    if (textPatch) {
      textPatches.push(textPatch)
    }
    refreshedBookIndices.push(bookIndex)
    refreshedSheetCount += referenceSheetCount
    refreshedCellCount += referenceCellCount
    referenceDiagnostics.push({
      bookIndex,
      ...(reference.workbookName === undefined ? {} : { workbookName: reference.workbookName }),
      ...(reference.target === undefined ? {} : { target: reference.target }),
      status: 'refreshed',
      candidateCount: resolved.candidateCount,
      ...(resolved.referenceCandidateCount === undefined ? {} : { referenceCandidateCount: resolved.referenceCandidateCount }),
      ...(resolved.matchKind === undefined ? {} : { matchKind: resolved.matchKind }),
      ...(resolved.input.fileName === undefined ? {} : { matchedFileName: resolved.input.fileName }),
      ...(resolved.input.workbookName === undefined ? {} : { matchedWorkbookName: resolved.input.workbookName }),
      ...(resolved.input.target === undefined ? {} : { matchedTarget: resolved.input.target }),
      refreshedSheetCount: referenceSheetCount,
      refreshedCellCount: referenceCellCount,
    })
  }
  const warnings = [
    ...(skippedNoMatchCount > 0 ? [externalWorkbookCompanionNoMatchWarning] : []),
    ...(skippedAmbiguousMatchCount > 0 ? [externalWorkbookCompanionAmbiguousMatchWarning] : []),
  ]
  return {
    rowsByAlias: output,
    textPatches,
    warnings,
    diagnostics: {
      externalWorkbookCount: externalWorkbooks.length,
      externalReferenceCount: targetsByBookIndex.size,
      refreshedBookIndices,
      refreshedSheetCount,
      refreshedCellCount,
      skippedNoMatchCount,
      skippedAmbiguousMatchCount,
      skippedEmptyRefreshCount,
      references: referenceDiagnostics,
    },
  }
}

function resolveExternalWorkbookInput(
  inputs: readonly XlsxExternalWorkbookInput[],
  referencesByBookIndex: ReadonlyMap<number, ExternalLinkReference>,
  reference: ExternalLinkReference,
): ResolvedExternalWorkbookInput {
  const exactTargetCandidates = inputs.filter((input) => externalWorkbookInputMatchesReferenceTarget(input, reference))
  if (exactTargetCandidates.length === 1) {
    return {
      input: exactTargetCandidates[0]!,
      status: 'matched',
      candidateCount: 1,
      matchKind: 'exact-target',
    }
  }
  if (exactTargetCandidates.length > 1) {
    return {
      status: 'skipped-ambiguous-match',
      candidateCount: exactTargetCandidates.length,
      matchKind: 'exact-target',
    }
  }
  const identityCandidates = inputs.filter((input) => externalWorkbookInputMatchesReferenceIdentity(input, reference))
  if (identityCandidates.length === 0) {
    return {
      status: 'skipped-no-match',
      candidateCount: 0,
    }
  }
  const referenceCandidateCount = externalReferenceIdentityCandidateCount(referencesByBookIndex, reference)
  if (identityCandidates.length > 1 || referenceCandidateCount !== 1) {
    return {
      status: 'skipped-ambiguous-match',
      candidateCount: identityCandidates.length,
      referenceCandidateCount,
      matchKind: 'unique-workbook-identity',
    }
  }
  return {
    input: identityCandidates[0]!,
    status: 'matched',
    candidateCount: 1,
    referenceCandidateCount,
    matchKind: 'unique-workbook-identity',
  }
}

function countExternalCachedCells(rows: ReadonlyMap<number, ReadonlyMap<number, PendingCellValue>>): number {
  let count = 0
  for (const row of rows.values()) {
    count += row.size
  }
  return count
}

interface CompanionWorkbookCachedRows {
  readonly rowsByAlias: ReadonlyMap<string, ReadonlyMap<number, Map<number, PendingCellValue>>>
  readonly rowsBySheetName: ReadonlyMap<string, ReadonlyMap<number, Map<number, PendingCellValue>>>
}

function readCompanionWorkbookCachedRowsByAlias(
  input: XlsxExternalWorkbookInput,
  bookIndex: number,
  targets: readonly ExternalReferenceScanTarget[],
): CompanionWorkbookCachedRows {
  const zip = readXlsxZipEntries(toUint8Array(input.bytes))
  const sheetsByName = new Map(workbookSheetPathEntriesForSource(zip).map((sheet) => [sheetNameKey(sheet.name), sheet]))
  const sharedStrings = readSharedStrings(zip)
  const output = new Map<string, ReadonlyMap<number, Map<number, PendingCellValue>>>()
  const rowsBySheetName = new Map<string, ReadonlyMap<number, Map<number, PendingCellValue>>>()
  for (const target of targets) {
    const sheet = sheetsByName.get(sheetNameKey(target.sheetName))
    const sheetXml = sheet ? getZipText(zip, sheet.path) : null
    if (!sheetXml) {
      continue
    }
    const rows = readExternalCachedRows(sheetXml, target.cellsByRow, true, sharedStrings)
    if (rows.size > 0) {
      output.set(externalReferenceAlias(bookIndex, target.sheetName), rows)
      rowsBySheetName.set(target.sheetName, rows)
    }
  }
  return { rowsByAlias: output, rowsBySheetName }
}

function externalLinkCacheTextPatch(
  reference: ExternalLinkReference,
  rowsBySheetName: ReadonlyMap<string, ReadonlyMap<number, Map<number, PendingCellValue>>>,
): XlsxSourceTextPatch | null {
  const rowsBySheetId = new Map<number, ReadonlyMap<number, Map<number, PendingCellValue>>>()
  for (const [sheetName, rows] of rowsBySheetName.entries()) {
    const sheetId = reference.sheetNames.findIndex((candidate) => sheetNameKey(candidate) === sheetNameKey(sheetName))
    if (sheetId >= 0) {
      rowsBySheetId.set(sheetId, rows)
    }
  }
  return rowsBySheetId.size === 0
    ? null
    : {
        path: reference.path,
        patchText: (text) => patchExternalLinkCacheXml(text, rowsBySheetId),
      }
}

function patchExternalLinkCacheXml(
  externalLinkXml: string,
  rowsBySheetId: ReadonlyMap<number, ReadonlyMap<number, Map<number, PendingCellValue>>>,
): string {
  const patchedSheetIds = new Set<number>()
  let nextXml = externalLinkXml.replace(
    /<((?:[A-Za-z_][\w.-]*:)?sheetData)\b((?:[^/>"']|"[^"]*"|'[^']*')*)(?:\/>|>([\s\S]*?)<\/\1>)/gu,
    (match: string, tagName: string, attributes: string, body: string | undefined) => {
      const sheetIdText = readXmlAttribute(attributes, 'sheetId')
      const sheetId = sheetIdText === null ? Number.NaN : Number(sheetIdText)
      const rows = Number.isSafeInteger(sheetId) ? rowsBySheetId.get(sheetId) : undefined
      if (!rows) {
        return match
      }
      patchedSheetIds.add(sheetId)
      return `<${tagName}${attributes}>${patchExternalSheetDataRows(body ?? '', rows)}</${tagName}>`
    },
  )
  const missingSheetDataXml = [...rowsBySheetId.entries()]
    .filter(([sheetId]) => !patchedSheetIds.has(sheetId))
    .toSorted((left, right) => left[0] - right[0])
    .map(([sheetId, rows]) => `<sheetData sheetId="${String(sheetId)}">${externalCacheRowsXml(rows)}</sheetData>`)
    .join('')
  if (missingSheetDataXml.length === 0) {
    return nextXml
  }
  if (/<(?:[A-Za-z_][\w.-]*:)?sheetDataSet\b/u.test(nextXml)) {
    return nextXml.replace(/<\/((?:[A-Za-z_][\w.-]*:)?sheetDataSet)>/u, `${missingSheetDataXml}</$1>`)
  }
  nextXml = nextXml.replace(/<\/((?:[A-Za-z_][\w.-]*:)?externalBook)>/u, `<sheetDataSet>${missingSheetDataXml}</sheetDataSet></$1>`)
  return nextXml
}

function patchExternalSheetDataRows(sheetDataBody: string, rows: ReadonlyMap<number, ReadonlyMap<number, PendingCellValue>>): string {
  const patchedRows = new Set<number>()
  let nextBody = sheetDataBody.replace(
    /<((?:[A-Za-z_][\w.-]*:)?row)\b((?:[^>"']|"[^"]*"|'[^']*')*)(?:\/>|>([\s\S]*?)<\/\1>)/gu,
    (match: string, tagName: string, attributes: string, body: string | undefined) => {
      const rowText = readXmlAttribute(attributes, 'r')
      const rowNumber = rowText === null ? Number.NaN : Number(rowText)
      const row = Number.isSafeInteger(rowNumber) && rowNumber > 0 ? rowNumber - 1 : Number.NaN
      const cells = Number.isSafeInteger(row) ? rows.get(row) : undefined
      if (!cells) {
        return match
      }
      patchedRows.add(row)
      return `<${tagName}${attributes}>${patchExternalRowCells(row, body ?? '', cells)}</${tagName}>`
    },
  )
  for (const [row, cells] of [...rows.entries()].toSorted((left, right) => left[0] - right[0])) {
    if (!patchedRows.has(row)) {
      nextBody = insertExternalCacheRowXml(nextBody, row, cells)
    }
  }
  return nextBody
}

function patchExternalRowCells(row: number, rowBody: string, cells: ReadonlyMap<number, PendingCellValue>): string {
  const patchedCols = new Set<number>()
  let nextBody = rowBody.replace(
    /<((?:[A-Za-z_][\w.-]*:)?(?:cell|c))\b((?:[^>"']|"[^"]*"|'[^']*')*)(?:\/>|>([\s\S]*?)<\/\1>)/gu,
    (match: string, _tagName: string, attributes: string) => {
      const addressText = readXmlAttribute(attributes, 'r')
      let address
      try {
        address = addressText ? decodeCellAddress(addressText.replaceAll('$', '')) : undefined
      } catch {
        address = undefined
      }
      if (!address) {
        return match
      }
      const value = cells.get(address.c)
      if (value === undefined) {
        return match
      }
      patchedCols.add(address.c)
      return externalCacheCellXml(row, address.c, value)
    },
  )
  for (const [col, value] of [...cells.entries()].toSorted((left, right) => left[0] - right[0])) {
    if (!patchedCols.has(col)) {
      nextBody = insertExternalCacheCellXml(nextBody, row, col, value)
    }
  }
  return nextBody
}

function insertExternalCacheRowXml(sheetDataBody: string, row: number, cells: ReadonlyMap<number, PendingCellValue>): string {
  const rowXml = `<row r="${String(row + 1)}">${externalCacheCellsXml(row, cells)}</row>`
  for (const match of sheetDataBody.matchAll(/<((?:[A-Za-z_][\w.-]*:)?row)\b((?:[^>"']|"[^"]*"|'[^']*')*)(?:\/>|>[\s\S]*?<\/\1>)/gu)) {
    const rowText = readXmlAttribute(match[2] ?? '', 'r')
    const rowNumber = rowText === null ? Number.NaN : Number(rowText)
    if (Number.isSafeInteger(rowNumber) && rowNumber > row + 1) {
      return `${sheetDataBody.slice(0, match.index)}${rowXml}${sheetDataBody.slice(match.index)}`
    }
  }
  return `${sheetDataBody}${rowXml}`
}

function insertExternalCacheCellXml(rowBody: string, row: number, col: number, value: PendingCellValue): string {
  const cellXml = externalCacheCellXml(row, col, value)
  for (const match of rowBody.matchAll(/<((?:[A-Za-z_][\w.-]*:)?(?:cell|c))\b((?:[^>"']|"[^"]*"|'[^']*')*)(?:\/>|>[\s\S]*?<\/\1>)/gu)) {
    const addressText = readXmlAttribute(match[2] ?? '', 'r')
    let address
    try {
      address = addressText ? decodeCellAddress(addressText.replaceAll('$', '')) : undefined
    } catch {
      address = undefined
    }
    if (address && address.c > col) {
      return `${rowBody.slice(0, match.index)}${cellXml}${rowBody.slice(match.index)}`
    }
  }
  return `${rowBody}${cellXml}`
}

function externalCacheRowsXml(rows: ReadonlyMap<number, ReadonlyMap<number, PendingCellValue>>): string {
  return [...rows.entries()]
    .toSorted((left, right) => left[0] - right[0])
    .map(([row, cells]) => `<row r="${String(row + 1)}">${externalCacheCellsXml(row, cells)}</row>`)
    .join('')
}

function externalCacheCellsXml(row: number, cells: ReadonlyMap<number, PendingCellValue>): string {
  return [...cells.entries()]
    .toSorted((left, right) => left[0] - right[0])
    .map(([col, value]) => externalCacheCellXml(row, col, value))
    .join('')
}

function externalCacheCellXml(row: number, col: number, value: PendingCellValue): string {
  const address = escapeXmlAttribute(encodeCellAddress({ r: row, c: col }))
  if ('kind' in value) {
    return `<cell r="${address}"/>`
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return `<cell r="${address}"/>`
    case ValueTag.Number:
      return Number.isFinite(value.value) ? `<cell r="${address}"><v>${String(value.value)}</v></cell>` : `<cell r="${address}"/>`
    case ValueTag.Boolean:
      return `<cell r="${address}" t="b"><v>${value.value ? '1' : '0'}</v></cell>`
    case ValueTag.String:
      return `<cell r="${address}" t="str"><v>${escapeXmlText(value.value)}</v></cell>`
    case ValueTag.Error:
      return `<cell r="${address}" t="e"><v>${escapeXmlText(formatErrorCode(value.code))}</v></cell>`
  }
}

function readSharedStrings(zip: XlsxZipEntries): readonly string[] {
  const sharedStringsXml = getZipText(zip, 'xl/sharedStrings.xml')
  if (!sharedStringsXml) {
    return []
  }
  return [...sharedStringsXml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?si)\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/\1>/gu)].map((match) =>
    readTextRuns(match[2] ?? ''),
  )
}

function readExternalCachedRows(
  xml: string,
  targetCellsByRow: ReadonlyMap<number, ReadonlySet<number>>,
  fillMissingCells: boolean,
  sharedStrings: readonly string[] = [],
): Map<number, Map<number, PendingCellValue>> {
  const output = new Map<number, Map<number, PendingCellValue>>()
  if (fillMissingCells) {
    for (const [row, cols] of targetCellsByRow.entries()) {
      output.set(row, new Map([...cols].map((col) => [col, emptyCellValue] as const)))
    }
  }
  for (const rowMatch of xml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?row)\b((?:[^>"']|"[^"]*"|'[^']*')*)>([\s\S]*?)<\/\1>/gu)) {
    const rowText = readXmlAttribute(rowMatch[2] ?? '', 'r')
    const rowNumber = rowText === null ? Number.NaN : Number(rowText)
    if (!Number.isSafeInteger(rowNumber) || rowNumber < 1) {
      continue
    }
    const row = rowNumber - 1
    const targetCols = targetCellsByRow.get(row)
    if (!targetCols) {
      continue
    }
    const rowValues = output.get(row) ?? new Map<number, PendingCellValue>()
    for (const cellMatch of (rowMatch[3] ?? '').matchAll(
      /<((?:[A-Za-z_][\w.-]*:)?(?:cell|c))\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?(?:cell|c))\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\2>/gu,
    )) {
      const cellXml = cellMatch[0]
      const openingTag = /<(?:[A-Za-z_][\w.-]*:)?(?:cell|c)\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(cellXml)?.[0]
      const addressText = openingTag ? readXmlAttribute(openingTag, 'r') : null
      if (!openingTag || !addressText) {
        continue
      }
      let address
      try {
        address = decodeCellAddress(addressText.replaceAll('$', ''))
      } catch {
        continue
      }
      if (!targetCols.has(address.c)) {
        continue
      }
      rowValues.set(address.c, readExternalCachedCellValue(cellXml, openingTag, sharedStrings))
    }
    output.set(row, rowValues)
  }
  return output
}

function readExternalCachedCellValue(cellXml: string, openingTag: string, sharedStrings: readonly string[] = []): PendingCellValue {
  const type = readXmlAttribute(openingTag, 't')
  if (type === 'inlineStr') {
    return stringCellValue(readTextRuns(cellXml))
  }
  const rawValue = readElementText(cellXml, 'v')
  if (rawValue === null) {
    return emptyCellValue
  }
  if (type === 'str') {
    return stringCellValue(decodeXmlText(rawValue))
  }
  if (type === 's') {
    const index = Number(decodeXmlText(rawValue).trim())
    const value = Number.isSafeInteger(index) && index >= 0 ? sharedStrings[index] : undefined
    return value === undefined ? emptyCellValue : stringCellValue(value)
  }
  if (type === 'b') {
    return { tag: ValueTag.Boolean, value: rawValue === '1' || rawValue.toLowerCase() === 'true' }
  }
  if (type === 'e') {
    return { tag: ValueTag.Error, code: errorCodeForText(decodeXmlText(rawValue)) }
  }
  const numeric = Number(rawValue)
  return Number.isFinite(numeric) ? { tag: ValueTag.Number, value: numeric } : stringCellValue(decodeXmlText(rawValue))
}

function toUint8Array(bytes: Uint8Array | ArrayBuffer): Uint8Array {
  return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
}

function normalizedWorkbookTarget(value: string | undefined): string | null {
  if (!value) {
    return null
  }
  const withoutFragment = value.split('#')[0]?.trim()
  if (!withoutFragment) {
    return null
  }
  const normalized = withoutFragment.replace(/\\/gu, '/').replace(/^file:\/+/iu, '/')
  try {
    return decodeURIComponent(normalized).toLocaleLowerCase('en-US')
  } catch {
    return normalized.toLocaleLowerCase('en-US')
  }
}

function workbookIdentityFromName(value: string | undefined): string | null {
  const target = normalizedWorkbookTarget(value)
  if (!target) {
    return null
  }
  const segments = target.split('/').filter((segment) => segment.length > 0)
  return segments.at(-1) ?? target
}

function externalWorkbookInputIdentityNames(input: XlsxExternalWorkbookInput): ReadonlySet<string> {
  return new Set(
    [input.workbookName, input.fileName, input.target]
      .map(workbookIdentityFromName)
      .filter((value): value is string => value !== null && value.length > 0),
  )
}

function externalReferenceIdentityNames(reference: ExternalLinkReference): ReadonlySet<string> {
  return new Set(
    [reference.workbookName, reference.target]
      .map(workbookIdentityFromName)
      .filter((value): value is string => value !== null && value.length > 0),
  )
}

function externalWorkbookInputMatchesReferenceTarget(input: XlsxExternalWorkbookInput, reference: ExternalLinkReference): boolean {
  const inputTarget = normalizedWorkbookTarget(input.target)
  const referenceTarget = normalizedWorkbookTarget(reference.target)
  return Boolean(inputTarget && referenceTarget && inputTarget === referenceTarget)
}

function externalWorkbookInputMatchesReferenceIdentity(input: XlsxExternalWorkbookInput, reference: ExternalLinkReference): boolean {
  const inputTarget = normalizedWorkbookTarget(input.target)
  const referenceTarget = normalizedWorkbookTarget(reference.target)
  if (inputTarget && referenceTarget) {
    return false
  }
  const inputNames = externalWorkbookInputIdentityNames(input)
  if (inputNames.size === 0) {
    return false
  }
  for (const referenceName of externalReferenceIdentityNames(reference)) {
    if (inputNames.has(referenceName)) {
      return true
    }
  }
  return false
}

function externalReferenceIdentityCandidateCount(
  referencesByBookIndex: ReadonlyMap<number, ExternalLinkReference>,
  reference: ExternalLinkReference,
): number {
  const referenceNames = externalReferenceIdentityNames(reference)
  if (referenceNames.size === 0) {
    return 0
  }
  let count = 0
  for (const candidate of referencesByBookIndex.values()) {
    for (const referenceName of externalReferenceIdentityNames(candidate)) {
      if (referenceNames.has(referenceName)) {
        count += 1
        break
      }
    }
  }
  return count
}

function resolveTargetPath(fromPath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target)
  }
  const baseParts = normalizeZipPath(fromPath).split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part.length === 0 || part === '.') {
      continue
    }
    if (part === '..') {
      baseParts.pop()
      continue
    }
    baseParts.push(part)
  }
  return normalizeZipPath(baseParts.join('/'))
}

function quoteFormulaSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function externalReferenceAlias(bookIndex: number, sheetName: string): string {
  let encodedSheetName = ''
  for (const character of sheetName) {
    encodedSheetName += /^[A-Za-z0-9_]$/u.test(character) ? character : `_u${character.codePointAt(0)!.toString(16).padStart(4, '0')}_`
  }
  return `__bilig_external_${String(bookIndex)}_${encodedSheetName}`
}

function sheetNameKey(sheetName: string): string {
  return sheetName.toLocaleLowerCase('en-US')
}

function stringCellValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function errorCodeForText(value: string): ErrorCode {
  switch (value.toUpperCase()) {
    case '#DIV/0!':
      return ErrorCode.Div0
    case '#REF!':
      return ErrorCode.Ref
    case '#VALUE!':
      return ErrorCode.Value
    case '#NAME?':
      return ErrorCode.Name
    case '#N/A':
      return ErrorCode.NA
    case '#NUM!':
      return ErrorCode.Num
    case '#FIELD!':
      return ErrorCode.Field
    case '#NULL!':
      return ErrorCode.Null
    default:
      return ErrorCode.Value
  }
}

function readTextRuns(xml: string): string {
  const runs = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)].map(
    (match) => decodeXmlText(match[1] ?? ''),
  )
  return runs.length > 0 ? runs.join('') : ''
}

function readElementText(xml: string, name: string): string | null {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  return (
    new RegExp(
      `<(?:[A-Za-z_][\\w.-]*:)?${escapedName}\\b(?:[^>"']|"[^"]*"|'[^']*')*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${escapedName}>`,
      'u',
    ).exec(xml)?.[1] ?? null
  )
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}
