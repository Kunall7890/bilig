import { parseCellAddress, translateFormulaReferences } from '@bilig/formula'
import type { CellRangeRef, WorkbookSnapshot } from '@bilig/protocol'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { findWorkbookFormulaIssues, summarizeWorkbookStructure, type WorkbookFormulaIssue } from './workbook-agent-comprehension.js'
import { clampAuditLimit, MAX_AUDIT_LIMIT } from './workbook-agent-audit-limits.js'
export { type WorkbookInvariantVerificationReport, verifyWorkbookInvariants } from './workbook-agent-invariants.js'
export {
  scanWorkbookUsedRangeBloat,
  type WorkbookUsedRangeBloatReport,
  type WorkbookUsedRangeBloatSheetReport,
} from './workbook-agent-audit-used-range.js'

const DEFAULT_HIDDEN_PRECEDENT_DEPTH = 4
const MAX_HIDDEN_PRECEDENT_DEPTH = 6
const MAX_HIDDEN_PRECEDENT_NODES = 160
const MIN_INCONSISTENT_GROUP_SIZE = 3

interface FormulaCellRef {
  sheetName: string
  address: string
  formula: string
  row: number
  col: number
}

interface FormulaRunGroup {
  axis: 'row' | 'column'
  sheetName: string
  cells: readonly FormulaCellRef[]
}

export interface WorkbookBrokenReferenceReport {
  summary: {
    scannedFormulaCells: number
    brokenReferenceCount: number
    truncated: boolean
  }
  issues: WorkbookFormulaIssue[]
}

export interface WorkbookHiddenRowDependencyHit {
  sheetName: string
  address: string
  rowNumber: number
  depth: number
}

export interface WorkbookHiddenRowDependencyIssue {
  sheetName: string
  address: string
  formula: string
  hiddenPrecedentCount: number
  hiddenPrecedents: WorkbookHiddenRowDependencyHit[]
}

export interface WorkbookHiddenRowDependencyReport {
  summary: {
    scannedFormulaCells: number
    affectedFormulaCount: number
    hiddenPrecedentCount: number
    truncated: boolean
  }
  issues: WorkbookHiddenRowDependencyIssue[]
}

export interface WorkbookInconsistentFormulaOutlier {
  address: string
  actualFormula: string
  expectedFormula: string
}

export interface WorkbookInconsistentFormulaGroupReport {
  axis: 'row' | 'column'
  sheetName: string
  groupRange: CellRangeRef
  formulaCellCount: number
  dominantFormula: string
  dominantCount: number
  outliers: WorkbookInconsistentFormulaOutlier[]
}

export interface WorkbookInconsistentFormulaReport {
  summary: {
    scannedFormulaCells: number
    inconsistentGroupCount: number
    outlierCount: number
    truncated: boolean
  }
  groups: WorkbookInconsistentFormulaGroupReport[]
}

export interface WorkbookPerformanceHotspot {
  sheetName: string
  cellCount: number
  formulaCellCount: number
  jsOnlyFormulaCount: number
  issueCount: number
  pivotCount: number
  spillCount: number
  usedRange: {
    startAddress: string
    endAddress: string
  } | null
  reasons: string[]
}

export interface WorkbookPerformanceHotspotReport {
  summary: {
    scannedSheetCount: number
    hotspotCount: number
    truncated: boolean
    recalcMetrics: ReturnType<WorkbookRuntime['engine']['getLastMetrics']>
  }
  hotspots: WorkbookPerformanceHotspot[]
}

function clampHiddenDepth(depth: number | undefined): number {
  if (!Number.isFinite(depth) || typeof depth !== 'number') {
    return DEFAULT_HIDDEN_PRECEDENT_DEPTH
  }
  return Math.max(1, Math.min(MAX_HIDDEN_PRECEDENT_DEPTH, Math.trunc(depth)))
}

function collectFormulaCells(snapshot: WorkbookSnapshot, sheetName?: string): readonly FormulaCellRef[] {
  const cells: FormulaCellRef[] = []
  for (const sheet of snapshot.sheets) {
    if (sheetName !== undefined && sheet.name !== sheetName) {
      continue
    }
    for (const cell of sheet.cells) {
      if (!cell.formula) {
        continue
      }
      const parsed = parseCellAddress(cell.address, sheet.name)
      cells.push({
        sheetName: sheet.name,
        address: cell.address,
        formula: cell.formula,
        row: parsed.row,
        col: parsed.col,
      })
    }
  }
  return cells
}

function splitQualifiedAddress(qualifiedAddress: string): {
  sheetName: string
  address: string
} {
  const separator = qualifiedAddress.lastIndexOf('!')
  if (separator <= 0 || separator >= qualifiedAddress.length - 1) {
    throw new Error(`Invalid qualified workbook address: ${qualifiedAddress}`)
  }
  return {
    sheetName: qualifiedAddress.slice(0, separator),
    address: qualifiedAddress.slice(separator + 1),
  }
}

function buildHiddenRowIntervals(runtime: WorkbookRuntime): Map<string, Array<[number, number]>> {
  const hiddenRows = new Map<string, Array<[number, number]>>()
  const snapshot = runtime.engine.exportSnapshot()
  for (const sheet of snapshot.sheets) {
    const intervals = runtime.engine
      .getRowMetadata(sheet.name)
      .filter((entry) => entry.hidden === true)
      .map((entry) => [entry.start, entry.start + entry.count - 1] as [number, number])
    if (intervals.length > 0) {
      hiddenRows.set(sheet.name, intervals)
    }
  }
  return hiddenRows
}

function isRowHidden(hiddenRows: readonly [number, number][], row: number): boolean {
  return hiddenRows.some(([start, end]) => row >= start && row <= end)
}

function collectHiddenPrecedents(
  runtime: WorkbookRuntime,
  hiddenRowsBySheet: ReadonlyMap<string, readonly [number, number][]>,
  rootSheetName: string,
  rootAddress: string,
  maxDepth: number,
): WorkbookHiddenRowDependencyHit[] {
  const hiddenHits: WorkbookHiddenRowDependencyHit[] = []
  const queue: Array<{ qualifiedAddress: string; depth: number }> = [{ qualifiedAddress: `${rootSheetName}!${rootAddress}`, depth: 0 }]
  const visited = new Set<string>(queue.map((entry) => entry.qualifiedAddress))
  let visitedNodes = 0
  while (queue.length > 0) {
    const current = queue.shift()
    if (!current || current.depth >= maxDepth) {
      continue
    }
    const { sheetName, address } = splitQualifiedAddress(current.qualifiedAddress)
    const dependencies = runtime.engine.getDependencies(sheetName, address)
    for (const precedent of dependencies.directPrecedents) {
      if (visited.has(precedent)) {
        continue
      }
      visited.add(precedent)
      visitedNodes += 1
      if (visitedNodes > MAX_HIDDEN_PRECEDENT_NODES) {
        return hiddenHits
      }
      const parsed = splitQualifiedAddress(precedent)
      const location = parseCellAddress(parsed.address, parsed.sheetName)
      const hiddenRows = hiddenRowsBySheet.get(parsed.sheetName)
      if (hiddenRows && isRowHidden(hiddenRows, location.row)) {
        hiddenHits.push({
          sheetName: parsed.sheetName,
          address: parsed.address,
          rowNumber: location.row + 1,
          depth: current.depth + 1,
        })
      }
      queue.push({
        qualifiedAddress: precedent,
        depth: current.depth + 1,
      })
    }
  }
  return hiddenHits
}

function buildFormulaRunGroups(snapshot: WorkbookSnapshot, axis: 'row' | 'column', sheetName?: string): readonly FormulaRunGroup[] {
  const groups: FormulaRunGroup[] = []
  for (const sheet of snapshot.sheets) {
    if (sheetName !== undefined && sheet.name !== sheetName) {
      continue
    }
    const formulaCells = collectFormulaCells(
      {
        ...snapshot,
        sheets: [sheet],
      },
      sheet.name,
    )
    const cellsByAxis = new Map<number, FormulaCellRef[]>()
    for (const cell of formulaCells) {
      const key = axis === 'column' ? cell.col : cell.row
      const entries = cellsByAxis.get(key) ?? []
      entries.push(cell)
      cellsByAxis.set(key, entries)
    }
    for (const entries of cellsByAxis.values()) {
      const sorted = [...entries].toSorted((left, right) => (axis === 'column' ? left.row - right.row : left.col - right.col))
      let current: FormulaCellRef[] = []
      let lastIndex = Number.NaN
      for (const cell of sorted) {
        const index = axis === 'column' ? cell.row : cell.col
        if (current.length === 0 || index === lastIndex + 1) {
          current.push(cell)
          lastIndex = index
          continue
        }
        if (current.length >= MIN_INCONSISTENT_GROUP_SIZE) {
          groups.push({
            axis,
            sheetName: sheet.name,
            cells: [...current],
          })
        }
        current = [cell]
        lastIndex = index
      }
      if (current.length >= MIN_INCONSISTENT_GROUP_SIZE) {
        groups.push({
          axis,
          sheetName: sheet.name,
          cells: [...current],
        })
      }
    }
  }
  return groups
}

function normalizeFormulaForAnchor(formula: string, rowDelta: number, colDelta: number): string {
  try {
    return translateFormulaReferences(formula, rowDelta, colDelta)
  } catch {
    return `__raw__:${formula}`
  }
}

function translateNormalizedFormula(normalizedFormula: string, rowDelta: number, colDelta: number, fallbackFormula: string): string {
  if (normalizedFormula.startsWith('__raw__:')) {
    return `=${fallbackFormula}`
  }
  try {
    return `=${translateFormulaReferences(normalizedFormula, rowDelta, colDelta)}`
  } catch {
    return `=${fallbackFormula}`
  }
}

function summarizeGroupRange(group: FormulaRunGroup): CellRangeRef {
  const first = group.cells[0]
  const last = group.cells[group.cells.length - 1]
  if (!first || !last) {
    throw new Error('Formula run group must contain at least one cell')
  }
  return {
    sheetName: group.sheetName,
    startAddress: first.address,
    endAddress: last.address,
  }
}

export function scanWorkbookBrokenReferences(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined
    limit?: number | undefined
  } = {},
): WorkbookBrokenReferenceReport {
  const limit = clampAuditLimit(input.limit)
  const report = findWorkbookFormulaIssues(runtime, {
    ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
    limit: MAX_AUDIT_LIMIT,
  })
  const issues = report.issues.filter((issue) => issue.errorText === '#REF!')
  return {
    summary: {
      scannedFormulaCells: report.summary.scannedFormulaCells,
      brokenReferenceCount: issues.length,
      truncated: issues.length > limit,
    },
    issues: issues.slice(0, limit),
  }
}

export function scanWorkbookHiddenRowsAffectingResults(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined
    limit?: number | undefined
    depth?: number | undefined
  } = {},
): WorkbookHiddenRowDependencyReport {
  const snapshot = runtime.engine.exportSnapshot()
  const hiddenRowsBySheet = buildHiddenRowIntervals(runtime)
  const depth = clampHiddenDepth(input.depth)
  const limit = clampAuditLimit(input.limit)
  const issues: WorkbookHiddenRowDependencyIssue[] = []
  let scannedFormulaCells = 0
  let hiddenPrecedentCount = 0

  for (const cell of collectFormulaCells(snapshot, input.sheetName)) {
    scannedFormulaCells += 1
    const hits = collectHiddenPrecedents(runtime, hiddenRowsBySheet, cell.sheetName, cell.address, depth)
    if (hits.length === 0) {
      continue
    }
    hiddenPrecedentCount += hits.length
    issues.push({
      sheetName: cell.sheetName,
      address: cell.address,
      formula: `=${cell.formula}`,
      hiddenPrecedentCount: hits.length,
      hiddenPrecedents: hits.slice(0, 10),
    })
  }

  const sortedIssues = issues.toSorted((left, right) => {
    if (left.hiddenPrecedentCount !== right.hiddenPrecedentCount) {
      return right.hiddenPrecedentCount - left.hiddenPrecedentCount
    }
    if (left.sheetName !== right.sheetName) {
      return left.sheetName.localeCompare(right.sheetName)
    }
    return left.address.localeCompare(right.address, undefined, { numeric: true })
  })

  return {
    summary: {
      scannedFormulaCells,
      affectedFormulaCount: sortedIssues.length,
      hiddenPrecedentCount,
      truncated: sortedIssues.length > limit,
    },
    issues: sortedIssues.slice(0, limit),
  }
}

export function scanWorkbookInconsistentFormulas(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined
    limit?: number | undefined
  } = {},
): WorkbookInconsistentFormulaReport {
  const snapshot = runtime.engine.exportSnapshot()
  const limit = clampAuditLimit(input.limit)
  const groups: WorkbookInconsistentFormulaGroupReport[] = []
  const candidateGroups = [
    ...buildFormulaRunGroups(snapshot, 'column', input.sheetName),
    ...buildFormulaRunGroups(snapshot, 'row', input.sheetName),
  ]

  for (const group of candidateGroups) {
    const anchor = group.cells[0]
    if (!anchor) {
      continue
    }
    const signatureMap = new Map<
      string,
      {
        count: number
        representative: FormulaCellRef
      }
    >()
    const signatures = group.cells.map((cell) => {
      const signature = normalizeFormulaForAnchor(cell.formula, anchor.row - cell.row, anchor.col - cell.col)
      const existing = signatureMap.get(signature)
      if (existing) {
        existing.count += 1
      } else {
        signatureMap.set(signature, {
          count: 1,
          representative: cell,
        })
      }
      return {
        cell,
        signature,
      }
    })

    const dominantEntry = [...signatureMap.entries()].toSorted((left, right) => {
      if (left[1].count !== right[1].count) {
        return right[1].count - left[1].count
      }
      return left[1].representative.address.localeCompare(right[1].representative.address, undefined, {
        numeric: true,
      })
    })[0]
    if (!dominantEntry || dominantEntry[1].count < 2) {
      continue
    }
    const [dominantSignature, dominantData] = dominantEntry
    const outliers = signatures
      .filter((entry) => entry.signature !== dominantSignature)
      .map((entry) => ({
        address: entry.cell.address,
        actualFormula: `=${entry.cell.formula}`,
        expectedFormula: translateNormalizedFormula(
          dominantSignature,
          entry.cell.row - anchor.row,
          entry.cell.col - anchor.col,
          dominantData.representative.formula,
        ),
      }))
    if (outliers.length === 0) {
      continue
    }
    groups.push({
      axis: group.axis,
      sheetName: group.sheetName,
      groupRange: summarizeGroupRange(group),
      formulaCellCount: group.cells.length,
      dominantFormula: `=${dominantData.representative.formula}`,
      dominantCount: dominantData.count,
      outliers,
    })
  }

  const sortedGroups = groups.toSorted((left, right) => {
    if (left.outliers.length !== right.outliers.length) {
      return right.outliers.length - left.outliers.length
    }
    if (left.formulaCellCount !== right.formulaCellCount) {
      return right.formulaCellCount - left.formulaCellCount
    }
    if (left.sheetName !== right.sheetName) {
      return left.sheetName.localeCompare(right.sheetName)
    }
    return left.groupRange.startAddress.localeCompare(right.groupRange.startAddress, undefined, {
      numeric: true,
    })
  })

  return {
    summary: {
      scannedFormulaCells: collectFormulaCells(snapshot, input.sheetName).length,
      inconsistentGroupCount: sortedGroups.length,
      outlierCount: sortedGroups.reduce((sum, group) => sum + group.outliers.length, 0),
      truncated: sortedGroups.length > limit,
    },
    groups: sortedGroups.slice(0, limit),
  }
}

export function scanWorkbookPerformanceHotspots(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined
    limit?: number | undefined
  } = {},
): WorkbookPerformanceHotspotReport {
  const limit = clampAuditLimit(input.limit)
  const structure = summarizeWorkbookStructure(runtime)
  const formulaIssues = findWorkbookFormulaIssues(runtime, {
    ...(input.sheetName !== undefined ? { sheetName: input.sheetName } : {}),
    limit: MAX_AUDIT_LIMIT,
  })
  const jsOnlyBySheet = new Map<string, number>()
  const issueCountBySheet = new Map<string, number>()
  for (const issue of formulaIssues.issues) {
    issueCountBySheet.set(issue.sheetName, (issueCountBySheet.get(issue.sheetName) ?? 0) + 1)
    if (issue.issueKinds.includes('unsupported')) {
      jsOnlyBySheet.set(issue.sheetName, (jsOnlyBySheet.get(issue.sheetName) ?? 0) + 1)
    }
  }

  const hotspots = structure.sheets
    .filter((sheet) => input.sheetName === undefined || sheet.name === input.sheetName)
    .map((sheet) => {
      const jsOnlyFormulaCount = jsOnlyBySheet.get(sheet.name) ?? 0
      const issueCount = issueCountBySheet.get(sheet.name) ?? 0
      const reasons: string[] = []
      if (jsOnlyFormulaCount > 0) {
        reasons.push(`${String(jsOnlyFormulaCount)} JS-only formula${jsOnlyFormulaCount === 1 ? '' : 's'}`)
      }
      if (sheet.pivotCount > 0) {
        reasons.push(`${String(sheet.pivotCount)} pivot output${sheet.pivotCount === 1 ? '' : 's'}`)
      }
      if (sheet.spillCount > 0) {
        reasons.push(`${String(sheet.spillCount)} spill range${sheet.spillCount === 1 ? '' : 's'}`)
      }
      if (sheet.formulaCellCount > 0) {
        reasons.push(`${String(sheet.formulaCellCount)} formula cell${sheet.formulaCellCount === 1 ? '' : 's'}`)
      }
      if (issueCount > 0) {
        reasons.push(`${String(issueCount)} formula issue${issueCount === 1 ? '' : 's'}`)
      }
      return {
        sheetName: sheet.name,
        cellCount: sheet.cellCount,
        formulaCellCount: sheet.formulaCellCount,
        jsOnlyFormulaCount,
        issueCount,
        pivotCount: sheet.pivotCount,
        spillCount: sheet.spillCount,
        usedRange: sheet.usedRange,
        reasons,
      } satisfies WorkbookPerformanceHotspot
    })
    .filter((sheet) => sheet.reasons.length > 0)
    .toSorted((left, right) => {
      if (left.jsOnlyFormulaCount !== right.jsOnlyFormulaCount) {
        return right.jsOnlyFormulaCount - left.jsOnlyFormulaCount
      }
      if (left.pivotCount !== right.pivotCount) {
        return right.pivotCount - left.pivotCount
      }
      if (left.formulaCellCount !== right.formulaCellCount) {
        return right.formulaCellCount - left.formulaCellCount
      }
      if (left.spillCount !== right.spillCount) {
        return right.spillCount - left.spillCount
      }
      if (left.issueCount !== right.issueCount) {
        return right.issueCount - left.issueCount
      }
      if (left.cellCount !== right.cellCount) {
        return right.cellCount - left.cellCount
      }
      return left.sheetName.localeCompare(right.sheetName)
    })

  return {
    summary: {
      scannedSheetCount:
        input.sheetName === undefined
          ? structure.summary.sheetCount
          : structure.sheets.filter((sheet) => sheet.name === input.sheetName).length,
      hotspotCount: hotspots.length,
      truncated: hotspots.length > limit,
      recalcMetrics: runtime.engine.getLastMetrics(),
    },
    hotspots: hotspots.slice(0, limit),
  }
}
