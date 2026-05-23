import { SpreadsheetEngine } from '@bilig/core'
import type { CellRangeRef, CellSnapshot, WorkbookSnapshot, WorkbookTableSnapshot } from '@bilig/protocol'
import {
  applyWorkbookAgentCommandBundle,
  describeWorkbookAgentCommand,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewCellDiff,
  type WorkbookAgentPreviewSemanticTarget,
  type WorkbookAgentPreviewSummary,
} from './workbook-agent-bundles.js'
import { toWorkbookCommandBundle } from './workbook-agent-command-handoff.js'
import { formatAddress, parseCellAddress } from '@bilig/formula'

const MAX_PREVIEW_DIFFS = 64

function buildChangeKinds(beforeCell: CellSnapshot, afterCell: CellSnapshot): WorkbookAgentPreviewCellDiff['changeKinds'] {
  return [
    ...(beforeCell.formula !== afterCell.formula ? (['formula'] as const) : []),
    ...(beforeCell.input !== afterCell.input ? (['input'] as const) : []),
    ...(beforeCell.styleId !== afterCell.styleId ? (['style'] as const) : []),
    ...(beforeCell.format !== afterCell.format ? (['numberFormat'] as const) : []),
  ]
}

function cloneCellDiff(beforeCell: CellSnapshot, afterCell: CellSnapshot): WorkbookAgentPreviewCellDiff | null {
  const changeKinds = buildChangeKinds(beforeCell, afterCell)
  if (changeKinds.length === 0) {
    return null
  }
  return {
    sheetName: afterCell.sheetName,
    address: afterCell.address,
    beforeInput: beforeCell.input ?? null,
    beforeFormula: beforeCell.formula ? `=${beforeCell.formula}` : null,
    afterInput: afterCell.input ?? null,
    afterFormula: afterCell.formula ? `=${afterCell.formula}` : null,
    changeKinds,
  }
}

function collectTargetAddresses(bundle: WorkbookAgentCommandBundle): readonly {
  sheetName: string
  address: string
}[] {
  const addresses: Array<{ sheetName: string; address: string }> = []
  bundle.affectedRanges
    .filter((range) => range.role === 'target')
    .forEach((range) => {
      const start = parseCellAddress(range.startAddress, range.sheetName)
      const end = parseCellAddress(range.endAddress, range.sheetName)
      const rowStart = Math.min(start.row, end.row)
      const rowEnd = Math.max(start.row, end.row)
      const colStart = Math.min(start.col, end.col)
      const colEnd = Math.max(start.col, end.col)
      for (let row = rowStart; row <= rowEnd && addresses.length < MAX_PREVIEW_DIFFS; row += 1) {
        for (let col = colStart; col <= colEnd && addresses.length < MAX_PREVIEW_DIFFS; col += 1) {
          const address = formatAddress(row, col)
          if (!addresses.some((entry) => entry.sheetName === range.sheetName && entry.address === address)) {
            addresses.push({ sheetName: range.sheetName, address })
          }
        }
      }
    })
  return addresses
}

function buildStructuralChanges(bundle: WorkbookAgentCommandBundle): string[] {
  const structuralChanges: string[] = []
  bundle.commands.forEach((command) => {
    if (
      command.kind === 'createSheet' ||
      command.kind === 'renameSheet' ||
      command.kind === 'updateRowMetadata' ||
      command.kind === 'updateColumnMetadata'
    ) {
      const description = describeWorkbookAgentCommand(command)
      if (!structuralChanges.includes(description)) {
        structuralChanges.push(description)
      }
    }
  })
  return structuralChanges
}

function buildEffectSummary(input: {
  cellDiffs: readonly WorkbookAgentPreviewCellDiff[]
  structuralChanges: readonly string[]
  truncatedCellDiffs: boolean
}): WorkbookAgentPreviewSummary['effectSummary'] {
  return {
    displayedCellDiffCount: input.cellDiffs.length,
    truncatedCellDiffs: input.truncatedCellDiffs,
    inputChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('input')).length,
    formulaChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('formula')).length,
    styleChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('style')).length,
    numberFormatChangeCount: input.cellDiffs.filter((diff) => diff.changeKinds.includes('numberFormat')).length,
    structuralChangeCount: input.structuralChanges.length,
  }
}

function buildTableSemanticTargets(input: {
  beforeEngine: SpreadsheetEngine
  previewEngine: SpreadsheetEngine
  bundle: WorkbookAgentCommandBundle
}): WorkbookAgentPreviewSemanticTarget[] {
  const tableNames = new Set<string>()
  input.bundle.commands.forEach((command) => {
    if (command.kind === 'upsertTable') {
      tableNames.add(command.table.name)
      return
    }
    if (command.kind === 'deleteTable') {
      tableNames.add(command.name)
    }
  })
  input.bundle.affectedRanges.forEach((range) => {
    input.previewEngine.getTables().forEach((table) => {
      if (rangesIntersect(tableRange(table), range)) {
        tableNames.add(table.name)
      }
    })
    input.beforeEngine.getTables().forEach((table) => {
      if (rangesIntersect(tableRange(table), range)) {
        tableNames.add(table.name)
      }
    })
  })
  return [...tableNames].flatMap((tableName) => {
    const table = input.previewEngine.getTable(tableName) ?? input.beforeEngine.getTable(tableName)
    return table ? tableSemanticTargets(table) : []
  })
}

function tableSemanticTargets(table: WorkbookTableSnapshot): WorkbookAgentPreviewSemanticTarget[] {
  const range = tableRange(table)
  const start = parseCellAddress(table.startAddress, table.sheetName)
  const end = parseCellAddress(table.endAddress, table.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  const headerRange = table.headerRow
    ? previewRange(table.sheetName, formatAddress(startRow, startCol), formatAddress(startRow, endCol), 'target')
    : undefined
  const dataStartRow = startRow + (table.headerRow ? 1 : 0)
  const dataEndRow = endRow - (table.totalsRow ? 1 : 0)
  const dataRange =
    dataStartRow <= dataEndRow
      ? previewRange(table.sheetName, formatAddress(dataStartRow, startCol), formatAddress(dataEndRow, endCol), 'target')
      : undefined
  const totalsRange = table.totalsRow
    ? previewRange(table.sheetName, formatAddress(endRow, startCol), formatAddress(endRow, endCol), 'target')
    : undefined
  return [
    {
      kind: 'table',
      tableName: table.name,
      label: `Table ${table.name}`,
      range,
    },
    ...(headerRange
      ? [{ kind: 'tableHeaderRow' as const, tableName: table.name, label: `${table.name} header row`, range: headerRange }]
      : []),
    ...(dataRange ? [{ kind: 'tableDataBody' as const, tableName: table.name, label: `${table.name} data body`, range: dataRange }] : []),
    ...(totalsRange
      ? [{ kind: 'tableTotalsRow' as const, tableName: table.name, label: `${table.name} totals row`, range: totalsRange }]
      : []),
    ...table.columnNames.map((columnName, columnIndex) => ({
      kind: 'tableColumn' as const,
      tableName: table.name,
      label: `${table.name}[${columnName}]`,
      columnName,
      columnIndex,
      range: previewRange(
        table.sheetName,
        formatAddress(startRow, startCol + columnIndex),
        formatAddress(endRow, startCol + columnIndex),
        'target',
      ),
    })),
  ]
}

function previewRange(sheetName: string, startAddress: string, endAddress: string, role: 'target' | 'source') {
  return { sheetName, startAddress, endAddress, role }
}

function tableRange(table: WorkbookTableSnapshot): CellRangeRef & { role: 'target' } {
  return {
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
    role: 'target',
  }
}

function rangesIntersect(
  left: Pick<CellRangeRef, 'sheetName' | 'startAddress' | 'endAddress'>,
  right: Pick<CellRangeRef, 'sheetName' | 'startAddress' | 'endAddress'>,
): boolean {
  if (left.sheetName !== right.sheetName) {
    return false
  }
  const leftStart = parseCellAddress(left.startAddress, left.sheetName)
  const leftEnd = parseCellAddress(left.endAddress, left.sheetName)
  const rightStart = parseCellAddress(right.startAddress, right.sheetName)
  const rightEnd = parseCellAddress(right.endAddress, right.sheetName)
  const leftTop = Math.min(leftStart.row, leftEnd.row)
  const leftBottom = Math.max(leftStart.row, leftEnd.row)
  const leftStartCol = Math.min(leftStart.col, leftEnd.col)
  const leftEndCol = Math.max(leftStart.col, leftEnd.col)
  const rightTop = Math.min(rightStart.row, rightEnd.row)
  const rightBottom = Math.max(rightStart.row, rightEnd.row)
  const rightStartCol = Math.min(rightStart.col, rightEnd.col)
  const rightEndCol = Math.max(rightStart.col, rightEnd.col)
  return !(leftBottom < rightTop || rightBottom < leftTop || leftEndCol < rightStartCol || rightEndCol < leftStartCol)
}

export async function buildWorkbookAgentPreview(input: {
  snapshot: WorkbookSnapshot
  replicaId: string
  bundle: WorkbookAgentCommandBundle
}): Promise<WorkbookAgentPreviewSummary> {
  if (!isWorkbookAgentCommandBundle(input.bundle)) {
    throw new Error('Invalid workbook agent command bundle')
  }
  toWorkbookCommandBundle(input.bundle)
  const previewEngine = new SpreadsheetEngine({
    workbookName: input.snapshot.workbook.name,
    replicaId: `${input.replicaId}:agent-preview`,
  })
  await previewEngine.ready()
  previewEngine.importSnapshot(input.snapshot)
  applyWorkbookAgentCommandBundle(previewEngine, input.bundle)
  const beforeEngine = new SpreadsheetEngine({
    workbookName: input.snapshot.workbook.name,
    replicaId: `${input.replicaId}:agent-preview-base`,
  })
  await beforeEngine.ready()
  beforeEngine.importSnapshot(input.snapshot)
  const targetAddresses = collectTargetAddresses(input.bundle)
  const resolvedCellDiffs = targetAddresses
    .flatMap(({ sheetName, address }) => {
      const beforeCell = beforeEngine.getCell(sheetName, address)
      const afterCell = previewEngine.getCell(sheetName, address)
      const diff = cloneCellDiff(beforeCell, afterCell)
      return diff ? [diff] : []
    })
    .slice(0, MAX_PREVIEW_DIFFS)
  const structuralChanges = buildStructuralChanges(input.bundle)
  const truncatedCellDiffs = input.bundle.affectedRanges.some((range) => {
    if (range.role !== 'target') {
      return false
    }
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    const rowCount = Math.abs(end.row - start.row) + 1
    const colCount = Math.abs(end.col - start.col) + 1
    return rowCount * colCount > MAX_PREVIEW_DIFFS
  })
  return {
    ranges: input.bundle.affectedRanges.map((range) => ({ ...range })),
    structuralChanges,
    cellDiffs: resolvedCellDiffs,
    semanticTargets: buildTableSemanticTargets({
      beforeEngine,
      previewEngine,
      bundle: input.bundle,
    }),
    effectSummary: buildEffectSummary({
      cellDiffs: resolvedCellDiffs,
      structuralChanges,
      truncatedCellDiffs,
    }),
  }
}
