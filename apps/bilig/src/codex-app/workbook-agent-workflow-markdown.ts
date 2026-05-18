import { formatErrorCode, ValueTag } from '@bilig/protocol'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookSearchReport, summarizeWorkbookStructure, traceWorkbookDependencies } from './workbook-agent-comprehension.js'

export function summarizeWorkbookMarkdown(summary: ReturnType<typeof summarizeWorkbookStructure>): string {
  const lines = [
    '## Workbook Summary',
    '',
    `Sheets: ${String(summary.summary.sheetCount)}`,
    `Formula cells: ${String(summary.summary.totalFormulaCellCount)}`,
    `Tables: ${String(summary.summary.tableCount)}`,
    `Pivots: ${String(summary.summary.pivotCount)}`,
    `Spills: ${String(summary.summary.spillCount)}`,
    '',
    '### Sheets',
  ]
  summary.sheets.forEach((sheet) => {
    lines.push(
      `- ${sheet.name}: ${String(sheet.cellCount)} populated cells, ${String(sheet.formulaCellCount)} formulas${sheet.usedRange ? `, used range ${sheet.usedRange.startAddress}:${sheet.usedRange.endAddress}` : ''}`,
    )
  })
  return lines.join('\n')
}

export function summarizeCurrentSheetMarkdown(sheet: ReturnType<typeof summarizeWorkbookStructure>['sheets'][number]): string {
  const lines = [
    '## Current Sheet Summary',
    '',
    `Sheet: ${sheet.name}`,
    `Order: ${String(sheet.order)}`,
    `Used range: ${sheet.usedRange ? `${sheet.usedRange.startAddress}:${sheet.usedRange.endAddress}` : '(empty)'}`,
    `Populated cells: ${String(sheet.cellCount)}`,
    `Formula cells: ${String(sheet.formulaCellCount)}`,
    `Tables: ${String(sheet.tableCount)}`,
    `Pivots: ${String(sheet.pivotCount)}`,
    `Spills: ${String(sheet.spillCount)}`,
    `Filters: ${String(sheet.filterCount)}`,
    `Sorts: ${String(sheet.sortCount)}`,
    `Freeze panes: ${sheet.freezePane ? `${String(sheet.freezePane.rows)} row(s), ${String(sheet.freezePane.cols)} column(s)` : 'none'}`,
    `Hidden row indexes: ${String(sheet.rowMetadata.hiddenIndexCount)}`,
    `Hidden column indexes: ${String(sheet.columnMetadata.hiddenIndexCount)}`,
    `Explicit row sizes: ${String(sheet.rowMetadata.explicitSizeIndexCount)}`,
    `Explicit column sizes: ${String(sheet.columnMetadata.explicitSizeIndexCount)}`,
    '',
    '### Tables',
  ]
  if (sheet.tables.length === 0) {
    lines.push('- None')
  } else {
    for (const table of sheet.tables) {
      lines.push(`- ${table.name}: ${table.startAddress}:${table.endAddress} (${String(table.columnCount)} columns)`)
    }
  }
  lines.push('', '### Pivots')
  if (sheet.pivots.length === 0) {
    lines.push('- None')
  } else {
    for (const pivot of sheet.pivots) {
      const source = pivot.source ?? 'cache-only source'
      lines.push(`- ${pivot.name}: ${pivot.address} from ${source} (${String(pivot.valueCount)} values)`)
    }
  }
  lines.push('', '### Spill Ranges')
  if (sheet.spills.length === 0) {
    lines.push('- None')
  } else {
    for (const spill of sheet.spills) {
      lines.push(`- ${spill.address}: ${String(spill.rows)}x${String(spill.cols)}`)
    }
  }
  return lines.join('\n')
}

export function summarizeRecentChangesMarkdown(changes: Awaited<ReturnType<ZeroSyncService['listWorkbookChanges']>>): string {
  const lines = ['## Recent Changes', '']
  if (changes.length === 0) {
    lines.push('No durable workbook changes are available yet.')
    return lines.join('\n')
  }
  changes.forEach((record) => {
    const location =
      record.range?.sheetName && record.range?.startAddress && record.range?.endAddress
        ? ` ${record.range.sheetName}!${record.range.startAddress}:${record.range.endAddress}`
        : record.sheetName && record.anchorAddress
          ? ` ${record.sheetName}!${record.anchorAddress}`
          : ''
    lines.push(`- r${String(record.revision)}: ${record.summary}${location}`)
  })
  return lines.join('\n')
}

export function summarizeDependencyTraceMarkdown(report: ReturnType<typeof traceWorkbookDependencies>): string {
  const lines = [
    '## Dependency Trace',
    '',
    `Root: ${report.root.sheetName}!${report.root.address}`,
    `Direction: ${report.direction}`,
    `Depth: ${String(report.depth)}`,
    `Direct precedents discovered: ${String(report.summary.precedentCount)}`,
    `Direct dependents discovered: ${String(report.summary.dependentCount)}`,
  ]
  if (report.summary.truncated) {
    lines.push('Trace output was truncated to stay inside the workflow node budget.')
  }
  lines.push('')
  if (report.layers.length === 0) {
    lines.push('No workbook precedents or dependents were found from the current selection.')
    return lines.join('\n')
  }
  report.layers.forEach((layer) => {
    lines.push(`### Depth ${String(layer.depth)}`)
    if (layer.precedents.length > 0) {
      lines.push('Precedents:')
      layer.precedents.forEach((node) => {
        lines.push(`- ${node.sheetName}!${node.address}: ${node.valueText}`)
      })
    }
    if (layer.dependents.length > 0) {
      lines.push('Dependents:')
      layer.dependents.forEach((node) => {
        lines.push(`- ${node.sheetName}!${node.address}: ${node.valueText}`)
      })
    }
    lines.push('')
  })
  return lines.join('\n').trimEnd()
}

export function serializeWorkflowCellValue(value: { tag: ValueTag; value?: number | boolean | string; code?: number }): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return '(empty)'
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return String(value.value ?? '')
    case ValueTag.Error:
      return typeof value.code === 'number' ? formatErrorCode(value.code) : '#ERROR!'
    default:
      return '(empty)'
  }
}

export function summarizeCellExplanationMarkdown(explanation: {
  readonly sheetName: string
  readonly address: string
  readonly valueText: string
  readonly formula: string | null
  readonly format: string | null
  readonly version: number
  readonly inCycle: boolean
  readonly mode: string | null
  readonly topoRank: number | null
  readonly directPrecedents: readonly string[]
  readonly directDependents: readonly string[]
}): string {
  const lines = [
    '## Current Cell',
    '',
    `Cell: ${explanation.sheetName}!${explanation.address}`,
    `Value: ${explanation.valueText}`,
    `Formula: ${explanation.formula ?? '(none)'}`,
    `Calculation mode: ${explanation.mode ?? '(unknown)'}`,
    `Version: ${String(explanation.version)}`,
    `In cycle: ${explanation.inCycle ? 'yes' : 'no'}`,
    `Direct precedents: ${String(explanation.directPrecedents.length)}`,
    `Direct dependents: ${String(explanation.directDependents.length)}`,
  ]
  if (explanation.topoRank !== null) {
    lines.push(`Topological rank: ${String(explanation.topoRank)}`)
  }
  if (explanation.format) {
    lines.push(`Number format: ${explanation.format}`)
  }
  lines.push('', '### Direct precedents')
  if (explanation.directPrecedents.length === 0) {
    lines.push('- None')
  } else {
    for (const precedent of explanation.directPrecedents) {
      lines.push(`- ${precedent}`)
    }
  }
  lines.push('', '### Direct dependents')
  if (explanation.directDependents.length === 0) {
    lines.push('- None')
  } else {
    for (const dependent of explanation.directDependents) {
      lines.push(`- ${dependent}`)
    }
  }
  return lines.join('\n')
}

export function summarizeSearchResultsMarkdown(report: WorkbookSearchReport): string {
  const lines = ['## Workbook Search', '', `Query: ${report.query}`, `Matches: ${String(report.summary.matchCount)}`]
  if (report.summary.truncated) {
    lines.push('Results were truncated to stay inside the workflow result budget.')
  }
  lines.push('')
  if (report.matches.length === 0) {
    lines.push('No workbook matches were found for the requested query.')
    return lines.join('\n')
  }
  lines.push('### Top matches')
  for (const match of report.matches) {
    if (match.kind === 'sheet') {
      lines.push(`- Sheet ${match.sheetName} [${match.reasons.join(', ')}]`)
      continue
    }
    const location = `${match.sheetName}!${match.address ?? '?'}`
    const snippet = match.formula ?? match.valueText ?? match.inputText ?? match.snippet ?? '(no snippet)'
    lines.push(`- ${location}: ${snippet} [${match.reasons.join(', ')}]`)
  }
  return lines.join('\n')
}
