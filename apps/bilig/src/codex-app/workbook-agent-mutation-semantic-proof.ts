import type { WorkbookAgentCommand, WorkbookAgentPreviewCellDiff } from '@bilig/agent-api'
import { diffWorkbookSemanticSnapshots } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  buildCellNumberFormatCode,
  sanitizeCellStyleRecord,
  type CellStylePatch,
  type LiteralInput,
  type WorkbookSnapshot,
} from '@bilig/protocol'
import type { WorkbookVerificationMismatch } from './workbook-agent-rendered-readback.js'

type WorkbookSemanticStyleRecord = NonNullable<NonNullable<WorkbookSnapshot['workbook']['metadata']>['styles']>[number]

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asLiteralInput(value: unknown): LiteralInput | undefined {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : undefined
}

function semanticStyleFromRecord(style: Record<string, unknown>, id: string): WorkbookSemanticStyleRecord {
  return sanitizeCellStyleRecord(id, style) ?? { id }
}

function readbackSemanticSnapshot(readbacks: readonly unknown[]): WorkbookSnapshot {
  const sheets = new Map<string, WorkbookSnapshot['sheets'][number]>()
  const styles = new Map<string, WorkbookSemanticStyleRecord>()
  readbacks.forEach((readback) => {
    if (!isRecord(readback) || !isRecord(readback['range']) || !Array.isArray(readback['rows'])) {
      return
    }
    const sheetName = typeof readback['range']['sheetName'] === 'string' ? readback['range']['sheetName'] : null
    if (!sheetName) {
      return
    }
    const sheet =
      sheets.get(sheetName) ??
      ({
        name: sheetName,
        order: sheets.size,
        metadata: {
          styleRanges: [],
        },
        cells: [],
      } satisfies WorkbookSnapshot['sheets'][number])
    readback['rows'].forEach((row) => {
      if (!Array.isArray(row)) {
        return
      }
      row.forEach((cell) => {
        if (!isRecord(cell) || typeof cell['address'] !== 'string') {
          return
        }
        const snapshotCell: WorkbookSnapshot['sheets'][number]['cells'][number] = {
          address: cell['address'],
        }
        const value = asLiteralInput(cell['value'] ?? null)
        if (value !== undefined) {
          snapshotCell.value = value
        }
        if (typeof cell['formula'] === 'string' && cell['formula'].length > 0) {
          snapshotCell.formula = cell['formula']
        }
        if (typeof cell['displayFormat'] === 'string' && cell['displayFormat'].length > 0) {
          snapshotCell.format = cell['displayFormat']
        }
        sheet.cells.push(snapshotCell)
        if (typeof cell['styleId'] === 'string') {
          sheet.metadata ??= {}
          sheet.metadata.styleRanges ??= []
          sheet.metadata.styleRanges.push({
            range: {
              sheetName,
              startAddress: cell['address'],
              endAddress: cell['address'],
            },
            styleId: cell['styleId'],
          })
        }
      })
    })
    if (Array.isArray(readback['styles'])) {
      readback['styles'].forEach((style) => {
        if (isRecord(style) && typeof style['id'] === 'string') {
          styles.set(style['id'], semanticStyleFromRecord(style, style['id']))
        }
      })
    }
    sheets.set(sheetName, sheet)
  })
  const semanticSheets = [...sheets.values()]
  semanticSheets.forEach((sheet, order) => {
    sheet.order = order
  })
  return {
    version: 1,
    workbook: {
      name: 'authoritative-readback',
      metadata: {
        styles: [...styles.values()],
      },
    },
    sheets: semanticSheets,
  }
}

function mergeStylePatch(style: WorkbookSemanticStyleRecord, patch: CellStylePatch): WorkbookSemanticStyleRecord {
  const merged = {
    ...style,
    ...(patch.fill ? { fill: { ...style.fill, ...patch.fill } } : {}),
    ...(patch.font ? { font: { ...style.font, ...patch.font } } : {}),
    ...(patch.alignment ? { alignment: { ...style.alignment, ...patch.alignment } } : {}),
    ...(patch.borders ? { borders: { ...style.borders, ...patch.borders } } : {}),
  }
  return sanitizeCellStyleRecord(style.id, merged) ?? { id: style.id }
}

function findSemanticSheet(snapshot: WorkbookSnapshot, sheetName: string): WorkbookSnapshot['sheets'][number] | null {
  return snapshot.sheets.find((sheet) => sheet.name === sheetName) ?? null
}

function findSemanticCell(
  snapshot: WorkbookSnapshot,
  sheetName: string,
  address: string,
): WorkbookSnapshot['sheets'][number]['cells'][number] | null {
  return findSemanticSheet(snapshot, sheetName)?.cells.find((cell) => cell.address === address) ?? null
}

function replaceCellStyleRange(input: {
  readonly snapshot: WorkbookSnapshot
  readonly sheetName: string
  readonly address: string
  readonly styleId: string
}): void {
  const sheet = findSemanticSheet(input.snapshot, input.sheetName)
  if (!sheet) {
    return
  }
  sheet.metadata ??= {}
  sheet.metadata.styleRanges = (sheet.metadata.styleRanges ?? []).filter(
    (styleRange) =>
      !(
        styleRange.range.sheetName === input.sheetName &&
        styleRange.range.startAddress === input.address &&
        styleRange.range.endAddress === input.address
      ),
  )
  sheet.metadata.styleRanges.push({
    range: {
      sheetName: input.sheetName,
      startAddress: input.address,
      endAddress: input.address,
    },
    styleId: input.styleId,
  })
}

function expectedSemanticSnapshot(input: {
  readonly actualSnapshot: WorkbookSnapshot
  readonly previewDiffs: readonly WorkbookAgentPreviewCellDiff[]
  readonly commands: readonly WorkbookAgentCommand[]
}): WorkbookSnapshot {
  const expected = structuredClone(input.actualSnapshot)
  input.previewDiffs.forEach((diff) => {
    const cell = findSemanticCell(expected, diff.sheetName, diff.address)
    if (!cell) {
      return
    }
    if (diff.changeKinds.includes('input')) {
      cell.value = diff.afterInput
      delete cell.formula
    }
    if (diff.changeKinds.includes('formula') && diff.afterFormula !== null) {
      cell.formula = diff.afterFormula
    }
  })
  const styles = new Map((expected.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
  let expectedStyleIndex = 0
  input.commands.forEach((command) => {
    if (command.kind !== 'formatRange') {
      return
    }
    const start = parseCellAddress(command.range.startAddress, command.range.sheetName)
    const end = parseCellAddress(command.range.endAddress, command.range.sheetName)
    const rowStart = Math.min(start.row, end.row)
    const rowEnd = Math.max(start.row, end.row)
    const colStart = Math.min(start.col, end.col)
    const colEnd = Math.max(start.col, end.col)
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        const address = formatAddress(row, col)
        const cell = findSemanticCell(expected, command.range.sheetName, address)
        if (!cell) {
          continue
        }
        if (command.numberFormat !== undefined) {
          cell.format = buildCellNumberFormatCode(command.numberFormat)
        }
        if (command.patch !== undefined) {
          const currentStyleRange = findSemanticSheet(expected, command.range.sheetName)?.metadata?.styleRanges?.find(
            (styleRange) =>
              styleRange.range.sheetName === command.range.sheetName &&
              styleRange.range.startAddress === address &&
              styleRange.range.endAddress === address,
          )
          const currentStyle = currentStyleRange ? (styles.get(currentStyleRange.styleId) ?? { id: currentStyleRange.styleId }) : { id: '' }
          const styleId = `expected-agent-style-${String(expectedStyleIndex)}`
          expectedStyleIndex += 1
          styles.set(styleId, mergeStylePatch({ ...currentStyle, id: styleId }, command.patch))
          replaceCellStyleRange({
            snapshot: expected,
            sheetName: command.range.sheetName,
            address,
            styleId,
          })
        }
      }
    }
  })
  expected.workbook.metadata ??= {}
  expected.workbook.metadata.styles = [...styles.values()]
  return expected
}

export function collectSemanticReadbackMismatches(input: {
  readonly commands: readonly WorkbookAgentCommand[]
  readonly previewDiffs: readonly WorkbookAgentPreviewCellDiff[]
  readonly readbacks: readonly unknown[]
}): {
  readonly matched: boolean | null
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
} {
  if (input.previewDiffs.length === 0 && input.commands.every((command) => command.kind !== 'formatRange')) {
    return { matched: null, mismatches: [], incompleteReason: null }
  }
  const actualSnapshot = readbackSemanticSnapshot(input.readbacks)
  const expectedSnapshot = expectedSemanticSnapshot({
    actualSnapshot,
    previewDiffs: input.previewDiffs,
    commands: input.commands,
  })
  const diffs = diffWorkbookSemanticSnapshots(expectedSnapshot, actualSnapshot)
  return {
    matched: diffs.length === 0,
    mismatches: diffs.map((diff) => ({
      sheetName: input.previewDiffs[0]?.sheetName ?? input.commands[0]?.kind ?? 'workbook',
      address: input.previewDiffs[0]?.address ?? 'semantic',
      field: diff.path,
      expected: diff.left,
      actual: diff.right,
      source: 'authoritative',
    })),
    incompleteReason: diffs.length === 0 ? null : 'Authoritative semantic readback did not match expected workbook state.',
  }
}
