import { formatAddress, parseCellAddress, parseFormula, serializeFormula } from '@bilig/formula'
import { buildCellNumberFormatCode, type CellRangeRef, type CellStylePatch, type CellStyleRecord } from '@bilig/protocol'
import type { WorkbookRef, WorkbookRefData } from './find.js'
import { materializeFormulaLabels, type WorkbookFormulaLabelReplacement } from './formula-usage.js'
import type { WorkbookActionCommand } from './model.js'
import type { WorkbookCellNumberFormatOp, WorkbookCellStyleOp, WorkbookOp } from './ops.js'
import type { WorkbookCommandResolvedRefs, WorkbookResolvedRefValue } from './result.js'

export type WorkbookConcreteCommandOp = Extract<WorkbookOp, { kind: 'setCellFormula' | 'setCellValue' | 'setCellFormat' | 'clearCell' }>

interface WorkbookConcreteCell {
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
}

export interface WorkbookCommandOpsExpectation {
  readonly ops: readonly WorkbookOp[]
}

function concreteSingleCell(target: WorkbookRef): { sheetName: string; address: string } | null {
  if (target.kind !== 'range') {
    return null
  }
  const range = target.range
  return range.startAddress === range.endAddress ? { sheetName: range.sheetName, address: range.startAddress } : null
}

function isRangeRefData(value: WorkbookRefData): value is Extract<WorkbookRefData, { readonly kind: 'range' }> {
  return value.kind === 'range'
}

function concreteRangesFromResolvedRef(value: WorkbookResolvedRefValue | undefined): readonly CellRangeRef[] | null {
  if (value === undefined) {
    return null
  }
  const values = Array.isArray(value) ? value : [value]
  if (!values.every(isRangeRefData)) {
    return null
  }
  return values.map((entry) => entry.range)
}

function cellsFromRange(range: CellRangeRef): readonly WorkbookConcreteCell[] {
  const start = parseCellAddress(range.startAddress)
  const end = parseCellAddress(range.endAddress)
  if (end.row < start.row || end.col < start.col) {
    return []
  }
  const cells: WorkbookConcreteCell[] = []
  for (let row = start.row; row <= end.row; row += 1) {
    for (let col = start.col; col <= end.col; col += 1) {
      cells.push({
        sheetName: range.sheetName,
        address: formatAddress(row, col),
        row,
        col,
      })
    }
  }
  return cells
}

function concreteCellsFromResolvedRefs(resolvedRefs: WorkbookCommandResolvedRefs | undefined): readonly WorkbookConcreteCell[] | null {
  const ranges = concreteRangesFromResolvedRef(resolvedRefs?.target)
  if (ranges === null) {
    return null
  }
  return ranges.flatMap(cellsFromRange)
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function cellSource(cell: WorkbookConcreteCell): string {
  return `${quoteSheetName(cell.sheetName)}!${cell.address}`
}

function rangeSource(range: CellRangeRef): string {
  const sheet = quoteSheetName(range.sheetName)
  return range.startAddress === range.endAddress ? `${sheet}!${range.startAddress}` : `${sheet}!${range.startAddress}:${range.endAddress}`
}

function cellKey(cell: Pick<WorkbookConcreteCell, 'sheetName' | 'address'>): string {
  return `${cell.sheetName}!${cell.address}`
}

function cellSet(cells: readonly WorkbookConcreteCell[]): ReadonlySet<string> {
  return new Set(cells.map(cellKey))
}

function coverRangeCells(range: CellRangeRef, targetCells: ReadonlySet<string>, coveredCells: Set<string>): boolean {
  const cells = cellsFromRange(range)
  if (cells.length === 0) {
    return false
  }
  for (const cell of cells) {
    if (!targetCells.has(cellKey(cell))) {
      return false
    }
  }
  for (const cell of cells) {
    coveredCells.add(cellKey(cell))
  }
  return true
}

function coversEveryTargetCell(targetCells: ReadonlySet<string>, coveredCells: ReadonlySet<string>): boolean {
  for (const targetCell of targetCells) {
    if (!coveredCells.has(targetCell)) {
      return false
    }
  }
  return true
}

interface FormatSupportCatalog {
  readonly styles: ReadonlyMap<string, WorkbookCellStyleOp>
  readonly formats: ReadonlyMap<string, WorkbookCellNumberFormatOp>
}

function formatSupportCatalog(ops: readonly WorkbookOp[]): FormatSupportCatalog {
  const styles = new Map<string, WorkbookCellStyleOp>()
  const formats = new Map<string, WorkbookCellNumberFormatOp>()
  for (const op of ops) {
    if (op.kind === 'upsertCellStyle') {
      styles.set(op.style.id, op.style)
    } else if (op.kind === 'upsertCellNumberFormat') {
      formats.set(op.format.id, op.format)
    }
  }
  return { styles, formats }
}

function normalizedNumberFormatCode(value: string): string {
  return buildCellNumberFormatCode(value)
}

function numberFormatRecordMatches(format: WorkbookCellNumberFormatOp | undefined, expected: string): boolean {
  return format !== undefined && normalizedNumberFormatCode(format.code) === normalizedNumberFormatCode(expected)
}

function styleFieldMatches(actual: unknown, expected: unknown): boolean {
  if (expected === undefined) {
    return true
  }
  return expected === null ? actual === undefined : actual === expected
}

function borderPatchSideMatchesRecord(
  patch: NonNullable<NonNullable<CellStylePatch['borders']>['top']> | null | undefined,
  record: NonNullable<NonNullable<CellStyleRecord['borders']>['top']> | undefined,
): boolean {
  if (patch === undefined) {
    return true
  }
  if (patch === null || patch.style === undefined || patch.weight === undefined || patch.color === undefined) {
    return record === undefined
  }
  return record?.style === patch.style && record.weight === patch.weight && record.color === patch.color
}

function stylePatchMatchesRecord(patch: CellStylePatch, record: WorkbookCellStyleOp): boolean {
  return (
    styleFieldMatches(record.fill?.backgroundColor, patch.fill?.backgroundColor) &&
    styleFieldMatches(record.font?.family, patch.font?.family) &&
    styleFieldMatches(record.font?.size, patch.font?.size) &&
    styleFieldMatches(record.font?.bold, patch.font?.bold) &&
    styleFieldMatches(record.font?.italic, patch.font?.italic) &&
    styleFieldMatches(record.font?.underline, patch.font?.underline) &&
    styleFieldMatches(record.font?.color, patch.font?.color) &&
    styleFieldMatches(record.alignment?.horizontal, patch.alignment?.horizontal) &&
    styleFieldMatches(record.alignment?.vertical, patch.alignment?.vertical) &&
    styleFieldMatches(record.alignment?.wrap, patch.alignment?.wrap) &&
    styleFieldMatches(record.alignment?.indent, patch.alignment?.indent) &&
    styleFieldMatches(record.alignment?.shrinkToFit, patch.alignment?.shrinkToFit) &&
    styleFieldMatches(record.alignment?.readingOrder, patch.alignment?.readingOrder) &&
    styleFieldMatches(record.alignment?.textRotation, patch.alignment?.textRotation) &&
    styleFieldMatches(record.alignment?.justifyLastLine, patch.alignment?.justifyLastLine) &&
    borderPatchSideMatchesRecord(patch.borders?.top, record.borders?.top) &&
    borderPatchSideMatchesRecord(patch.borders?.right, record.borders?.right) &&
    borderPatchSideMatchesRecord(patch.borders?.bottom, record.borders?.bottom) &&
    borderPatchSideMatchesRecord(patch.borders?.left, record.borders?.left)
  )
}

function borderPatchSideNeedsRecord(patch: NonNullable<NonNullable<CellStylePatch['borders']>['top']> | null | undefined): boolean {
  return patch !== undefined && patch !== null && patch.style !== undefined && patch.weight !== undefined && patch.color !== undefined
}

function stylePatchNeedsRecord(patch: CellStylePatch): boolean {
  return (
    (patch.fill?.backgroundColor !== undefined && patch.fill.backgroundColor !== null) ||
    (patch.font?.family !== undefined && patch.font.family !== null) ||
    (patch.font?.size !== undefined && patch.font.size !== null) ||
    (patch.font?.bold !== undefined && patch.font.bold !== null) ||
    (patch.font?.italic !== undefined && patch.font.italic !== null) ||
    (patch.font?.underline !== undefined && patch.font.underline !== null) ||
    (patch.font?.color !== undefined && patch.font.color !== null) ||
    (patch.alignment?.horizontal !== undefined && patch.alignment.horizontal !== null) ||
    (patch.alignment?.vertical !== undefined && patch.alignment.vertical !== null) ||
    (patch.alignment?.wrap !== undefined && patch.alignment.wrap !== null) ||
    (patch.alignment?.indent !== undefined && patch.alignment.indent !== null) ||
    (patch.alignment?.shrinkToFit !== undefined && patch.alignment.shrinkToFit !== null) ||
    (patch.alignment?.readingOrder !== undefined && patch.alignment.readingOrder !== null) ||
    (patch.alignment?.textRotation !== undefined && patch.alignment.textRotation !== null) ||
    (patch.alignment?.justifyLastLine !== undefined && patch.alignment.justifyLastLine !== null) ||
    borderPatchSideNeedsRecord(patch.borders?.top) ||
    borderPatchSideNeedsRecord(patch.borders?.right) ||
    borderPatchSideNeedsRecord(patch.borders?.bottom) ||
    borderPatchSideNeedsRecord(patch.borders?.left)
  )
}

function styleRangeMatchesExpected(
  command: Extract<WorkbookActionCommand, { readonly kind: 'format' }>,
  op: Extract<WorkbookOp, { readonly kind: 'setStyleRange' }>,
  catalog: FormatSupportCatalog,
  usedStyleIds: Set<string>,
): boolean {
  if (command.style === undefined) {
    return false
  }
  const style = catalog.styles.get(op.styleId)
  if (style === undefined) {
    return !stylePatchNeedsRecord(command.style)
  }
  if (!stylePatchMatchesRecord(command.style, style)) {
    return false
  }
  usedStyleIds.add(op.styleId)
  return true
}

function formatRangeMatchesExpected(
  command: Extract<WorkbookActionCommand, { readonly kind: 'format' }>,
  op: Extract<WorkbookOp, { readonly kind: 'setFormatRange' }>,
  catalog: FormatSupportCatalog,
  usedFormatIds: Set<string>,
): boolean {
  if (command.numberFormat === undefined) {
    return false
  }
  if (command.numberFormat === null) {
    return false
  }
  if (!numberFormatRecordMatches(catalog.formats.get(op.formatId), command.numberFormat)) {
    return false
  }
  usedFormatIds.add(op.formatId)
  return true
}

function everySupportOpWasUsed(
  catalog: FormatSupportCatalog,
  usedStyleIds: ReadonlySet<string>,
  usedFormatIds: ReadonlySet<string>,
): boolean {
  for (const id of catalog.styles.keys()) {
    if (!usedStyleIds.has(id)) {
      return false
    }
  }
  for (const id of catalog.formats.keys()) {
    if (!usedFormatIds.has(id)) {
      return false
    }
  }
  return true
}

function formatOpBindsToTarget(
  command: Extract<WorkbookActionCommand, { readonly kind: 'format' }>,
  op: WorkbookOp,
  targetCells: ReadonlySet<string>,
  catalog: FormatSupportCatalog,
  usedStyleIds: Set<string>,
  usedFormatIds: Set<string>,
  coveredStyleCells: Set<string>,
  coveredNumberFormatCells: Set<string>,
): {
  readonly binding: 'target' | 'support'
} | null {
  if (op.kind === 'upsertCellStyle' || op.kind === 'upsertCellNumberFormat') {
    return { binding: 'support' }
  }
  if (op.kind === 'setCellFormat') {
    const key = `${op.sheetName}!${op.address}`
    if (command.numberFormat === undefined || !targetCells.has(key) || op.format !== command.numberFormat) {
      return null
    }
    coveredNumberFormatCells.add(key)
    return { binding: 'target' }
  }
  if (op.kind === 'setFormatRange') {
    if (
      !formatRangeMatchesExpected(command, op, catalog, usedFormatIds) ||
      !coverRangeCells(op.range, targetCells, coveredNumberFormatCells)
    ) {
      return null
    }
    return { binding: 'target' }
  }
  if (op.kind === 'setStyleRange') {
    if (!styleRangeMatchesExpected(command, op, catalog, usedStyleIds) || !coverRangeCells(op.range, targetCells, coveredStyleCells)) {
      return null
    }
    return { binding: 'target' }
  }
  return null
}

function formatOpsMatchExpected(
  command: Extract<WorkbookActionCommand, { readonly kind: 'format' }>,
  ops: readonly WorkbookOp[],
  resolvedRefs?: WorkbookCommandResolvedRefs,
): boolean {
  if (command.style === undefined && command.numberFormat === undefined) {
    return ops.length === 0
  }
  const targetCells = concreteCellsFromResolvedRefs(resolvedRefs)
  if (targetCells === null || targetCells.length === 0) {
    return false
  }
  const targetCellSet = cellSet(targetCells)
  const catalog = formatSupportCatalog(ops)
  const usedStyleIds = new Set<string>()
  const usedFormatIds = new Set<string>()
  const coveredStyleCells = new Set<string>()
  const coveredNumberFormatCells = new Set<string>()
  for (const op of ops) {
    const binding = formatOpBindsToTarget(
      command,
      op,
      targetCellSet,
      catalog,
      usedStyleIds,
      usedFormatIds,
      coveredStyleCells,
      coveredNumberFormatCells,
    )
    if (binding === null) {
      return false
    }
  }
  return (
    (command.style === undefined || coversEveryTargetCell(targetCellSet, coveredStyleCells)) &&
    (command.numberFormat === undefined || coversEveryTargetCell(targetCellSet, coveredNumberFormatCells)) &&
    everySupportOpWasUsed(catalog, usedStyleIds, usedFormatIds)
  )
}

function workbookRefKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
}

function resolvedInputForLabel(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  label: { readonly ref: WorkbookRef },
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
): WorkbookResolvedRefValue | null {
  const inputs = resolvedRefs?.inputs
  if (inputs === undefined) {
    return null
  }
  const labelKey = workbookRefKey(label.ref)
  const inputIndex = command.inputs.findIndex((input) => workbookRefKey(input) === labelKey)
  return inputIndex < 0 ? null : (inputs[inputIndex] ?? null)
}

function labelReplacementForCell(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  label: { readonly name: string; readonly ref: WorkbookRef },
  targetCells: readonly WorkbookConcreteCell[],
  cellIndex: number,
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
): WorkbookFormulaLabelReplacement | null {
  const value = resolvedInputForLabel(command, label, resolvedRefs)
  const inputRanges = concreteRangesFromResolvedRef(value ?? undefined)
  if (inputRanges === null) {
    return null
  }
  const inputCells = inputRanges.flatMap(cellsFromRange)
  if (inputCells.length === targetCells.length) {
    const inputCell = inputCells[cellIndex]
    return inputCell === undefined ? null : { name: label.name, source: cellSource(inputCell) }
  }
  if (inputCells.length === 1) {
    return { name: label.name, source: cellSource(inputCells[0]!) }
  }
  if (inputRanges.length === 1) {
    return { name: label.name, source: rangeSource(inputRanges[0]!) }
  }
  return null
}

function resolvedFormulaLabelsForCell(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  targetCells: readonly WorkbookConcreteCell[],
  cellIndex: number,
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
): readonly WorkbookFormulaLabelReplacement[] | null {
  const replacements: WorkbookFormulaLabelReplacement[] = []
  for (const label of command.labels) {
    const replacement = labelReplacementForCell(command, label, targetCells, cellIndex, resolvedRefs)
    if (replacement === null) {
      return null
    }
    replacements.push(replacement)
  }
  return replacements
}

function formulaForExpectedCell(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  targetCells: readonly WorkbookConcreteCell[],
  cellIndex: number,
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
  formulaLabels: readonly WorkbookFormulaLabelReplacement[] = [],
): string | null {
  if (command.labels.length === 0) {
    return command.inputs.length > 0 ? null : command.formula
  }

  const resolvedLabels = resolvedFormulaLabelsForCell(command, targetCells, cellIndex, resolvedRefs)
  if (resolvedLabels !== null) {
    return materializeFormulaLabels(command.formula, resolvedLabels)
  }
  if (targetCells.length === 1 && formulaLabels.length > 0) {
    return materializeFormulaLabels(command.formula, formulaLabels)
  }
  return null
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalValue)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalValue(entry)]),
    )
  }
  return value
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

export function workbookOpMatches(expected: WorkbookOp, actual: WorkbookOp): boolean {
  return canonicalJson(actual) === canonicalJson(expected)
}

export function workbookOpsMatch(left: readonly WorkbookOp[], right: readonly WorkbookOp[]): boolean {
  return (
    left.length === right.length &&
    left.every((op, index) => {
      const other = right[index]
      return other !== undefined && workbookOpMatches(op, other)
    })
  )
}

export function workbookConcreteOpMatches(expected: WorkbookConcreteCommandOp, actual: WorkbookOp): boolean {
  if (expected.kind !== actual.kind) {
    return false
  }
  switch (expected.kind) {
    case 'setCellFormula':
      return (
        actual.kind === 'setCellFormula' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.formula === expected.formula
      )
    case 'setCellValue':
      return (
        actual.kind === 'setCellValue' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.value === expected.value
      )
    case 'setCellFormat':
      return (
        actual.kind === 'setCellFormat' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.format === expected.format
      )
    case 'clearCell':
      return actual.kind === 'clearCell' && actual.sheetName === expected.sheetName && actual.address === expected.address
  }
}

export function expectedConcreteCommandOp(command: WorkbookActionCommand): WorkbookConcreteCommandOp | null {
  if (command.kind === 'op') {
    return null
  }

  const target = concreteSingleCell(command.target)
  if (target === null) {
    return null
  }
  switch (command.kind) {
    case 'writeFormula':
      return {
        kind: 'setCellFormula',
        sheetName: target.sheetName,
        address: target.address,
        formula: command.formula,
      }
    case 'writeValue':
      return {
        kind: 'setCellValue',
        sheetName: target.sheetName,
        address: target.address,
        value: command.value,
      }
    case 'clear':
      return {
        kind: 'clearCell',
        sheetName: target.sheetName,
        address: target.address,
      }
    case 'format':
      if (command.numberFormat === undefined) {
        return null
      }
      return {
        kind: 'setCellFormat',
        sheetName: target.sheetName,
        address: target.address,
        format: command.numberFormat,
      }
  }
}

export function expectedCommandOps(command: WorkbookActionCommand): WorkbookCommandOpsExpectation | null {
  if (command.kind === 'op') {
    return {
      ops: [command.op],
    }
  }

  const op = expectedConcreteCommandOp(command)
  return op === null
    ? null
    : {
        ops: [op],
      }
}

function formulasMatchExpected(expected: string, actual: string, formulaLabels: readonly WorkbookFormulaLabelReplacement[] = []): boolean {
  try {
    const expectedSource = formulaLabels.length > 0 ? materializeFormulaLabels(expected, formulaLabels) : expected
    return serializeFormula(parseFormula(expectedSource)) === serializeFormula(parseFormula(actual))
  } catch {
    return expected === actual
  }
}

export function commandOpsMatchExpected(
  command: WorkbookActionCommand,
  ops: readonly WorkbookOp[],
  formulaLabels: readonly WorkbookFormulaLabelReplacement[] = [],
  resolvedRefs?: WorkbookCommandResolvedRefs,
): boolean {
  if (ops.length === 0) {
    return true
  }
  if (command.kind === 'writeFormula') {
    const expected = expectedConcreteCommandOp(command)
    const concreteCells = expected === null ? concreteCellsFromResolvedRefs(resolvedRefs) : null
    if (expected === null && concreteCells === null) {
      return false
    }
    const expectedCells =
      concreteCells ??
      (expected === null
        ? []
        : [
            {
              sheetName: expected.sheetName,
              address: expected.address,
              ...parseCellAddress(expected.address),
            },
          ])
    return (
      expectedCells.length === ops.length &&
      expectedCells.every((cell, index) => {
        const actual = ops[index]
        const expectedFormula = formulaForExpectedCell(command, expectedCells, index, resolvedRefs, formulaLabels)
        return (
          actual !== undefined &&
          actual.kind === 'setCellFormula' &&
          actual.sheetName === cell.sheetName &&
          actual.address === cell.address &&
          expectedFormula !== null &&
          formulasMatchExpected(expectedFormula, actual.formula)
        )
      })
    )
  }
  const expected = expectedCommandOps(command)
  if (expected === null) {
    const concreteCells = concreteCellsFromResolvedRefs(resolvedRefs)
    if (concreteCells === null) {
      return false
    }
    if (command.kind === 'writeValue') {
      return (
        concreteCells.length === ops.length &&
        concreteCells.every((cell, index) => {
          const actual = ops[index]
          return (
            actual !== undefined &&
            actual.kind === 'setCellValue' &&
            actual.sheetName === cell.sheetName &&
            actual.address === cell.address &&
            actual.value === command.value
          )
        })
      )
    }
    if (command.kind === 'clear') {
      return (
        concreteCells.length === ops.length &&
        concreteCells.every((cell, index) => {
          const actual = ops[index]
          return (
            actual !== undefined && actual.kind === 'clearCell' && actual.sheetName === cell.sheetName && actual.address === cell.address
          )
        })
      )
    }
    if (command.kind === 'format') {
      return formatOpsMatchExpected(command, ops, resolvedRefs)
    }
    return false
  }
  return workbookOpsMatch(expected.ops, ops)
}
