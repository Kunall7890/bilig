import { formatAddress, parseCellAddress, type FormulaNode, type RangeRefNode } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import type { StringPool } from '../../string-pool.js'
import type { WorkbookStore } from '../../workbook-store.js'
import type { ParsedCompiledFormula } from './formula-binding-direct-descriptors.js'
import { type DynamicRangeBounds, normalizeRangeDependency, rangeBounds } from './formula-binding-dynamic-range-bounds.js'
import {
  branchChoice,
  criterionMatches,
  evaluateScalar,
  integerScalar,
  readCellValue,
  type DynamicFormulaAstResolver,
} from './formula-binding-dynamic-scalar.js'

const MAX_COMPACT_INDEX_DEPENDENCY_CELLS = 4096
const MAX_DYNAMIC_OFFSET_DEPENDENCY_CELLS = 4096

export interface DynamicIndexDependencyPlan {
  readonly compactedRangeDependencies: ReadonlySet<string>
  readonly selectedCells: readonly DynamicIndexSelectedCell[]
}

export interface DynamicIndexSelectedCell {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly address: string
}

function isCriteriaAggregateCallee(callee: string): boolean {
  return (
    callee === 'SUMIF' ||
    callee === 'AVERAGEIF' ||
    callee === 'SUMIFS' ||
    callee === 'AVERAGEIFS' ||
    callee === 'MINIFS' ||
    callee === 'MAXIFS'
  )
}

export function formulaMayNeedDynamicIndexDependencyPlan(compiled: ParsedCompiledFormula): boolean {
  let found = false
  const visit = (node: FormulaNode): void => {
    if (found) {
      return
    }
    switch (node.kind) {
      case 'CallExpr': {
        const callee = node.callee.trim().toUpperCase()
        if (callee === 'INDEX' || callee === 'OFFSET' || isCriteriaAggregateCallee(callee)) {
          found = true
          return
        }
        node.args.forEach(visit)
        return
      }
      case 'UnaryExpr':
        visit(node.argument)
        return
      case 'BinaryExpr':
        visit(node.left)
        visit(node.right)
        return
      case 'InvokeExpr':
        visit(node.callee)
        node.args.forEach(visit)
        return
      case 'ArrayConstant':
        node.rows.forEach((row) => row.forEach(visit))
        return
      case 'BooleanLiteral':
      case 'CellRef':
      case 'ColumnRef':
      case 'ErrorLiteral':
      case 'NameRef':
      case 'NumberLiteral':
      case 'OmittedArgument':
      case 'RangeRef':
      case 'RowRef':
      case 'SpillRef':
      case 'StringLiteral':
      case 'StructuredRef':
        return
    }
  }

  visit(compiled.optimizedAst)
  return found
}

function selectedIndexCells(args: {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly ownerSheetName: string
  readonly getFormulaAst?: DynamicFormulaAstResolver | undefined
  readonly range: RangeRefNode
  readonly rowArg: FormulaNode | undefined
  readonly colArg: FormulaNode | undefined
}): DynamicIndexSelectedCell[] {
  const bounds = rangeBounds(args.range, args.ownerSheetName)
  if (!bounds) {
    return []
  }

  const rowStart = bounds.rowStart
  const rowEnd = bounds.rowEnd
  const colStart = bounds.colStart
  const colEnd = bounds.colEnd
  const cols = colEnd - colStart + 1
  const rowNumber = integerScalar({ ...args, node: args.rowArg })
  const colNumber = args.colArg ? integerScalar({ ...args, node: args.colArg }) : cols === 1 ? 1 : undefined
  if (rowNumber === undefined || colNumber === undefined || rowNumber < 0 || colNumber < 0) {
    return []
  }
  const selectedRowStart = rowNumber === 0 ? rowStart : rowStart + rowNumber - 1
  const selectedRowEnd = rowNumber === 0 ? rowEnd : selectedRowStart
  const selectedColStart = colNumber === 0 ? colStart : colStart + colNumber - 1
  const selectedColEnd = colNumber === 0 ? colEnd : selectedColStart
  const selectedCellCount = (selectedRowEnd - selectedRowStart + 1) * (selectedColEnd - selectedColStart + 1)
  if (
    selectedRowStart < rowStart ||
    selectedRowEnd > rowEnd ||
    selectedColStart < colStart ||
    selectedColEnd > colEnd ||
    selectedCellCount <= 0 ||
    selectedCellCount > MAX_COMPACT_INDEX_DEPENDENCY_CELLS
  ) {
    return []
  }

  const cells: DynamicIndexSelectedCell[] = []
  for (let row = selectedRowStart; row <= selectedRowEnd; row += 1) {
    for (let col = selectedColStart; col <= selectedColEnd; col += 1) {
      cells.push({
        sheetName: bounds.sheetName,
        row,
        col,
        address: formatAddress(row, col),
      })
    }
  }
  return cells
}

function offsetReferenceBounds(node: FormulaNode | undefined, ownerSheetName: string): DynamicRangeBounds | undefined {
  if (!node) {
    return undefined
  }
  if (node.kind === 'CellRef') {
    try {
      const parsed = parseCellAddress(node.ref, node.sheetName ?? ownerSheetName)
      return {
        sheetName: parsed.sheetName ?? ownerSheetName,
        rowStart: parsed.row,
        rowEnd: parsed.row,
        colStart: parsed.col,
        colEnd: parsed.col,
      }
    } catch {
      return undefined
    }
  }
  return rangeBounds(node, ownerSheetName)
}

function selectedOffsetCells(args: {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly ownerSheetName: string
  readonly getFormulaAst?: DynamicFormulaAstResolver | undefined
  readonly node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>
}): DynamicIndexSelectedCell[] {
  if (args.node.args.length < 3 || args.node.args.length > 5) {
    return []
  }
  const reference = offsetReferenceBounds(args.node.args[0], args.ownerSheetName)
  if (!reference) {
    return []
  }
  const referenceRows = reference.rowEnd - reference.rowStart + 1
  const referenceCols = reference.colEnd - reference.colStart + 1
  const rowOffset = integerScalar({ ...args, node: args.node.args[1] })
  const colOffset = integerScalar({ ...args, node: args.node.args[2] })
  const height = args.node.args[3] ? integerScalar({ ...args, node: args.node.args[3] }) : referenceRows
  const width = args.node.args[4] ? integerScalar({ ...args, node: args.node.args[4] }) : referenceCols
  if (rowOffset === undefined || colOffset === undefined || height === undefined || width === undefined || height < 1 || width < 1) {
    return []
  }

  const rowStart = reference.rowStart + rowOffset
  const rowEnd = rowStart + height - 1
  const colStart = reference.colStart + colOffset
  const colEnd = colStart + width - 1
  const selectedCellCount = (rowEnd - rowStart + 1) * (colEnd - colStart + 1)
  if (
    rowStart < 0 ||
    colStart < 0 ||
    rowEnd < rowStart ||
    colEnd < colStart ||
    selectedCellCount <= 0 ||
    selectedCellCount > MAX_DYNAMIC_OFFSET_DEPENDENCY_CELLS
  ) {
    return []
  }

  const cells: DynamicIndexSelectedCell[] = []
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      cells.push({
        sheetName: reference.sheetName,
        row,
        col,
        address: formatAddress(row, col),
      })
    }
  }
  return cells
}

function cellWithStaticOffset(node: FormulaNode):
  | {
      readonly cell: Extract<FormulaNode, { readonly kind: 'CellRef' }>
      readonly offset: number
    }
  | undefined {
  if (node.kind === 'CellRef') {
    return { cell: node, offset: 0 }
  }
  if (
    node.kind === 'BinaryExpr' &&
    (node.operator === '+' || node.operator === '-') &&
    node.left.kind === 'CellRef' &&
    node.right.kind === 'NumberLiteral' &&
    Number.isInteger(node.right.value)
  ) {
    return { cell: node.left, offset: node.operator === '+' ? node.right.value : -node.right.value }
  }
  return undefined
}

function relativeCriterionSelectedRow(args: {
  readonly criterion: FormulaNode
  readonly range: DynamicRangeBounds
  readonly ownerSheetName: string
}): number | undefined {
  if (args.range.colStart !== args.range.colEnd) {
    return undefined
  }

  const parsed = cellWithStaticOffset(args.criterion)
  if (!parsed) {
    return undefined
  }
  const cell = parseCellAddress(parsed.cell.ref, parsed.cell.sheetName ?? args.ownerSheetName)
  const cellSheetName = cell.sheetName ?? args.ownerSheetName
  return cellSheetName === args.range.sheetName && cell.col === args.range.colStart ? cell.row + parsed.offset : undefined
}

function selectedCriteriaAggregateCells(args: {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly ownerSheetName: string
  readonly getFormulaAst?: DynamicFormulaAstResolver | undefined
  readonly node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>
}):
  | {
      readonly aggregateRangeDependency: string
      readonly selectedCells: readonly DynamicIndexSelectedCell[]
    }
  | undefined {
  const callee = args.node.callee.trim().toUpperCase()
  if (!isCriteriaAggregateCallee(callee)) {
    return undefined
  }
  const isSumIf = callee === 'SUMIF' || callee === 'AVERAGEIF'

  const aggregateNode = isSumIf ? (args.node.args[2] ?? args.node.args[0]) : args.node.args[0]
  const aggregateRangeDependency = aggregateNode?.kind === 'RangeRef' ? normalizeRangeDependency(aggregateNode) : undefined
  const aggregateRange = rangeBounds(aggregateNode, args.ownerSheetName)
  if (!aggregateRange || !aggregateRangeDependency) {
    return undefined
  }

  const criteriaPairs: Array<{ readonly range: DynamicRangeBounds; readonly criterion: FormulaNode }> = []
  if (isSumIf) {
    const range = rangeBounds(args.node.args[0], args.ownerSheetName)
    const criterion = args.node.args[1]
    if (!range || !criterion) {
      return undefined
    }
    criteriaPairs.push({ range, criterion })
  } else {
    if (args.node.args.length < 3 || args.node.args.length % 2 === 0) {
      return undefined
    }
    for (let index = 1; index < args.node.args.length; index += 2) {
      const range = rangeBounds(args.node.args[index], args.ownerSheetName)
      const criterion = args.node.args[index + 1]
      if (!range || !criterion) {
        return undefined
      }
      criteriaPairs.push({ range, criterion })
    }
  }

  const rowCount = aggregateRange.rowEnd - aggregateRange.rowStart + 1
  const colCount = aggregateRange.colEnd - aggregateRange.colStart + 1
  if (rowCount < 1 || colCount < 1) {
    return undefined
  }
  const hasVerticalCriteria = criteriaPairs.every(
    (pair) =>
      pair.range.rowEnd - pair.range.rowStart + 1 === rowCount &&
      pair.range.colStart === pair.range.colEnd &&
      pair.range.sheetName === aggregateRange.sheetName,
  )
  const hasHorizontalCriteria =
    rowCount === 1 &&
    criteriaPairs.every(
      (pair) =>
        pair.range.rowStart === pair.range.rowEnd &&
        pair.range.colEnd - pair.range.colStart + 1 === colCount &&
        pair.range.sheetName === aggregateRange.sheetName,
    )
  if (!hasVerticalCriteria && !hasHorizontalCriteria) {
    return undefined
  }
  const criteriaOrientation = hasVerticalCriteria ? 'vertical' : 'horizontal'

  const criteriaValues = criteriaPairs.map((pair) =>
    evaluateScalar({
      workbook: args.workbook,
      strings: args.strings,
      ownerSheetName: args.ownerSheetName,
      getFormulaAst: args.getFormulaAst,
      node: pair.criterion,
    }),
  )
  if (criteriaValues.every((value) => value !== undefined)) {
    const selectedCells: DynamicIndexSelectedCell[] = []
    let hadUnknownCriteriaCell = false
    const scanCount = criteriaOrientation === 'vertical' ? rowCount : colCount
    for (let offset = 0; offset < scanCount; offset += 1) {
      let matched = true
      for (let pairIndex = 0; pairIndex < criteriaPairs.length; pairIndex += 1) {
        const pair = criteriaPairs[pairIndex]!
        const row = criteriaOrientation === 'vertical' ? pair.range.rowStart + offset : pair.range.rowStart
        const col = criteriaOrientation === 'vertical' ? pair.range.colStart : pair.range.colStart + offset
        const address = formatAddress(row, col)
        const value = readCellValue({
          workbook: args.workbook,
          strings: args.strings,
          ownerSheetName: pair.range.sheetName,
          getFormulaAst: args.getFormulaAst,
          node: { kind: 'CellRef', sheetName: pair.range.sheetName, ref: address },
        })
        if (value.tag === ValueTag.Empty && args.getFormulaAst?.(pair.range.sheetName, address) === undefined) {
          hadUnknownCriteriaCell = true
          break
        }
        const didMatch = criterionMatches(value, criteriaValues[pairIndex]!)
        if (didMatch !== true) {
          matched = false
          break
        }
      }
      if (hadUnknownCriteriaCell) {
        break
      }
      if (!matched) {
        continue
      }
      const selectedRowStart = criteriaOrientation === 'vertical' ? aggregateRange.rowStart + offset : aggregateRange.rowStart
      const selectedRowEnd = criteriaOrientation === 'vertical' ? selectedRowStart : aggregateRange.rowEnd
      const selectedColStart = criteriaOrientation === 'vertical' ? aggregateRange.colStart : aggregateRange.colStart + offset
      const selectedColEnd = criteriaOrientation === 'vertical' ? aggregateRange.colEnd : selectedColStart
      for (let row = selectedRowStart; row <= selectedRowEnd; row += 1) {
        for (let col = selectedColStart; col <= selectedColEnd; col += 1) {
          selectedCells.push({
            sheetName: aggregateRange.sheetName,
            row,
            col,
            address: formatAddress(row, col),
          })
        }
      }
    }
    if (!hadUnknownCriteriaCell) {
      return { aggregateRangeDependency, selectedCells }
    }
  }

  if (criteriaOrientation !== 'vertical') {
    return undefined
  }

  const selectedRows = criteriaPairs.map((pair) =>
    relativeCriterionSelectedRow({ criterion: pair.criterion, range: pair.range, ownerSheetName: args.ownerSheetName }),
  )
  if (selectedRows.some((row) => row === undefined)) {
    return undefined
  }
  const selectedRow = selectedRows[0]!
  if (selectedRows.some((row) => row !== selectedRow)) {
    return undefined
  }
  if (selectedRow < aggregateRange.rowStart || selectedRow > aggregateRange.rowEnd) {
    return { aggregateRangeDependency, selectedCells: [] }
  }
  const selectedCells: DynamicIndexSelectedCell[] = []
  for (let col = aggregateRange.colStart; col <= aggregateRange.colEnd; col += 1) {
    selectedCells.push({
      sheetName: aggregateRange.sheetName,
      row: selectedRow,
      col,
      address: formatAddress(selectedRow, col),
    })
  }
  return { aggregateRangeDependency, selectedCells }
}

export function collectDynamicIndexDependencyPlan(args: {
  readonly compiled: ParsedCompiledFormula
  readonly ownerSheetName: string
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly getFormulaAst?: DynamicFormulaAstResolver | undefined
}): DynamicIndexDependencyPlan | undefined {
  const compactedRangeDependencies = new Set<string>()
  const selectedCells: DynamicIndexSelectedCell[] = []
  const selectedCellKeys = new Set<string>()

  const addSelectedCell = (cell: DynamicIndexSelectedCell): void => {
    const key = `${cell.sheetName}!${cell.address}`
    if (selectedCellKeys.has(key)) {
      return
    }
    selectedCellKeys.add(key)
    selectedCells.push(cell)
  }

  const visit = (node: FormulaNode, active: boolean): void => {
    if (node.kind === 'CallExpr') {
      const callee = node.callee.trim().toUpperCase()
      if (callee === 'IF' && node.args.length === 3) {
        visit(node.args[0]!, active)
        const choice = active
          ? branchChoice(
              evaluateScalar({
                workbook: args.workbook,
                strings: args.strings,
                ownerSheetName: args.ownerSheetName,
                getFormulaAst: args.getFormulaAst,
                node: node.args[0],
              }),
            )
          : 'no-branch'
        if (choice === 'truthy') {
          visit(node.args[1]!, true)
          visit(node.args[2]!, false)
        } else if (choice === 'falsy') {
          visit(node.args[1]!, false)
          visit(node.args[2]!, true)
        } else if (choice === 'no-branch') {
          visit(node.args[1]!, false)
          visit(node.args[2]!, false)
        } else {
          visit(node.args[1]!, true)
          visit(node.args[2]!, true)
        }
        return
      }
      if (callee === 'INDEX' && node.args[0]?.kind === 'RangeRef') {
        const rangeDependency = normalizeRangeDependency(node.args[0])
        if (rangeDependency) {
          compactedRangeDependencies.add(rangeDependency)
        }
        visit(node.args[1]!, active)
        if (node.args[2]) {
          visit(node.args[2], active)
        }
        if (active) {
          selectedIndexCells({
            workbook: args.workbook,
            strings: args.strings,
            ownerSheetName: args.ownerSheetName,
            getFormulaAst: args.getFormulaAst,
            range: node.args[0],
            rowArg: node.args[1],
            colArg: node.args[2],
          }).forEach(addSelectedCell)
        }
        return
      }
      if ((callee === 'ROWS' || callee === 'COLUMNS') && node.args.length === 1 && node.args[0]?.kind === 'RangeRef') {
        const rangeDependency = normalizeRangeDependency(node.args[0])
        if (rangeDependency) {
          compactedRangeDependencies.add(rangeDependency)
        }
        return
      }
      if (callee === 'OFFSET') {
        node.args.forEach((arg) => visit(arg, active))
        if (active) {
          selectedOffsetCells({
            workbook: args.workbook,
            strings: args.strings,
            ownerSheetName: args.ownerSheetName,
            getFormulaAst: args.getFormulaAst,
            node,
          }).forEach(addSelectedCell)
        }
        return
      }
      const criteriaAggregate = selectedCriteriaAggregateCells({
        workbook: args.workbook,
        strings: args.strings,
        ownerSheetName: args.ownerSheetName,
        getFormulaAst: args.getFormulaAst,
        node,
      })
      if (criteriaAggregate) {
        compactedRangeDependencies.add(criteriaAggregate.aggregateRangeDependency)
        node.args.forEach((arg) => visit(arg, active))
        if (active) {
          criteriaAggregate.selectedCells.forEach(addSelectedCell)
        }
        return
      }
      node.args.forEach((arg) => visit(arg, active))
      return
    }
    if (node.kind === 'UnaryExpr') {
      visit(node.argument, active)
      return
    }
    if (node.kind === 'BinaryExpr') {
      visit(node.left, active)
      visit(node.right, active)
      return
    }
    if (node.kind === 'InvokeExpr') {
      visit(node.callee, active)
      node.args.forEach((arg) => visit(arg, active))
    }
  }

  visit(args.compiled.optimizedAst, true)
  return compactedRangeDependencies.size === 0 && selectedCells.length === 0
    ? undefined
    : {
        compactedRangeDependencies,
        selectedCells,
      }
}
