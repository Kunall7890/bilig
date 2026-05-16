import { excelPower, parseCellAddress, parseRangeAddress, type FormulaNode } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import type { StringPool } from '../../string-pool.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { residentRangeShape } from './formula-binding-dynamic-range-bounds.js'

export type BranchChoice = 'truthy' | 'falsy' | 'no-branch' | 'unknown'

export interface DynamicIndexFormulaAst {
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
  readonly ast: FormulaNode
}

export type DynamicFormulaAstResolver = (sheetName: string, address: string) => DynamicIndexFormulaAst | undefined

export interface DynamicScalarContext {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly ownerSheetName: string
  readonly currentRow?: number
  readonly currentCol?: number
  readonly getFormulaAst?: DynamicFormulaAstResolver | undefined
  readonly visitingFormulaCells?: ReadonlySet<string>
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function booleanValue(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty }
}

function literalInputToCellValue(value: LiteralInput): CellValue {
  if (value === null) {
    return emptyValue()
  }
  if (typeof value === 'number') {
    return numberValue(value)
  }
  if (typeof value === 'boolean') {
    return booleanValue(value)
  }
  return { tag: ValueTag.String, value, stringId: 0 }
}

export function readCellValue(args: {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly ownerSheetName: string
  readonly getFormulaAst?: DynamicFormulaAstResolver | undefined
  readonly visitingFormulaCells?: ReadonlySet<string>
  readonly node: Extract<FormulaNode, { readonly kind: 'CellRef' }>
}): CellValue {
  const sheetName = args.node.sheetName ?? args.ownerSheetName
  const cellIndex = args.workbook.getCellIndex(sheetName, args.node.ref)
  if (cellIndex === undefined) {
    return emptyValue()
  }
  const value = args.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : args.strings.get(stringId)))
  if (value.tag !== ValueTag.Empty) {
    return value
  }

  const formula = args.getFormulaAst?.(sheetName, args.node.ref)
  const formulaKey = formula ? `${formula.sheetName}!${formula.address}` : undefined
  if (!formula || !formulaKey || args.visitingFormulaCells?.has(formulaKey) === true) {
    return value
  }
  const visitingFormulaCells = new Set(args.visitingFormulaCells)
  visitingFormulaCells.add(formulaKey)
  return (
    evaluateScalar({
      workbook: args.workbook,
      strings: args.strings,
      ownerSheetName: formula.sheetName,
      currentRow: formula.row,
      currentCol: formula.col,
      getFormulaAst: args.getFormulaAst,
      visitingFormulaCells,
      node: formula.ast,
    }) ?? value
  )
}

function coerceNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
      return undefined
  }
}

function concatText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(Object.is(value.value, -0) ? 0 : value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return String(value.code)
  }
}

function compareText(left: string, right: string): number {
  const normalizedLeft = left.toUpperCase()
  const normalizedRight = right.toUpperCase()
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}

function compareScalar(left: CellValue, right: CellValue, operator: '=' | '<>' | '>' | '>=' | '<' | '<='): CellValue {
  if (left.tag === ValueTag.Error) {
    return left
  }
  if (right.tag === ValueTag.Error) {
    return right
  }

  const leftNumber = coerceNumber(left)
  const rightNumber = coerceNumber(right)
  if (leftNumber !== undefined && rightNumber !== undefined) {
    switch (operator) {
      case '=':
        return booleanValue(leftNumber === rightNumber)
      case '<>':
        return booleanValue(leftNumber !== rightNumber)
      case '>':
        return booleanValue(leftNumber > rightNumber)
      case '>=':
        return booleanValue(leftNumber >= rightNumber)
      case '<':
        return booleanValue(leftNumber < rightNumber)
      case '<=':
        return booleanValue(leftNumber <= rightNumber)
    }
  }

  let comparison: number | undefined
  if (left.tag === ValueTag.String && right.tag === ValueTag.String) {
    comparison = compareText(left.value, right.value)
  } else if (left.tag === ValueTag.Empty && right.tag === ValueTag.Empty) {
    comparison = 0
  } else if (left.tag === ValueTag.String && right.tag === ValueTag.Empty) {
    comparison = compareText(left.value, '')
  } else if (left.tag === ValueTag.Empty && right.tag === ValueTag.String) {
    comparison = compareText('', right.value)
  } else if (left.tag === ValueTag.String && (right.tag === ValueTag.Number || right.tag === ValueTag.Boolean)) {
    comparison = 1
  } else if ((left.tag === ValueTag.Number || left.tag === ValueTag.Boolean) && right.tag === ValueTag.String) {
    comparison = -1
  }

  if (comparison !== undefined) {
    switch (operator) {
      case '=':
        return booleanValue(comparison === 0)
      case '<>':
        return booleanValue(comparison !== 0)
      case '>':
        return booleanValue(comparison > 0)
      case '>=':
        return booleanValue(comparison >= 0)
      case '<':
        return booleanValue(comparison < 0)
      case '<=':
        return booleanValue(comparison <= 0)
    }
  }

  return errorValue(ErrorCode.Value)
}

function criterionTextParts(value: string): { readonly operator: '=' | '<>' | '>' | '>=' | '<' | '<='; readonly operand: string } {
  const trimmed = value.trim()
  if (trimmed.startsWith('>=')) {
    return { operator: '>=', operand: trimmed.slice(2).trim() }
  }
  if (trimmed.startsWith('<=')) {
    return { operator: '<=', operand: trimmed.slice(2).trim() }
  }
  if (trimmed.startsWith('<>')) {
    return { operator: '<>', operand: trimmed.slice(2).trim() }
  }
  if (trimmed.startsWith('>')) {
    return { operator: '>', operand: trimmed.slice(1).trim() }
  }
  if (trimmed.startsWith('<')) {
    return { operator: '<', operand: trimmed.slice(1).trim() }
  }
  if (trimmed.startsWith('=')) {
    return { operator: '=', operand: trimmed.slice(1).trim() }
  }
  return { operator: '=', operand: trimmed }
}

function criterionOperandValue(criterion: CellValue): CellValue {
  if (criterion.tag !== ValueTag.String) {
    return criterion
  }
  const numeric = Number(criterion.value)
  return criterion.value.trim() !== '' && Number.isFinite(numeric)
    ? numberValue(numeric)
    : { tag: ValueTag.String, value: criterion.value, stringId: 0 }
}

export function criterionMatches(value: CellValue, criterion: CellValue): boolean | undefined {
  if (value.tag === ValueTag.Error || criterion.tag === ValueTag.Error) {
    return undefined
  }
  if (criterion.tag === ValueTag.String) {
    const parts = criterionTextParts(criterion.value)
    const operand = criterionOperandValue({ tag: ValueTag.String, value: parts.operand, stringId: 0 })
    const result = compareScalar(value, operand, parts.operator)
    return result.tag === ValueTag.Boolean ? result.value : undefined
  }
  const result = compareScalar(value, criterion, '=')
  return result.tag === ValueTag.Boolean ? result.value : undefined
}

export function evaluateScalar(args: {
  readonly workbook: WorkbookStore
  readonly strings: StringPool
  readonly ownerSheetName: string
  readonly currentRow?: number
  readonly currentCol?: number
  readonly getFormulaAst?: DynamicFormulaAstResolver | undefined
  readonly visitingFormulaCells?: ReadonlySet<string>
  readonly node: FormulaNode | undefined
}): CellValue | undefined {
  const node = args.node
  if (!node) {
    return undefined
  }
  switch (node.kind) {
    case 'NumberLiteral':
      return numberValue(node.value)
    case 'BooleanLiteral':
      return booleanValue(node.value)
    case 'StringLiteral':
      return { tag: ValueTag.String, value: node.value, stringId: 0 }
    case 'ErrorLiteral':
      return errorValue(node.code as ErrorCode)
    case 'OmittedArgument':
      return undefined
    case 'ArrayConstant':
      return undefined
    case 'CellRef':
      return readCellValue({ ...args, node })
    case 'NameRef': {
      const definedName = args.workbook.getDefinedName(node.name, args.ownerSheetName)
      if (!definedName) {
        return undefined
      }
      const value = definedName.value
      if (typeof value !== 'object' || value === null || !('kind' in value)) {
        return literalInputToCellValue(value)
      }
      if (value.kind === 'scalar') {
        return literalInputToCellValue(value.value)
      }
      if (value.kind === 'cell-ref') {
        const cellIndex = args.workbook.getCellIndex(value.sheetName, value.address)
        return cellIndex === undefined
          ? emptyValue()
          : args.workbook.cellStore.getValue(cellIndex, (stringId) => (stringId === 0 ? '' : args.strings.get(stringId)))
      }
      return undefined
    }
    case 'UnaryExpr': {
      const value = evaluateScalar({ ...args, node: node.argument })
      const numeric = value ? coerceNumber(value) : undefined
      return numeric === undefined ? undefined : numberValue(node.operator === '-' ? -numeric : numeric)
    }
    case 'BinaryExpr': {
      const left = evaluateScalar({ ...args, node: node.left })
      const right = evaluateScalar({ ...args, node: node.right })
      if (!left || !right) {
        return undefined
      }
      if (node.operator === '&') {
        if (left.tag === ValueTag.Error) {
          return left
        }
        if (right.tag === ValueTag.Error) {
          return right
        }
        return stringValue(`${concatText(left)}${concatText(right)}`)
      }
      if (
        node.operator === '=' ||
        node.operator === '<>' ||
        node.operator === '>' ||
        node.operator === '>=' ||
        node.operator === '<' ||
        node.operator === '<='
      ) {
        return compareScalar(left, right, node.operator)
      }
      const leftNumber = coerceNumber(left)
      const rightNumber = coerceNumber(right)
      if (leftNumber === undefined || rightNumber === undefined) {
        return undefined
      }
      switch (node.operator) {
        case '+':
          return numberValue(leftNumber + rightNumber)
        case '-':
          return numberValue(leftNumber - rightNumber)
        case '*':
          return numberValue(leftNumber * rightNumber)
        case '/':
          return rightNumber === 0 ? errorValue(ErrorCode.Div0) : numberValue(leftNumber / rightNumber)
        case '^': {
          const value = excelPower(leftNumber, rightNumber)
          return Number.isFinite(value) ? numberValue(value) : errorValue(ErrorCode.Value)
        }
        case ':':
          return undefined
      }
    }
    case 'CallExpr': {
      const callee = node.callee.trim().toUpperCase()
      if (callee === 'ROW') {
        if (node.args.length === 0) {
          return args.currentRow === undefined ? undefined : numberValue(args.currentRow + 1)
        }
        const reference = node.args[0]
        if (reference?.kind === 'CellRef') {
          try {
            return numberValue(parseCellAddress(reference.ref, reference.sheetName ?? args.ownerSheetName).row + 1)
          } catch {
            return undefined
          }
        }
        if (reference?.kind === 'RangeRef' && reference.refKind === 'cells') {
          try {
            const parsed = parseRangeAddress(`${reference.start}:${reference.end}`, reference.sheetName ?? args.ownerSheetName)
            return parsed.kind === 'cells' ? numberValue(parsed.start.row + 1) : undefined
          } catch {
            return undefined
          }
        }
      }
      if ((callee === 'ROWS' || callee === 'COLUMNS') && node.args.length === 1 && node.args[0]?.kind === 'RangeRef') {
        const shape = residentRangeShape({
          workbook: args.workbook,
          ownerSheetName: args.ownerSheetName,
          range: node.args[0],
        })
        if (shape) {
          return numberValue(callee === 'ROWS' ? shape.rows : shape.cols)
        }
      }
      return undefined
    }
    case 'ColumnRef':
    case 'InvokeExpr':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
      return undefined
  }
}

export function branchChoice(value: CellValue | undefined): BranchChoice {
  if (!value) {
    return 'unknown'
  }
  switch (value.tag) {
    case ValueTag.Boolean:
      return value.value ? 'truthy' : 'falsy'
    case ValueTag.Number:
      return value.value === 0 ? 'falsy' : 'truthy'
    case ValueTag.Empty:
      return 'falsy'
    case ValueTag.Error:
    case ValueTag.String:
      return 'no-branch'
  }
}

export function integerScalar(args: DynamicScalarContext & { readonly node: FormulaNode | undefined }): number | undefined {
  const value = evaluateScalar(args)
  const numeric = value ? coerceNumber(value) : undefined
  return numeric !== undefined && Number.isInteger(numeric) ? numeric : undefined
}
