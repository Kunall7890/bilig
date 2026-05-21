import type { FormulaNode } from './ast.js'
import { parseRangeAddress } from './addressing.js'
import { rewriteSpecialCall } from './special-call-rewrites.js'

const VOLATILE_BUILTINS = new Set(['TODAY', 'NOW', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT', 'SUBTOTAL', 'AGGREGATE'])

export const FORMULA_SPILL_PRODUCING_FUNCTION_NAMES = [
  'SEQUENCE',
  'EXPAND',
  'LINEST',
  'LOGEST',
  'OFFSET',
  'TAKE',
  'DROP',
  'CHOOSECOLS',
  'CHOOSEROWS',
  'SORT',
  'SORTBY',
  'TOCOL',
  'TOROW',
  'WRAPROWS',
  'WRAPCOLS',
  'FILTER',
  'UNIQUE',
  'FREQUENCY',
  'MODE.MULT',
  'TEXTSPLIT',
  'TRIMRANGE',
  'GROUPBY',
  'PIVOTBY',
  'MAKEARRAY',
  'MAP',
  'SCAN',
  'BYROW',
  'BYCOL',
  'RANDARRAY',
  'MUNIT',
  'MINVERSE',
  'MMULT',
] as const

const SPILL_PRODUCING_BUILTINS = new Set<string>(FORMULA_SPILL_PRODUCING_FUNCTION_NAMES)

export interface VolatileMetadata {
  volatile: boolean
  randCallCount: number
}

function cellRangeShape(node: FormulaNode): { rows: number; cols: number } | undefined {
  if (node.kind !== 'RangeRef' || node.refKind !== 'cells' || node.sheetEndName !== undefined) {
    return undefined
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : ''
    const range = parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`)
    if (range.kind !== 'cells') {
      return undefined
    }
    return {
      rows: range.end.row - range.start.row + 1,
      cols: range.end.col - range.start.col + 1,
    }
  } catch {
    return undefined
  }
}

function xlookupProducesSpill(args: readonly FormulaNode[]): boolean {
  const [lookupValue, lookupArray, returnArray] = args
  if (!lookupValue || !lookupArray || !returnArray) {
    return false
  }
  if (producesSpillResult(lookupValue)) {
    return true
  }

  const lookupShape = cellRangeShape(lookupArray)
  const returnShape = cellRangeShape(returnArray)
  if (!lookupShape || !returnShape) {
    return producesSpillResult(returnArray)
  }
  if (lookupShape.cols === 1) {
    return returnShape.rows === lookupShape.rows && returnShape.cols > 1
  }
  if (lookupShape.rows === 1) {
    return returnShape.cols === lookupShape.cols && returnShape.rows > 1
  }
  return false
}

export function producesSpillResult(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'InvokeExpr':
      return false
    case 'ArrayConstant':
      return true
    case 'RangeRef':
      return true
    case 'UnaryExpr':
      return producesSpillResult(node.argument)
    case 'BinaryExpr':
      return producesSpillResult(node.left) || producesSpillResult(node.right)
    case 'CallExpr':
      if (node.callee.toUpperCase() === 'CHOOSE') {
        return node.args.some((arg) => producesSpillResult(arg))
      }
      if (node.callee.toUpperCase() === 'XLOOKUP') {
        return xlookupProducesSpill(node.args)
      }
      if (node.callee.toUpperCase() === 'TREND' || node.callee.toUpperCase() === 'GROWTH') {
        const shapeArg = node.args[2] ?? node.args[1] ?? node.args[0]
        if (shapeArg === undefined) {
          return false
        }
        return producesSpillResult(shapeArg)
      }
      return SPILL_PRODUCING_BUILTINS.has(node.callee.toUpperCase())
  }
}

export function analyzeVolatileMetadata(node: FormulaNode): VolatileMetadata {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return { volatile: false, randCallCount: 0 }
    case 'ArrayConstant': {
      const children = node.rows.flatMap((row) => row.map(analyzeVolatileMetadata))
      return {
        volatile: children.some((child) => child.volatile),
        randCallCount: children.reduce((sum, child) => sum + child.randCallCount, 0),
      }
    }
    case 'UnaryExpr':
      return analyzeVolatileMetadata(node.argument)
    case 'BinaryExpr': {
      const left = analyzeVolatileMetadata(node.left)
      const right = analyzeVolatileMetadata(node.right)
      return {
        volatile: left.volatile || right.volatile,
        randCallCount: left.randCallCount + right.randCallCount,
      }
    }
    case 'CallExpr': {
      const rewritten = rewriteSpecialCall(node)
      if (rewritten) {
        return analyzeVolatileMetadata(rewritten)
      }
      const callee = node.callee.toUpperCase()
      let volatile = VOLATILE_BUILTINS.has(callee)
      let randCallCount = callee === 'RAND' ? 1 : 0
      node.args.forEach((arg) => {
        const child = analyzeVolatileMetadata(arg)
        volatile = volatile || child.volatile
        randCallCount += child.randCallCount
      })
      return { volatile, randCallCount }
    }
    case 'InvokeExpr': {
      const callee = analyzeVolatileMetadata(node.callee)
      const args = node.args.map(analyzeVolatileMetadata)
      return {
        volatile: callee.volatile || args.some((child) => child.volatile),
        randCallCount: callee.randCallCount + args.reduce((sum, child) => sum + child.randCallCount, 0),
      }
    }
  }
}
