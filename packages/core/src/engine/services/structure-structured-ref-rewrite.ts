import type { FormulaNode } from '@bilig/formula'
import { parseFormula, serializeFormula } from '@bilig/formula'
import { ErrorCode } from '@bilig/protocol'
import type { DeletedTableColumnReference } from './structure-metadata-rewrite.js'

export function rewriteFormulaSourceForDeletedStructuredReferences(
  source: string,
  deletedReferences: readonly DeletedTableColumnReference[],
): string | undefined {
  if (deletedReferences.length === 0) {
    return undefined
  }
  const deleted = new Set(deletedReferences.map(({ tableName, columnName }) => structuredReferenceKey(tableName, columnName)))
  const rewrite = rewriteDeletedStructuredReferenceNode(parseFormula(source), deleted)
  return rewrite.changed ? serializeFormula(rewrite.node) : undefined
}

function rewriteDeletedStructuredReferenceNode(
  node: FormulaNode,
  deletedReferences: ReadonlySet<string>,
): { readonly node: FormulaNode; readonly changed: boolean } {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'CellRef':
    case 'SpillRef':
    case 'ColumnRef':
    case 'RowRef':
    case 'RangeRef':
      return { node, changed: false }
    case 'StructuredRef':
      return deletedReferences.has(structuredReferenceKey(node.tableName, node.columnName))
        ? { node: { kind: 'ErrorLiteral', code: ErrorCode.Ref }, changed: true }
        : { node, changed: false }
    case 'ArrayConstant': {
      let changed = false
      const rows = node.rows.map((row) =>
        row.map((entry) => {
          const rewrite = rewriteDeletedStructuredReferenceNode(entry, deletedReferences)
          changed ||= rewrite.changed
          return rewrite.node
        }),
      )
      return changed ? { node: { ...node, rows }, changed } : { node, changed: false }
    }
    case 'UnaryExpr': {
      const argument = rewriteDeletedStructuredReferenceNode(node.argument, deletedReferences)
      return argument.changed ? { node: { ...node, argument: argument.node }, changed: true } : { node, changed: false }
    }
    case 'BinaryExpr': {
      const left = rewriteDeletedStructuredReferenceNode(node.left, deletedReferences)
      const right = rewriteDeletedStructuredReferenceNode(node.right, deletedReferences)
      return left.changed || right.changed
        ? { node: { ...node, left: left.node, right: right.node }, changed: true }
        : { node, changed: false }
    }
    case 'CallExpr': {
      let changed = false
      const args = node.args.map((arg) => {
        const rewrite = rewriteDeletedStructuredReferenceNode(arg, deletedReferences)
        changed ||= rewrite.changed
        return rewrite.node
      })
      return changed ? { node: { ...node, args }, changed } : { node, changed: false }
    }
    case 'InvokeExpr': {
      const callee = rewriteDeletedStructuredReferenceNode(node.callee, deletedReferences)
      let argsChanged = false
      const args = node.args.map((arg) => {
        const rewrite = rewriteDeletedStructuredReferenceNode(arg, deletedReferences)
        argsChanged ||= rewrite.changed
        return rewrite.node
      })
      return callee.changed || argsChanged ? { node: { ...node, callee: callee.node, args }, changed: true } : { node, changed: false }
    }
  }
}

function structuredReferenceKey(tableName: string, columnName: string): string {
  return `${tableName.trim().toUpperCase()}\u0000${columnName.trim().toUpperCase()}`
}
