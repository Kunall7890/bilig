import { ErrorCode } from '@bilig/protocol'
import type { BinaryExprNode, CallExprNode, FormulaNode } from './ast.js'

function errorNode(code: ErrorCode): FormulaNode {
  return { kind: 'ErrorLiteral', code }
}

function booleanNode(value: boolean): FormulaNode {
  return { kind: 'BooleanLiteral', value }
}

function callNode(callee: string, args: FormulaNode[]): CallExprNode {
  return { kind: 'CallExpr', callee: callee.toUpperCase(), args }
}

function binaryNode(operator: BinaryExprNode['operator'], left: FormulaNode, right: FormulaNode): FormulaNode {
  return { kind: 'BinaryExpr', operator, left, right }
}

function isTextLikeNode(node: FormulaNode): FormulaNode {
  return callNode('OR', [callNode('ISTEXT', [node]), callNode('ISBLANK', [node])])
}

function isLogicalLikeNode(node: FormulaNode): FormulaNode {
  return callNode('NOT', [callNode('OR', [callNode('ISNUMBER', [node]), callNode('ISTEXT', [node]), callNode('ISBLANK', [node])])])
}

function switchMatchCondition(expression: FormulaNode, candidate: FormulaNode): FormulaNode {
  const valueEqualsCandidate = binaryNode('=', expression, candidate)
  return callNode('OR', [
    callNode('AND', [callNode('ISNUMBER', [expression]), callNode('ISNUMBER', [candidate]), valueEqualsCandidate]),
    callNode('AND', [isLogicalLikeNode(expression), isLogicalLikeNode(candidate), valueEqualsCandidate]),
    callNode('AND', [isTextLikeNode(expression), isTextLikeNode(candidate), valueEqualsCandidate]),
  ])
}

function rewriteIfs(args: readonly FormulaNode[]): FormulaNode {
  if (args.length < 2 || args.length % 2 !== 0) {
    return errorNode(ErrorCode.Value)
  }

  let fallback: FormulaNode = errorNode(ErrorCode.NA)
  for (let index = args.length - 2; index >= 0; index -= 2) {
    fallback = callNode('IF', [args[index]!, args[index + 1]!, fallback])
  }
  return fallback
}

function rewriteSwitch(args: readonly FormulaNode[]): FormulaNode {
  if (args.length < 3) {
    return errorNode(ErrorCode.Value)
  }

  const expression = args[0]!
  const entries = args.slice(1)
  if (entries.length < 2) {
    return errorNode(ErrorCode.Value)
  }

  const hasDefault = entries.length % 2 === 1
  let fallback: FormulaNode = hasDefault ? entries[entries.length - 1]! : errorNode(ErrorCode.NA)
  const pairLimit = hasDefault ? entries.length - 1 : entries.length
  for (let index = pairLimit - 2; index >= 0; index -= 2) {
    fallback = callNode('IF', [switchMatchCondition(expression, entries[index]!), entries[index + 1]!, fallback])
  }
  return fallback
}

export function rewriteSpecialCall(node: CallExprNode): FormulaNode | undefined {
  switch (node.callee.toUpperCase()) {
    case 'IF':
      return node.args.length === 2 ? callNode('IF', [node.args[0]!, node.args[1]!, booleanNode(false)]) : undefined
    case 'TRUE':
      return node.args.length === 0 ? booleanNode(true) : errorNode(ErrorCode.Value)
    case 'FALSE':
      return node.args.length === 0 ? booleanNode(false) : errorNode(ErrorCode.Value)
    case 'IFS':
      return rewriteIfs(node.args)
    case 'SWITCH':
      return rewriteSwitch(node.args)
    default:
      return undefined
  }
}
