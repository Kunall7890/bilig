import type { FormulaNode } from '@bilig/formula'
import type { ErrorCode } from '@bilig/protocol'

export function dominantScalarErrorForCycleFormula(node: FormulaNode): ErrorCode | undefined {
  if (node.kind === 'ErrorLiteral') {
    return node.code as ErrorCode
  }
  if (node.kind === 'UnaryExpr') {
    return dominantScalarErrorForCycleFormula(node.argument)
  }
  if (node.kind === 'BinaryExpr') {
    return node.operator === ':'
      ? undefined
      : (dominantScalarErrorForCycleFormula(node.left) ?? dominantScalarErrorForCycleFormula(node.right))
  }
  return undefined
}
