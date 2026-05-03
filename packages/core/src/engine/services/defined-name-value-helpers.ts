import { parseFormula } from '@bilig/formula'
import type { WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'

export function isScalarOnlyDefinedNameValue(value: WorkbookDefinedNameValueSnapshot): boolean {
  if (typeof value === 'string' && value.startsWith('=')) {
    try {
      const ast = parseFormula(value)
      return ast.kind === 'NumberLiteral' || ast.kind === 'BooleanLiteral' || ast.kind === 'StringLiteral' || ast.kind === 'ErrorLiteral'
    } catch {
      return false
    }
  }
  if (value === null || typeof value !== 'object') {
    return true
  }
  if (!('kind' in value)) {
    return true
  }
  if (value.kind === 'scalar') {
    return true
  }
  if (value.kind !== 'formula') {
    return false
  }
  try {
    const ast = parseFormula(value.formula)
    return ast.kind === 'NumberLiteral' || ast.kind === 'BooleanLiteral' || ast.kind === 'StringLiteral' || ast.kind === 'ErrorLiteral'
  } catch {
    return false
  }
}
