import type { RuntimeFormula } from '../runtime-state.js'
import { tryCompileSimpleDirectScalarFormula } from '../../formula/simple-direct-scalar-compile.js'
import type { ParsedCompiledFormula } from './formula-binding-direct-descriptors.js'

export function tryRewriteSimpleDirectScalarFormulaSourcePreservingBinding(
  existing: RuntimeFormula | undefined,
  definedNameCount: number,
  source: string,
  rewriteCompiled: (compiled: ParsedCompiledFormula) => boolean,
): boolean {
  if (
    existing?.directScalar === undefined ||
    existing.directAggregate !== undefined ||
    existing.directCriteria !== undefined ||
    existing.directLookup !== undefined ||
    definedNameCount !== 0
  ) {
    return false
  }
  const compiled = tryCompileSimpleDirectScalarFormula(source)
  return compiled !== undefined && rewriteCompiled(compiled as ParsedCompiledFormula)
}
