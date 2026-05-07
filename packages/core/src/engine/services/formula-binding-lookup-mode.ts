import { FormulaMode } from '@bilig/protocol'
import { hasDirectApproximateLookupCandidate, hasIndexedExactLookupCandidate } from './formula-binding-lookup-candidates.js'
import type { ParsedCompiledFormula } from './formula-binding-direct-descriptors.js'

export function normalizeFormulaBindingLookupCompileMode(compiled: ParsedCompiledFormula): ParsedCompiledFormula {
  if (compiled.mode !== FormulaMode.WasmFastPath) {
    return compiled
  }
  if (!hasIndexedExactLookupCandidate(compiled.optimizedAst) && !hasDirectApproximateLookupCandidate(compiled.optimizedAst)) {
    return compiled
  }
  return {
    ...compiled,
    mode: FormulaMode.JsOnly,
  }
}
