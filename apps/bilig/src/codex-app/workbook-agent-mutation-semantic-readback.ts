import type { WorkbookRenderedReadbackProof } from './workbook-agent-rendered-readback.js'
import type { WorkbookAuthoritativeReadbackProof, WorkbookSemanticReadbackProof } from './workbook-agent-mutation-proof-types.js'

export function buildWorkbookSemanticReadbackProof(input: {
  readonly authoritativeReadback: WorkbookAuthoritativeReadbackProof
  readonly renderedReadback: WorkbookRenderedReadbackProof
}): WorkbookSemanticReadbackProof {
  const requested = input.authoritativeReadback.requested || input.renderedReadback.requested
  if (!requested) {
    return {
      requested: false,
      matched: null,
      incompleteReason: input.authoritativeReadback.incompleteReason ?? input.renderedReadback.incompleteReason,
    }
  }
  const matched =
    input.authoritativeReadback.matched === true && (!input.renderedReadback.requested || input.renderedReadback.matched === true)
  return {
    requested,
    matched,
    incompleteReason:
      input.authoritativeReadback.matched !== true
        ? (input.authoritativeReadback.incompleteReason ?? 'Authoritative semantic readback did not match.')
        : input.renderedReadback.requested && input.renderedReadback.matched !== true
          ? (input.renderedReadback.incompleteReason ?? 'Rendered semantic readback did not match.')
          : null,
  }
}
