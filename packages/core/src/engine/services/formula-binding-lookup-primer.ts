import type { RuntimeFormula } from '../runtime-state.js'
import type { collectDirectApproximateLookupCandidates, collectIndexedExactLookupCandidates } from './formula-binding-lookup-candidates.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export function primeFormulaBindingLookupCandidates(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly ownerSheetName: string
  readonly directLookup: RuntimeFormula['directLookup']
  readonly indexedExactLookupCandidates: ReturnType<typeof collectIndexedExactLookupCandidates>
  readonly directApproximateLookupCandidates: ReturnType<typeof collectDirectApproximateLookupCandidates>
}): void {
  if (args.directLookup) {
    return
  }
  args.indexedExactLookupCandidates.forEach((candidate) => {
    if (candidate.startCol !== candidate.endCol) {
      return
    }
    args.serviceArgs.exactLookup.primeColumnIndex({
      sheetName: candidate.sheetName ?? args.ownerSheetName,
      rowStart: candidate.startRow,
      rowEnd: candidate.endRow,
      col: candidate.startCol,
    })
  })
  args.directApproximateLookupCandidates.forEach((candidate) => {
    if (candidate.startCol !== candidate.endCol) {
      return
    }
    args.serviceArgs.sortedLookup.primeColumnIndex({
      sheetName: candidate.sheetName ?? args.ownerSheetName,
      rowStart: candidate.startRow,
      rowEnd: candidate.endRow,
      col: candidate.startCol,
    })
  })
}
