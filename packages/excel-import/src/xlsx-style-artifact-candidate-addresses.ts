import type { collectStyleCandidateAddresses } from './xlsx-import-cell-styles.js'
import type { readImportedWorkbookStyleArtifacts } from './xlsx-styles.js'

function addCandidateAddress(addressesBySheet: Map<string, Set<string>>, sheetName: string, address: string): boolean {
  const addresses = addressesBySheet.get(sheetName) ?? new Set<string>()
  const previousSize = addresses.size
  addresses.add(address)
  if (addresses.size > previousSize) {
    addressesBySheet.set(sheetName, addresses)
    return true
  }
  return false
}

export function addStyleArtifactCandidateAddresses(
  candidates: ReturnType<typeof collectStyleCandidateAddresses>,
  importedStyleArtifacts: ReturnType<typeof readImportedWorkbookStyleArtifacts>,
  maxCandidateCount: number,
): ReturnType<typeof collectStyleCandidateAddresses> {
  let count = candidates.count
  const addressesBySheet = new Map([...candidates.addressesBySheet].map(([sheetName, addresses]) => [sheetName, new Set(addresses)]))
  for (const [sheetName, artifacts] of importedStyleArtifacts.sheetArtifactsByName) {
    for (const entry of artifacts.cellStyleIndexes) {
      if (addCandidateAddress(addressesBySheet, sheetName, entry.address)) {
        count += 1
      }
      if (count > maxCandidateCount) {
        return candidates
      }
    }
    for (const address of artifacts.blankCellAddresses ?? []) {
      if (addCandidateAddress(addressesBySheet, sheetName, address)) {
        count += 1
      }
      if (count > maxCandidateCount) {
        return candidates
      }
    }
  }
  return { addressesBySheet, count }
}
