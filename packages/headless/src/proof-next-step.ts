export interface ProofNextStep {
  readonly ifUseful: string
  readonly star: string
  readonly watchReleases: string
  readonly ifBlocked: string
  readonly adoptionBlocker: string
}

export const proofStarUrl = 'https://github.com/proompteng/bilig/stargazers'
export const proofWatchReleasesUrl = 'https://github.com/proompteng/bilig/subscription'
export const proofAdoptionBlockerUrl = 'https://github.com/proompteng/bilig/discussions/new?category=general'

export function createProofNextStep(workflowLabel: string): ProofNextStep {
  return {
    ifUseful: `If this ${workflowLabel} proof matched your workflow, star or bookmark Bilig so you can find it again.`,
    star: proofStarUrl,
    watchReleases: proofWatchReleasesUrl,
    ifBlocked: 'If it almost worked, open the concrete workbook or agent blocker.',
    adoptionBlocker: proofAdoptionBlockerUrl,
  }
}
