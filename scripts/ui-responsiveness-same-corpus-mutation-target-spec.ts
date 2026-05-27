import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function sameCorpusMutationTargetRangeForSample(workload: UiResponsivenessSameCorpusMutatingWorkload, sampleIndex: number): string {
  const columnByWorkload: Record<UiResponsivenessSameCorpusMutatingWorkload, string> = {
    'edit-visible-cell': 'F',
    'fill-format-change': 'B',
    'formula-edit': 'C',
  }
  return `${columnByWorkload[workload]}${String(sampleIndex + 5)}`
}
