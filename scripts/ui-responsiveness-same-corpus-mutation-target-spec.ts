import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function sameCorpusMutationTargetRangeForSample(workload: UiResponsivenessSameCorpusMutatingWorkload, sampleIndex: number): string {
  const columnByWorkload: Record<UiResponsivenessSameCorpusMutatingWorkload, string> = {
    'edit-visible-cell': 'C',
    'fill-format-change': 'E',
    'formula-edit': 'D',
  }
  return `${columnByWorkload[workload]}${String(sampleIndex + 5)}`
}
