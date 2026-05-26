export const requiredUiResponsivenessSameCorpusWorkloads = [
  'open-workbook',
  'select-cell',
  'edit-visible-cell',
  'scroll-vertical',
  'scroll-horizontal',
  'jump-deep-row',
  'formula-edit',
  'fill-format-change',
  'wide-sheet-navigation',
] as const

export type UiResponsivenessSameCorpusWorkload = (typeof requiredUiResponsivenessSameCorpusWorkloads)[number]
export type UiResponsivenessSameCorpusMutatingWorkload = Extract<
  UiResponsivenessSameCorpusWorkload,
  'edit-visible-cell' | 'formula-edit' | 'fill-format-change'
>

export function isUiResponsivenessSameCorpusWorkload(value: string): value is UiResponsivenessSameCorpusWorkload {
  return (requiredUiResponsivenessSameCorpusWorkloads as readonly string[]).includes(value)
}

export function uiSameCorpusWorkloadRequiresScrollEventEvidence(workload: UiResponsivenessSameCorpusWorkload): boolean {
  return workload === 'scroll-vertical' || workload === 'scroll-horizontal' || workload === 'wide-sheet-navigation'
}

export function uiSameCorpusWorkloadMutatesWorkbook(
  workload: UiResponsivenessSameCorpusWorkload,
): workload is UiResponsivenessSameCorpusMutatingWorkload {
  return workload === 'edit-visible-cell' || workload === 'formula-edit' || workload === 'fill-format-change'
}
