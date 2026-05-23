import type { ReactNode } from 'react'

export type WorkbookUiSlot = 'toolbar' | 'sidePanel' | 'floatingOverlay' | 'status'

export interface WorkbookUiSlotContribution {
  readonly id: string
  readonly slot: WorkbookUiSlot
  readonly label: string
  readonly order?: number
  readonly render: () => ReactNode
}

export function sortWorkbookUiSlotContributions<T extends Pick<WorkbookUiSlotContribution, 'id' | 'order'>>(
  contributions: readonly T[],
): T[] {
  return [...contributions].toSorted((left, right) => {
    const orderDelta = (left.order ?? 0) - (right.order ?? 0)
    return orderDelta === 0 ? left.id.localeCompare(right.id) : orderDelta
  })
}

export function getWorkbookUiSlotContributions<T extends WorkbookUiSlotContribution>(
  contributions: readonly T[],
  slot: WorkbookUiSlot,
): T[] {
  return sortWorkbookUiSlotContributions(contributions.filter((contribution) => contribution.slot === slot))
}
