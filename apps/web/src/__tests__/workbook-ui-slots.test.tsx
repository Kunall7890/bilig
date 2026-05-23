import { describe, expect, it } from 'vitest'
import { getWorkbookUiSlotContributions } from '../workbook-ui-slots.js'

describe('workbook UI slots', () => {
  it('returns matching contributions in stable order', () => {
    const contributions = getWorkbookUiSlotContributions(
      [
        { id: 'z', slot: 'sidePanel', label: 'Z', order: 20, render: () => null },
        { id: 'a', slot: 'toolbar', label: 'A', order: 10, render: () => null },
        { id: 'b', slot: 'sidePanel', label: 'B', order: 10, render: () => null },
        { id: 'a-side', slot: 'sidePanel', label: 'A Side', order: 10, render: () => null },
      ],
      'sidePanel',
    )

    expect(contributions.map((contribution) => contribution.id)).toEqual(['a-side', 'b', 'z'])
  })
})
