// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { WorkbookTablesPanel } from '../WorkbookTablesPanel.js'

const roots: Array<{ unmount: () => void }> = []

afterEach(() => {
  roots.splice(0).forEach((root) => root.unmount())
})

describe('WorkbookTablesPanel', () => {
  it('renders table metadata and dispatches table actions', async () => {
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)
    roots.push(root)
    const onCreateFromSelection = vi.fn()
    const onDeleteTable = vi.fn()
    const onResizeTable = vi.fn()
    const onSelectRange = vi.fn()
    const selectionRange = {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C6',
    }

    await act(async () => {
      root.render(
        <WorkbookTablesPanel
          activeTableName="Sales"
          selectionRange={selectionRange}
          tables={[
            {
              name: 'Sales',
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'B5',
              columnNames: ['Region', 'Amount'],
              columns: [{ name: 'Region' }, { name: 'Amount' }],
              headerRow: true,
              totalsRow: false,
              style: { name: 'TableStyleMedium2' },
            },
          ]}
          onCreateFromSelection={onCreateFromSelection}
          onDeleteTable={onDeleteTable}
          onResizeTable={onResizeTable}
          onSelectRange={onSelectRange}
        />,
      )
    })

    expect(host.querySelector("[data-testid='workbook-tables-panel']")).not.toBeNull()
    expect(host.textContent).toContain('Sales')
    expect(host.textContent).toContain('TableStyleMedium2')

    await act(async () => {
      host.querySelector('button')?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(onCreateFromSelection).toHaveBeenCalledTimes(1)

    const buttons = [...host.querySelectorAll('button')]
    await act(async () => {
      buttons.find((button) => button.textContent?.includes('Resize'))?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      buttons.find((button) => button.textContent?.includes('Delete'))?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onResizeTable).toHaveBeenCalledWith('Sales', selectionRange)
    expect(onDeleteTable).toHaveBeenCalledWith('Sales')
  })
})
