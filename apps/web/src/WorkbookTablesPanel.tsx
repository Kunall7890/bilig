import { Button } from '@base-ui/react/button'
import { Maximize2, Plus, Trash2 } from 'lucide-react'
import { parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, WorkbookTableSnapshot } from '@bilig/protocol'
import { cn } from './cn.js'

function tableRange(table: WorkbookTableSnapshot): CellRangeRef {
  return {
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
  }
}

function rangeLabel(range: CellRangeRef): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

function rangeWidth(range: CellRangeRef): number {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return Math.abs(end.col - start.col) + 1
}

function tableCanResizeToSelection(table: WorkbookTableSnapshot, selectionRange: CellRangeRef): boolean {
  return table.sheetName === selectionRange.sheetName && rangeWidth(selectionRange) > 0
}

const panelButtonClass =
  'inline-flex h-8 items-center justify-center gap-1.5 rounded-[var(--wb-radius-control)] border border-[var(--color-mauve-300)] bg-white px-2.5 text-[12px] font-medium text-[var(--color-mauve-800)] shadow-sm transition-colors hover:bg-[var(--color-mauve-100)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] disabled:cursor-not-allowed disabled:opacity-50'

export function WorkbookTablesPanel(props: {
  readonly tables: readonly WorkbookTableSnapshot[]
  readonly selectionRange: CellRangeRef
  readonly activeTableName?: string | null
  readonly onCreateFromSelection: () => void
  readonly onDeleteTable: (name: string) => void
  readonly onResizeTable: (name: string, range: CellRangeRef) => void
  readonly onSelectRange: (range: CellRangeRef) => void
}) {
  return (
    <div className="flex h-full min-h-0 flex-col overflow-hidden bg-[var(--color-mauve-50)]" data-testid="workbook-tables-panel">
      <div className="border-b border-[var(--color-mauve-200)] px-3 py-2.5">
        <div className="flex items-center gap-2">
          <div className="min-w-0 flex-1">
            <div className="text-[13px] font-semibold text-[var(--color-mauve-950)]">Tables</div>
            <div className="mt-0.5 truncate text-[11px] text-[var(--color-mauve-600)]">{rangeLabel(props.selectionRange)}</div>
          </div>
          <Button className={cn(panelButtonClass, 'shrink-0')} type="button" onClick={props.onCreateFromSelection}>
            <Plus aria-hidden="true" className="size-3.5" strokeWidth={2} />
            Create
          </Button>
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.tables.length === 0 ? (
          <div className="rounded-lg border border-dashed border-[var(--color-mauve-300)] bg-white px-3 py-4 text-[12px] text-[var(--color-mauve-600)]">
            No workbook tables.
          </div>
        ) : (
          <div className="flex flex-col gap-2.5">
            {props.tables.map((table) => {
              const range = tableRange(table)
              const isActive = props.activeTableName === table.name
              const canResize = tableCanResizeToSelection(table, props.selectionRange)
              return (
                <section
                  className={cn(
                    'rounded-lg border bg-white px-3 py-3 shadow-sm',
                    isActive ? 'border-[var(--color-blue-500)]' : 'border-[var(--color-mauve-200)]',
                  )}
                  data-testid={`workbook-table-card-${table.name}`}
                  key={table.name}
                >
                  <div className="flex items-start gap-2">
                    <button
                      className="min-w-0 flex-1 text-left focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)]"
                      type="button"
                      onClick={() => props.onSelectRange(range)}
                    >
                      <div className="truncate text-[13px] font-semibold text-[var(--color-mauve-950)]">{table.name}</div>
                      <div className="mt-0.5 truncate text-[11px] text-[var(--color-mauve-600)]">{rangeLabel(range)}</div>
                    </button>
                  </div>
                  <dl className="mt-2 grid grid-cols-2 gap-x-3 gap-y-1.5 text-[11px]">
                    <div>
                      <dt className="text-[var(--color-mauve-500)]">Columns</dt>
                      <dd className="font-medium text-[var(--color-mauve-900)]">{String(table.columnNames.length)}</dd>
                    </div>
                    <div>
                      <dt className="text-[var(--color-mauve-500)]">Header</dt>
                      <dd className="font-medium text-[var(--color-mauve-900)]">{table.headerRow ? 'On' : 'Off'}</dd>
                    </div>
                    <div>
                      <dt className="text-[var(--color-mauve-500)]">Totals</dt>
                      <dd className="font-medium text-[var(--color-mauve-900)]">{table.totalsRow ? 'On' : 'Off'}</dd>
                    </div>
                    <div>
                      <dt className="text-[var(--color-mauve-500)]">Style</dt>
                      <dd className="truncate font-medium text-[var(--color-mauve-900)]">{table.style?.name ?? 'Default'}</dd>
                    </div>
                  </dl>
                  <div className="mt-2 flex flex-wrap gap-1.5">
                    <Button
                      className={panelButtonClass}
                      disabled={!canResize}
                      title={canResize ? 'Resize table to current selection' : 'Select a range on the same sheet to resize'}
                      type="button"
                      onClick={() => props.onResizeTable(table.name, props.selectionRange)}
                    >
                      <Maximize2 aria-hidden="true" className="size-3.5" strokeWidth={2} />
                      Resize
                    </Button>
                    <Button className={panelButtonClass} type="button" onClick={() => props.onDeleteTable(table.name)}>
                      <Trash2 aria-hidden="true" className="size-3.5" strokeWidth={2} />
                      Delete
                    </Button>
                  </div>
                  <div className="mt-2 truncate text-[11px] text-[var(--color-mauve-600)]">{table.columnNames.join(', ')}</div>
                </section>
              )
            })}
          </div>
        )}
      </div>
    </div>
  )
}
