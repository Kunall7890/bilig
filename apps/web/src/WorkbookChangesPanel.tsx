import { Button } from '@base-ui/react/button'
import { ListChecks, Undo2 } from 'lucide-react'
import { cn } from './cn.js'
import { agentPanelLabelTextClass, agentPanelMetaTextClass } from './workbook-agent-panel-primitives.js'
import { formatWorkbookChangeDay, formatWorkbookChangeTime, type WorkbookChangeEntry } from './workbook-changes-model.js'

interface WorkbookChangeDaySection {
  readonly dayLabel: string
  readonly changes: readonly WorkbookChangeEntry[]
}

function groupChangesByDay(changes: readonly WorkbookChangeEntry[]): readonly WorkbookChangeDaySection[] {
  const sections: Array<{ dayLabel: string; changes: WorkbookChangeEntry[] }> = []
  changes.forEach((change) => {
    const dayLabel = formatWorkbookChangeDay(change.createdAt)
    const lastSection = sections.at(-1)
    if (lastSection && lastSection.dayLabel === dayLabel) {
      lastSection.changes.push(change)
      return
    }
    sections.push({
      dayLabel,
      changes: [change],
    })
  })
  return sections
}

function renderChangeStatus(change: WorkbookChangeEntry): string | null {
  if (change.revertedByRevision !== null) {
    return `reverted by r${change.revertedByRevision}`
  }
  if (change.revertsRevision !== null) {
    return `reverts r${change.revertsRevision}`
  }
  return null
}

function WorkbookChangeRow(props: {
  readonly change: WorkbookChangeEntry
  readonly isRevertPending: boolean
  readonly onJump: (sheetName: string, address: string) => void
  readonly onRevert?: (revision: number) => void
}) {
  const { change } = props
  const statusLabel = renderChangeStatus(change)
  const canRevert = change.canRevert && props.onRevert !== undefined
  const metadata = [change.actorLabel, formatWorkbookChangeTime(change.createdAt), statusLabel].filter(
    (value): value is string => typeof value === 'string' && value.length > 0,
  )

  const content = (
    <div className="min-w-0 flex-1 px-3 py-2.5">
      <div className="min-w-0 text-[12px] font-medium leading-5 text-[var(--wb-text)]">{change.summary}</div>
      <div
        className="mt-0.5 flex flex-wrap items-center gap-x-1.5 gap-y-0.5 text-[11px] leading-4 text-[var(--wb-text-subtle)]"
        data-testid="workbook-change-meta"
      >
        {metadata.map((item, index) => (
          <span key={`${change.revision}:${item}`} className={cn(index === 0 ? 'text-[var(--wb-text-muted)]' : '')}>
            {index > 0 ? (
              <span aria-hidden="true" className="mr-1.5 text-[var(--wb-text-subtle)]">
                •
              </span>
            ) : null}
            {item}
          </span>
        ))}
      </div>
    </div>
  )
  const revertButton = canRevert ? (
    <Button
      aria-label={`Revert ${change.summary}`}
      className="mr-2 inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-[var(--wb-radius-control)] border border-transparent text-[var(--wb-text-muted)] transition-colors hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] disabled:cursor-not-allowed disabled:opacity-50"
      data-testid="workbook-change-revert"
      disabled={props.isRevertPending}
      title={`Revert ${change.summary}`}
      type="button"
      onClick={() => {
        props.onRevert?.(change.revision)
      }}
    >
      <Undo2 aria-hidden="true" className="h-3.5 w-3.5" strokeWidth={1.9} />
    </Button>
  ) : null

  return (
    <li className="border-b border-[var(--wb-border)] last:border-b-0" data-testid="workbook-change-row">
      {change.isJumpable ? (
        <div className="flex min-w-0 items-center transition-colors hover:bg-[var(--wb-hover)]">
          <Button
            className="block min-w-0 flex-1 text-left focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-inset"
            type="button"
            onClick={() => {
              if (change.sheetName && change.address) {
                props.onJump(change.sheetName, change.address)
              }
            }}
          >
            {content}
          </Button>
          {revertButton}
        </div>
      ) : (
        <div className="flex min-w-0 items-center">
          {content}
          {revertButton}
        </div>
      )}
    </li>
  )
}

export function WorkbookChangesPanel(props: {
  readonly changes: readonly WorkbookChangeEntry[]
  readonly pendingRevertRevision?: number | null
  readonly onJump: (sheetName: string, address: string) => void
  readonly onRevert?: (revision: number) => void
}) {
  const sections = groupChangesByDay(props.changes)
  const isEmpty = sections.length === 0

  return (
    <div
      aria-label="Workbook changes"
      className="flex h-full min-h-0 flex-col"
      data-testid="workbook-changes-panel"
      id="workbook-changes-panel"
    >
      {isEmpty ? (
        <div className="flex min-h-0 flex-1 items-center justify-center px-5 py-6" data-testid="workbook-changes-empty-state">
          <div className="flex max-w-52 flex-col items-center text-center">
            <div className="mb-3 flex h-9 w-9 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)]">
              <ListChecks aria-hidden="true" className="h-4 w-4" />
            </div>
            <div className={agentPanelLabelTextClass()}>No changes yet</div>
            <div className={cn(agentPanelMetaTextClass(), 'mt-1')}>Workbook is up to date.</div>
          </div>
        </div>
      ) : (
        <div className="min-h-0 flex-1 overflow-y-auto px-0 py-2">
          {sections.map((section) => (
            <section key={section.dayLabel} aria-label={section.dayLabel}>
              <div className="px-3 pb-1 pt-3 text-[10px] font-semibold uppercase tracking-[0.08em] text-[var(--wb-text-subtle)]">
                {section.dayLabel}
              </div>
              <ol className="m-0 list-none p-0">
                {section.changes.map((change) => (
                  <WorkbookChangeRow
                    key={`${change.revision}:${change.summary}`}
                    change={change}
                    isRevertPending={props.pendingRevertRevision === change.revision}
                    onJump={props.onJump}
                    {...(props.onRevert ? { onRevert: props.onRevert } : {})}
                  />
                ))}
              </ol>
            </section>
          ))}
        </div>
      )}
    </div>
  )
}
