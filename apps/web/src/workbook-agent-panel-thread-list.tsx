import { Button } from '@base-ui/react/button'
import type { WorkbookAgentThreadSummary } from '@bilig/contracts'
import { cn } from './cn.js'
import {
  agentPanelLabelTextClass,
  agentPanelMetaTextClass,
  agentPanelThreadButtonClass,
  agentPanelThreadListClass,
} from './workbook-agent-panel-primitives.js'
import { formatWorkbookAgentThreadEntryCount, summarizeWorkbookAgentThreadActivity } from './workbook-agent-thread-summary.js'
import { formatWorkbookCollaboratorLabel } from './workbook-presence-model.js'
import { workbookPillClass } from './workbook-shell-chrome.js'

export function getVisibleWorkbookAgentThreadSummaries(input: {
  readonly activeThreadId: string | null
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[]
}): readonly WorkbookAgentThreadSummary[] {
  return input.threadSummaries.filter((threadSummary) => threadSummary.threadId !== input.activeThreadId)
}

export function ThreadSummaryStrip(props: {
  readonly activeThreadId: string | null
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[]
  readonly onSelectThread: (threadId: string) => void
}) {
  const visibleThreadSummaries = getVisibleWorkbookAgentThreadSummaries({
    activeThreadId: props.activeThreadId,
    threadSummaries: props.threadSummaries,
  })
  if (visibleThreadSummaries.length === 0) {
    return null
  }

  return (
    <div className={agentPanelThreadListClass()}>
      {visibleThreadSummaries.map((threadSummary) => {
        const latestActivity = summarizeWorkbookAgentThreadActivity(threadSummary.latestEntryText, 64)
        return (
          <Button
            key={threadSummary.threadId}
            aria-label={`Open ${threadSummary.scope} thread ${threadSummary.threadId}`}
            aria-pressed={false}
            className={agentPanelThreadButtonClass({ active: false })}
            data-testid={`workbook-agent-thread-${threadSummary.threadId}`}
            type="button"
            onClick={() => {
              props.onSelectThread(threadSummary.threadId)
            }}
          >
            <div className="min-w-0 flex-1">
              <div className="flex items-center gap-2">
                <span className={cn(agentPanelLabelTextClass(), 'font-semibold')}>
                  {threadSummary.scope === 'shared' ? 'Shared' : 'Private'}
                </span>
                <span className={agentPanelMetaTextClass()}>
                  {threadSummary.scope === 'shared' ? formatWorkbookCollaboratorLabel(threadSummary.ownerUserId) : 'Just you'}
                </span>
                <span className={agentPanelMetaTextClass()}>{formatWorkbookAgentThreadEntryCount(threadSummary.entryCount)}</span>
              </div>
              {latestActivity ? <div className={cn(agentPanelMetaTextClass(), 'mt-0.5 truncate')}>{latestActivity}</div> : null}
            </div>
            {threadSummary.scope === 'shared' && threadSummary.reviewQueueItemCount > 0 ? (
              <span className={workbookPillClass({ tone: 'accent', weight: 'strong' })}>Review</span>
            ) : null}
          </Button>
        )
      })}
    </div>
  )
}
