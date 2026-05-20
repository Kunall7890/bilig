import type { WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { cn } from './cn.js'
import { agentPanelBodyTextClass } from './workbook-agent-panel-primitives.js'
import { TimelineCitationList } from './workbook-agent-panel-timeline-citations.js'
import { WorkbookAgentMarkdown } from './workbook-agent-markdown.js'
import { MessageCopyButton } from './workbook-agent-message-copy-button.js'

export function WorkbookAgentUserMessageRow(props: { readonly entry: WorkbookAgentTimelineEntry }) {
  const text = props.entry.text ?? ''
  return (
    <div className="flex min-w-0 max-w-full flex-col items-end px-3 py-3">
      <div
        className={cn(
          agentPanelBodyTextClass(),
          'min-w-0 max-w-[82%] overflow-hidden break-words rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-muted)] px-3 py-2 shadow-[0_1px_2px_rgba(15,23,42,0.04)]',
        )}
        data-testid={`workbook-agent-user-message-card-${props.entry.id}`}
      >
        <div className="min-w-0 max-w-full overflow-hidden">
          <WorkbookAgentMarkdown markdown={text} />
          <TimelineCitationList citations={props.entry.citations} />
        </div>
      </div>
      <div className="mt-1 flex max-w-[82%] justify-end" data-testid={`workbook-agent-user-message-actions-${props.entry.id}`}>
        <MessageCopyButton entryId={props.entry.id} messageKind="user" text={text} />
      </div>
    </div>
  )
}

export function WorkbookAgentAssistantMessageRow(props: { readonly entry: WorkbookAgentTimelineEntry }) {
  if (props.entry.phase === 'progress') {
    return null
  }
  if (!props.entry.text?.trim().length) {
    return null
  }

  return (
    <div className={cn(agentPanelBodyTextClass(), 'min-w-0 w-full max-w-full overflow-hidden px-3 py-3')}>
      <div className="min-w-0 max-w-full overflow-hidden" data-testid={`workbook-agent-assistant-message-body-${props.entry.id}`}>
        <WorkbookAgentMarkdown markdown={props.entry.text} />
        <TimelineCitationList citations={props.entry.citations} />
      </div>
      <div className="mt-1 flex justify-end" data-testid={`workbook-agent-assistant-message-actions-${props.entry.id}`}>
        <MessageCopyButton entryId={props.entry.id} messageKind="assistant" text={props.entry.text} />
      </div>
    </div>
  )
}
