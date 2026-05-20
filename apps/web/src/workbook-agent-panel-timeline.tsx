import type { WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { cn } from './cn.js'
import { agentPanelMetaTextClass } from './workbook-agent-panel-primitives.js'
import { TimelineCitationList } from './workbook-agent-panel-timeline-citations.js'
import { WorkbookAgentTextDisclosureEntryRow } from './workbook-agent-panel-timeline-disclosure-row.js'
import { WorkbookAgentAssistantMessageRow, WorkbookAgentUserMessageRow } from './workbook-agent-panel-timeline-message-row.js'
import { isHiddenTimelineEntry } from './workbook-agent-panel-timeline-model.js'
import { WorkbookAgentToolEntryRow } from './workbook-agent-panel-timeline-tool-row.js'

export {
  getVisibleWorkbookAgentTimelineEntries,
  getWorkbookAgentProgressAnchorIndex,
  isHiddenTimelineEntry,
  summarizeDisclosureText,
  summarizeTimelineCitations,
} from './workbook-agent-panel-timeline-model.js'

export function WorkbookAgentEntryRow(props: { readonly entry: WorkbookAgentTimelineEntry }) {
  const { entry } = props
  if (entry.kind === 'user') {
    return <WorkbookAgentUserMessageRow entry={entry} />
  }

  if (entry.kind === 'assistant') {
    return <WorkbookAgentAssistantMessageRow entry={entry} />
  }

  if (entry.kind === 'reasoning') {
    return <WorkbookAgentTextDisclosureEntryRow entry={entry} label="Thought" />
  }

  if (entry.kind === 'plan') {
    return <WorkbookAgentTextDisclosureEntryRow entry={entry} label="Plan" />
  }

  if (entry.kind === 'tool') {
    return <WorkbookAgentToolEntryRow entry={entry} />
  }

  if (isHiddenTimelineEntry(entry)) {
    return null
  }

  return (
    <div className="px-3 py-2.5">
      <div className={cn(agentPanelMetaTextClass(), 'break-words')}>{entry.text}</div>
      <TimelineCitationList citations={entry.citations} />
    </div>
  )
}
