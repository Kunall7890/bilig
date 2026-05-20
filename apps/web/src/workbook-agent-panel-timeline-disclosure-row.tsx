import type { WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { cn } from './cn.js'
import { WorkbookAgentDisclosureRow } from './workbook-agent-panel-disclosure-row.js'
import { agentPanelBodyMutedTextClass } from './workbook-agent-panel-primitives.js'
import { TimelineCitationList } from './workbook-agent-panel-timeline-citations.js'
import { summarizeDisclosureText } from './workbook-agent-panel-timeline-model.js'
import { WorkbookAgentMarkdown } from './workbook-agent-markdown.js'

export function WorkbookAgentTextDisclosureEntryRow(props: {
  readonly entry: WorkbookAgentTimelineEntry
  readonly label: 'Thought' | 'Plan'
}) {
  const bodyText = props.entry.kind === 'reasoning' || props.entry.kind === 'plan' ? props.entry.text : null
  if (!bodyText?.trim().length) {
    return null
  }
  const summary = summarizeDisclosureText(bodyText)
  const disclosureKey = props.entry.kind

  return (
    <WorkbookAgentDisclosureRow
      id={props.entry.id}
      label={props.label}
      panelTestId={`workbook-agent-${disclosureKey}-panel-${props.entry.id}`}
      summary={summary ?? `No ${props.label.toLowerCase()} details available.`}
      triggerLabel={{
        expanded: `Collapse ${props.label.toLowerCase()}`,
        collapsed: `Expand ${props.label.toLowerCase()}`,
      }}
      triggerTestId={`workbook-agent-${disclosureKey}-toggle-${props.entry.id}`}
    >
      <div className={cn(agentPanelBodyMutedTextClass(), 'min-w-0 max-w-full overflow-hidden')}>
        <WorkbookAgentMarkdown markdown={bodyText} />
      </div>
      <TimelineCitationList citations={props.entry.citations} />
    </WorkbookAgentDisclosureRow>
  )
}
