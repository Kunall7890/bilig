import { cva } from 'class-variance-authority'
import type { WorkbookAgentTimelineEntry } from '@bilig/contracts'
import { cn } from './cn.js'
import { WorkbookAgentDisclosureRow } from './workbook-agent-panel-disclosure-row.js'
import { agentPanelEyebrowTextClass, agentPanelLabelTextClass, agentPanelMetaTextClass } from './workbook-agent-panel-primitives.js'
import { TimelineCitationList } from './workbook-agent-panel-timeline-citations.js'
import { renderToolDisplayName, safeParseToolOutput, StructuredToolOutput, summarizeToolEntry } from './workbook-agent-tool-output.js'
import { workbookPillClass } from './workbook-shell-chrome.js'

const toolStatusPillClass = cva(
  'inline-flex h-5 items-center rounded-full border px-2 text-[10px] leading-none font-medium uppercase tracking-[0.05em]',
  {
    variants: {
      status: {
        completed: workbookPillClass({ tone: 'neutral', weight: 'strong' }),
        failed: workbookPillClass({ tone: 'danger', weight: 'strong' }),
        running: workbookPillClass({ tone: 'accent', weight: 'strong' }),
      },
    },
  },
)

export function WorkbookAgentToolEntryRow(props: { readonly entry: WorkbookAgentTimelineEntry }) {
  const displayName = renderToolDisplayName(props.entry.toolName)
  const summary = summarizeToolEntry(props.entry)
  const parsedOutput = safeParseToolOutput(props.entry.outputText)
  const isCommandExecution = props.entry.toolName === 'command_execution'
  const hasDetails = (props.entry.argumentsText?.trim().length ?? 0) > 0 || (props.entry.outputText?.trim().length ?? 0) > 0
  if (!hasDetails) {
    return (
      <div className="px-3 py-2.5">
        <div className="flex items-center justify-between gap-3">
          <div className="min-w-0 flex flex-1 items-center gap-2.5">
            <div className={cn(agentPanelLabelTextClass(), 'min-w-0')}>{displayName}</div>
            {summary ? (
              <div className={cn(agentPanelMetaTextClass(), 'min-w-0 flex-1 whitespace-normal break-words')}>{summary}</div>
            ) : null}
          </div>
          <ToolStatusPill status={props.entry.toolStatus} />
        </div>
        <TimelineCitationList citations={props.entry.citations} />
      </div>
    )
  }

  return (
    <WorkbookAgentDisclosureRow
      badge={<ToolStatusPill status={props.entry.toolStatus} />}
      id={props.entry.id}
      label={displayName}
      panelTestId={`workbook-agent-tool-panel-${props.entry.id}`}
      summary={summary}
      triggerLabel={{
        expanded: `Collapse ${displayName}`,
        collapsed: `Expand ${displayName}`,
      }}
      triggerTestId={`workbook-agent-tool-toggle-${props.entry.id}`}
    >
      {props.entry.argumentsText?.trim().length ? <ToolPreformattedBlock label="Arguments" text={props.entry.argumentsText} /> : null}
      {props.entry.outputText?.trim().length ? (
        <div className={props.entry.argumentsText?.trim().length ? 'mt-2' : undefined}>
          <div className={agentPanelEyebrowTextClass()}>Output</div>
          {parsedOutput !== null && !isCommandExecution ? (
            <StructuredToolOutput toolName={props.entry.toolName} outputText={props.entry.outputText} />
          ) : (
            <ToolPreformattedBody text={props.entry.outputText} />
          )}
        </div>
      ) : null}
      <TimelineCitationList citations={props.entry.citations} />
    </WorkbookAgentDisclosureRow>
  )
}

function ToolStatusPill(props: { readonly status: WorkbookAgentTimelineEntry['toolStatus'] }) {
  if (props.status === 'completed') {
    return null
  }
  const label = props.status === 'failed' ? 'Failed' : 'Running'
  return (
    <span
      className={toolStatusPillClass({
        status: props.status === 'failed' ? 'failed' : 'running',
      })}
    >
      {label}
    </span>
  )
}

function ToolPreformattedBlock(props: { readonly label: string; readonly text: string }) {
  return (
    <div>
      <div className={agentPanelEyebrowTextClass()}>{props.label}</div>
      <ToolPreformattedBody text={props.text} />
    </div>
  )
}

function ToolPreformattedBody(props: { readonly text: string }) {
  return (
    <pre
      className={cn(
        agentPanelMetaTextClass(),
        'mt-1 box-border w-full min-w-0 max-w-full overflow-x-hidden whitespace-pre-wrap break-all rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2',
      )}
    >
      {props.text}
    </pre>
  )
}
