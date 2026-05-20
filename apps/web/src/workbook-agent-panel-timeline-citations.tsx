import { Fragment } from 'react'
import type { WorkbookAgentTimelineCitation } from '@bilig/contracts'
import { cn } from './cn.js'
import { agentPanelMetaTextClass } from './workbook-agent-panel-primitives.js'
import { summarizeTimelineCitations } from './workbook-agent-panel-timeline-model.js'

export function TimelineCitationList(props: { readonly citations: readonly WorkbookAgentTimelineCitation[] }) {
  const segments = summarizeTimelineCitations(props.citations)
  if (segments.length === 0) {
    return null
  }
  return (
    <div className={cn(agentPanelMetaTextClass(), 'mt-1 break-words')}>
      {segments.map((segment, index) => (
        <Fragment key={segment}>
          {index > 0 ? <span aria-hidden="true"> · </span> : null}
          <span>{segment}</span>
        </Fragment>
      ))}
    </div>
  )
}
