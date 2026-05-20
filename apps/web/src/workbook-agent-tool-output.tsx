import { WORKBOOK_AGENT_TOOL_NAMES, normalizeWorkbookAgentToolName } from '@bilig/agent-api'
import { cn } from './cn.js'
import { agentPanelEyebrowTextClass, agentPanelLabelTextClass, agentPanelMetaTextClass } from './workbook-agent-panel-primitives.js'
import {
  formatCount,
  formatMissingCheckLabel,
  formatRangeLabel,
  formatScalarValue,
  formatStatusValue,
  formulaAuditLabel,
  humanizeKey,
  invariantAuditLabel,
  isRangeRecord,
  isRecord,
  joinHumanList,
  mutationReceiptRecord,
  mutationReceiptStatus,
  proofMatchedLabel,
  readNumber,
  readString,
  readStringItems,
  renderedProofLabel,
  renderReasonLabel,
  safeParseToolOutput,
  undoProofLabel,
  verificationReportStatus,
} from './workbook-agent-tool-output-model.js'
import { workbookInsetClass, workbookPillClass, workbookSurfaceClass } from './workbook-shell-chrome.js'

export { renderToolDisplayName, safeParseToolOutput, summarizeToolEntry } from './workbook-agent-tool-output-model.js'

function ToolFactGrid(props: { readonly facts: ReadonlyArray<{ label: string; value: string }> }) {
  if (props.facts.length === 0) {
    return null
  }
  return (
    <div className="grid gap-x-3 gap-y-2 sm:grid-cols-2">
      {props.facts.map((fact) => (
        <div key={fact.label} className="min-w-0">
          <div className={agentPanelEyebrowTextClass()}>{fact.label}</div>
          <div className={cn(agentPanelLabelTextClass(), 'mt-0.5 break-words font-semibold')}>{fact.value}</div>
        </div>
      ))}
    </div>
  )
}

function ToolKeyValueCard(props: {
  readonly title: string
  readonly subtitle?: string | null
  readonly facts?: ReadonlyArray<{ label: string; value: string }>
  readonly pills?: readonly string[]
}) {
  return (
    <div className={cn(workbookSurfaceClass({ emphasis: 'raised' }), 'px-3 py-2.5')}>
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className={cn(agentPanelLabelTextClass(), 'font-semibold')}>{props.title}</div>
          {props.subtitle ? <div className={cn(agentPanelMetaTextClass(), 'mt-0.5 break-words')}>{props.subtitle}</div> : null}
        </div>
      </div>
      {props.facts?.length ? (
        <div className="mt-2">
          <ToolFactGrid facts={props.facts} />
        </div>
      ) : null}
      {props.pills?.length ? (
        <div className="mt-2 flex flex-wrap gap-1.5">
          {props.pills.map((pill) => (
            <span key={pill} className={workbookPillClass({ tone: 'neutral' })}>
              {pill}
            </span>
          ))}
        </div>
      ) : null}
    </div>
  )
}

function GenericParsedObjectOutput(props: { readonly parsed: unknown }) {
  if (Array.isArray(props.parsed)) {
    return (
      <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
        <div className={cn(agentPanelMetaTextClass(), 'mb-2')}>{formatCount(props.parsed.length, 'result')}</div>
        <div className="flex flex-col gap-2">
          {props.parsed.slice(0, 8).map((item, index) => {
            const itemKey = isRecord(item)
              ? `${readString(item['name'])}:${readString(item['title'])}:${readString(item['sheetName'])}:${readString(item['address'])}:${index}`
              : `${String(item)}:${index}`
            if (!isRecord(item)) {
              return <ToolKeyValueCard key={itemKey} title={`Item ${index + 1}`} subtitle={formatScalarValue(item) ?? String(item)} />
            }
            return <GenericRecordCard key={itemKey} record={item} />
          })}
        </div>
      </div>
    )
  }
  if (isRecord(props.parsed)) {
    return (
      <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
        <GenericRecordCard record={props.parsed} />
      </div>
    )
  }
  return null
}

function GenericRecordCard(props: { readonly record: Record<string, unknown> }) {
  const title =
    readString(props.record['name']) ||
    readString(props.record['title']) ||
    readString(props.record['sheetName']) ||
    readString(props.record['address']) ||
    'Result'
  const subtitle = isRangeRecord(props.record)
    ? formatRangeLabel(props.record)
    : typeof props.record['sheetName'] === 'string' && typeof props.record['address'] === 'string'
      ? `${readString(props.record['sheetName'])}!${readString(props.record['address'])}`
      : null

  const facts = Object.entries(props.record)
    .filter(
      ([key]) => !['name', 'title', 'sheetName', 'address', 'startAddress', 'endAddress', 'columnNames', 'tables', 'sheets'].includes(key),
    )
    .flatMap(([key, value]) => {
      if (isRangeRecord(value)) {
        return [
          {
            label: humanizeKey(key),
            value: formatRangeLabel(value),
          },
        ]
      }
      const scalar = formatScalarValue(value)
      if (scalar !== null) {
        return [
          {
            label: humanizeKey(key),
            value: scalar,
          },
        ]
      }
      if (Array.isArray(value)) {
        return [
          {
            label: humanizeKey(key),
            value: formatCount(value.length, 'item'),
          },
        ]
      }
      return []
    })

  const pills = Array.isArray(props.record['columnNames'])
    ? props.record['columnNames'].flatMap((value) => (typeof value === 'string' ? [value] : [])).slice(0, 8)
    : []

  return <ToolKeyValueCard facts={facts} pills={pills} subtitle={subtitle} title={title} />
}

function renderListTablesOutput(parsed: Record<string, unknown>) {
  const tables = Array.isArray(parsed['tables']) ? parsed['tables'].flatMap((table) => (isRecord(table) ? [table] : [])) : []
  const tableCount = readNumber(parsed['tableCount'], tables.length)
  const documentId = readString(parsed['documentId'])
  return (
    <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
      <div className={cn(agentPanelMetaTextClass(), 'flex items-start justify-between gap-3')}>
        <div>{formatCount(tableCount, 'table')}</div>
        {documentId ? <div className={agentPanelEyebrowTextClass()}>{documentId}</div> : null}
      </div>
      <div className="mt-2 flex flex-col gap-2">
        {tables.map((table, index) => {
          const title = readString(table['name'], `Table ${index + 1}`)
          const sheetName = readString(table['sheetName'])
          const startAddress = readString(table['startAddress'])
          const endAddress = readString(table['endAddress'])
          const subtitle = sheetName && startAddress && endAddress ? formatRangeLabel({ sheetName, startAddress, endAddress }) : null
          const facts = [
            {
              label: 'Rows',
              value: formatCount(readNumber(table['rowCount']), 'row'),
            },
            {
              label: 'Columns',
              value: formatCount(readNumber(table['columnCount']), 'column'),
            },
            {
              label: 'Header rows',
              value: formatCount(readNumber(table['headerRowCount']), 'row'),
            },
          ]
          const columnNames = Array.isArray(table['columnNames'])
            ? table['columnNames'].flatMap((columnName) => (typeof columnName === 'string' ? [columnName] : [])).slice(0, 8)
            : []
          return <ToolKeyValueCard facts={facts} key={title} pills={columnNames} subtitle={subtitle} title={title} />
        })}
      </div>
    </div>
  )
}

function renderGetContextOutput(parsed: Record<string, unknown>) {
  const selection = isRecord(parsed['selection']) ? parsed['selection'] : null
  const visibleRange = isRangeRecord(parsed['visibleRange']) ? parsed['visibleRange'] : null
  const selectionFacts: Array<{ label: string; value: string }> = []
  if (selection && typeof selection['sheetName'] === 'string' && typeof selection['address'] === 'string') {
    selectionFacts.push({
      label: 'Active cell',
      value: `${readString(selection['sheetName'])}!${readString(selection['address'])}`,
    })
    const selectionRange = isRecord(selection['range']) ? selection['range'] : null
    if (selectionRange && typeof selectionRange['startAddress'] === 'string' && typeof selectionRange['endAddress'] === 'string') {
      selectionFacts.push({
        label: 'Selection',
        value: formatRangeLabel({
          sheetName: readString(selection['sheetName']),
          startAddress: readString(selectionRange['startAddress']),
          endAddress: readString(selectionRange['endAddress']),
        }),
      })
    }
  }

  const viewport = isRecord(parsed['viewport']) ? parsed['viewport'] : null
  const viewportFacts =
    viewport && typeof viewport['rowStart'] === 'number'
      ? [
          {
            label: 'Viewport rows',
            value: `${readNumber(viewport['rowStart']) + 1}-${readNumber(viewport['rowEnd']) + 1}`,
          },
          {
            label: 'Viewport columns',
            value: `${readNumber(viewport['colStart']) + 1}-${readNumber(viewport['colEnd']) + 1}`,
          },
        ]
      : []

  return (
    <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
      <div className="grid gap-2">
        <ToolKeyValueCard facts={selectionFacts} subtitle={visibleRange ? formatRangeLabel(visibleRange) : null} title="Workbook context" />
        {viewportFacts.length > 0 ? <ToolKeyValueCard facts={viewportFacts} title="Viewport" /> : null}
      </div>
    </div>
  )
}

function renderMutationReceiptOutput(parsed: Record<string, unknown>) {
  const receipt = mutationReceiptRecord(parsed)
  if (!receipt) {
    return null
  }
  const status = mutationReceiptStatus(parsed)
  const revision = isRecord(receipt['revision']) ? receipt['revision'] : null
  const authoritativeReadback = isRecord(receipt['authoritativeReadback']) ? receipt['authoritativeReadback'] : null
  const renderedReadback = isRecord(receipt['renderedReadback']) ? receipt['renderedReadback'] : null
  const undo = isRecord(receipt['undo']) ? receipt['undo'] : null
  const warnings = Array.isArray(receipt['warnings'])
    ? receipt['warnings'].flatMap((warning) => (typeof warning === 'string' && warning.trim().length > 0 ? [warning] : []))
    : []
  const facts = [
    {
      label: 'Status',
      value: formatStatusValue(status),
    },
    {
      label: 'Revision',
      value:
        typeof revision?.['after'] === 'number'
          ? `r${String(revision['after'])}`
          : typeof parsed['revision'] === 'number'
            ? `r${String(parsed['revision'])}`
            : 'Unknown',
    },
    {
      label: 'Authoritative proof',
      value: proofMatchedLabel(authoritativeReadback),
    },
    {
      label: 'Rendered proof',
      value: proofMatchedLabel(renderedReadback),
    },
    {
      label: 'Undo proof',
      value: undoProofLabel(undo),
    },
  ]
  const proofReason =
    [
      typeof renderedReadback?.['incompleteReason'] === 'string' ? renderedReadback['incompleteReason'] : null,
      typeof undo?.['reasonUnavailable'] === 'string' ? undo['reasonUnavailable'] : null,
    ].find((reason): reason is string => reason !== null) ?? null

  return (
    <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
      <ToolKeyValueCard
        facts={facts}
        pills={warnings.length > 0 ? ['Needs verification'] : []}
        subtitle={proofReason}
        title={status === 'verification_incomplete' ? 'Verification incomplete' : 'Mutation receipt'}
      />
      {warnings.length > 0 ? (
        <div className="mt-2 flex flex-col gap-1">
          {warnings.map((warning) => (
            <div
              key={warning}
              className={cn(
                agentPanelMetaTextClass(),
                'rounded-[var(--wb-radius-control)] border border-[#ead0d0] bg-[#fff8f8] px-2 py-1.5',
              )}
            >
              {warning}
            </div>
          ))}
        </div>
      ) : null}
    </div>
  )
}

function renderVerificationReportOutput(parsed: Record<string, unknown>) {
  const status = verificationReportStatus(parsed)
  if (!status) {
    return null
  }
  const missingChecks = readStringItems(parsed['verificationMissingChecks'])
  const revision = typeof parsed['appliedRevision'] === 'number' ? `r${String(parsed['appliedRevision'])}` : 'Unknown'
  const facts = [
    {
      label: 'Status',
      value: formatStatusValue(status),
    },
    {
      label: 'Revision',
      value: revision,
    },
    {
      label: 'Rendered proof',
      value: renderedProofLabel(parsed['renderedReadback']),
    },
    {
      label: 'Formula audit',
      value: formulaAuditLabel(parsed['formulaIssues']),
    },
    {
      label: 'Invariant audit',
      value: invariantAuditLabel(parsed['invariants']),
    },
  ]
  const missingLabels = missingChecks.map(formatMissingCheckLabel)
  return (
    <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
      <ToolKeyValueCard
        facts={facts}
        pills={missingLabels.length > 0 ? ['Missing checks', ...missingLabels] : []}
        subtitle={missingLabels.length > 0 ? `Missing checks: ${joinHumanList(missingLabels)}` : null}
        title={status === 'verification_incomplete' ? 'Verification incomplete' : 'Verification report'}
      />
    </div>
  )
}

function renderToolSpecificOutput(normalizedToolName: string | null, parsed: unknown) {
  if (!normalizedToolName || !isRecord(parsed)) {
    return null
  }

  const mutationReceiptOutput = renderMutationReceiptOutput(parsed)
  if (mutationReceiptOutput) {
    return mutationReceiptOutput
  }

  const verificationReportOutput = renderVerificationReportOutput(parsed)
  if (verificationReportOutput) {
    return verificationReportOutput
  }

  if (normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.listTables && Array.isArray(parsed['tables'])) {
    return renderListTablesOutput(parsed)
  }

  if (normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.getContext) {
    return renderGetContextOutput(parsed)
  }

  return null
}

export function StructuredToolOutput(props: { readonly toolName: string | null; readonly outputText: string | null }) {
  const parsed = safeParseToolOutput(props.outputText)
  const normalizedToolName = props.toolName ? normalizeWorkbookAgentToolName(props.toolName) : null
  if (!normalizedToolName || parsed === null) {
    return null
  }

  const toolSpecificOutput = renderToolSpecificOutput(normalizedToolName, parsed)
  if (toolSpecificOutput) {
    return toolSpecificOutput
  }

  if (normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues && isRecord(parsed) && Array.isArray(parsed['issues'])) {
    const summary = isRecord(parsed['summary']) ? parsed['summary'] : null
    const issues = parsed['issues'].flatMap((issue) => (isRecord(issue) ? [issue] : []))
    return (
      <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
        <div className={cn(agentPanelMetaTextClass(), 'flex items-start justify-between gap-3')}>
          <div>
            {readNumber(summary?.['issueCount'], issues.length)} issues · {readNumber(summary?.['scannedFormulaCells'])} formulas
          </div>
          <div className={cn(agentPanelEyebrowTextClass(), 'text-right')}>
            {readNumber(summary?.['errorCount'])} errors · {readNumber(summary?.['cycleCount'])} cycles ·{' '}
            {readNumber(summary?.['unsupportedCount'])} JS-only
          </div>
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {issues.slice(0, 8).map((issue) => (
            <div
              key={`${readString(issue['sheetName'])}:${readString(issue['address'])}`}
              className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2"
            >
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <div className={cn(agentPanelLabelTextClass(), 'font-semibold')}>
                    {readString(issue['sheetName'])}!{readString(issue['address'])}
                  </div>
                  <div className={cn(agentPanelMetaTextClass(), 'mt-1 break-all')}>{readString(issue['formula'])}</div>
                </div>
                <div className={cn(agentPanelEyebrowTextClass(), 'text-right')}>{readString(issue['valueText']) || '(empty)'}</div>
              </div>
              <div className="mt-2 flex flex-wrap gap-1.5">
                {Array.isArray(issue['issueKinds'])
                  ? issue['issueKinds'].map((kind) => (
                      <span key={readString(kind)} className={workbookPillClass({ tone: 'danger', weight: 'strong' })}>
                        {readString(kind)}
                      </span>
                    ))
                  : null}
              </div>
            </div>
          ))}
        </div>
      </div>
    )
  }

  if (normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook && isRecord(parsed) && Array.isArray(parsed['matches'])) {
    const matches = parsed['matches'].flatMap((match) => (isRecord(match) ? [match] : []))
    return (
      <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
        <div className={cn(agentPanelMetaTextClass(), 'flex items-start justify-between gap-3')}>
          <div className="truncate">“{readString(parsed['query'])}”</div>
          <div className={agentPanelEyebrowTextClass()}>
            {readNumber(isRecord(parsed['summary']) ? parsed['summary']['matchCount'] : undefined, matches.length)} matches
          </div>
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {matches.slice(0, 8).map((match) => (
            <div
              key={`${readString(match['kind'])}:${readString(match['sheetName'])}:${readString(match['address'])}`}
              className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2"
            >
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <div className={cn(agentPanelLabelTextClass(), 'font-semibold')}>
                    {readString(match['kind']) === 'sheet'
                      ? `Sheet ${readString(match['sheetName'])}`
                      : `${readString(match['sheetName'])}!${readString(match['address'])}`}
                  </div>
                  <div className={cn(agentPanelMetaTextClass(), 'mt-1 break-all')}>{readString(match['snippet'])}</div>
                </div>
                <div className={agentPanelEyebrowTextClass()}>score {readNumber(match['score'])}</div>
              </div>
              {Array.isArray(match['reasons']) ? (
                <div className="mt-2 flex flex-wrap gap-1.5">
                  {match['reasons'].map((reason) => (
                    <span key={readString(reason)} className={workbookPillClass({ tone: 'neutral' })}>
                      {renderReasonLabel(readString(reason))}
                    </span>
                  ))}
                </div>
              ) : null}
            </div>
          ))}
        </div>
      </div>
    )
  }

  if (normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.traceDependencies && isRecord(parsed) && Array.isArray(parsed['layers'])) {
    const root = isRecord(parsed['root']) ? parsed['root'] : null
    const layers = parsed['layers'].flatMap((layer) => (isRecord(layer) ? [layer] : []))
    return (
      <div className={cn(workbookInsetClass(), 'mt-2 px-3 py-3')}>
        <div className={cn(agentPanelMetaTextClass(), 'flex items-start justify-between gap-3')}>
          <div>
            {readString(root?.['sheetName'])}!{readString(root?.['address'])}
          </div>
          <div className={agentPanelEyebrowTextClass()}>
            {readString(parsed['direction'], 'both')} · {readNumber(parsed['depth'])} hops
          </div>
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {layers.map((layer) => (
            <div
              key={`trace-layer-${readNumber(layer['depth'])}`}
              className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2"
            >
              <div className={agentPanelEyebrowTextClass()}>Hop {readNumber(layer['depth'])}</div>
              <div className="mt-2 grid gap-2 md:grid-cols-2">
                <div>
                  <div className={cn(agentPanelLabelTextClass(), 'font-semibold')}>Precedents</div>
                  <div className="mt-1 flex flex-col gap-1">
                    {Array.isArray(layer['precedents']) && layer['precedents'].length > 0 ? (
                      layer['precedents']
                        .flatMap((node) => (isRecord(node) ? [node] : []))
                        .map((node) => (
                          <div
                            key={`precedent:${readString(node['sheetName'])}:${readString(node['address'])}`}
                            className={agentPanelMetaTextClass()}
                          >
                            {readString(node['sheetName'])}!{readString(node['address'])}{' '}
                            <span className="text-[var(--wb-text-muted)]">
                              {readString(node['formula']) || readString(node['valueText'])}
                            </span>
                          </div>
                        ))
                    ) : (
                      <div className={agentPanelMetaTextClass()}>None</div>
                    )}
                  </div>
                </div>
                <div>
                  <div className={cn(agentPanelLabelTextClass(), 'font-semibold')}>Dependents</div>
                  <div className="mt-1 flex flex-col gap-1">
                    {Array.isArray(layer['dependents']) && layer['dependents'].length > 0 ? (
                      layer['dependents']
                        .flatMap((node) => (isRecord(node) ? [node] : []))
                        .map((node) => (
                          <div
                            key={`dependent:${readString(node['sheetName'])}:${readString(node['address'])}`}
                            className={agentPanelMetaTextClass()}
                          >
                            {readString(node['sheetName'])}!{readString(node['address'])}{' '}
                            <span className="text-[var(--wb-text-muted)]">
                              {readString(node['formula']) || readString(node['valueText'])}
                            </span>
                          </div>
                        ))
                    ) : (
                      <div className={agentPanelMetaTextClass()}>None</div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    )
  }

  return <GenericParsedObjectOutput parsed={parsed} />
}
