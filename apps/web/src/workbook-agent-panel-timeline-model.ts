import type { WorkbookAgentTimelineCitation, WorkbookAgentTimelineEntry } from '@bilig/contracts'

export function summarizeTimelineCitations(citations: readonly WorkbookAgentTimelineCitation[]): readonly string[] {
  const seen = new Set<string>()
  const segments: string[] = []
  for (const citation of citations) {
    if (citation.kind === 'revision') {
      continue
    }
    const address =
      citation.startAddress === citation.endAddress
        ? `${citation.sheetName}!${citation.startAddress}`
        : `${citation.sheetName}!${citation.startAddress}:${citation.endAddress}`
    const segment = `${citation.role === 'target' ? 'Target' : 'Source'} ${address}`
    if (seen.has(segment)) {
      continue
    }
    seen.add(segment)
    segments.push(segment)
  }
  return segments
}

export function getVisibleWorkbookAgentTimelineEntries(input: {
  readonly optimisticEntries: readonly WorkbookAgentTimelineEntry[]
  readonly snapshotEntries: readonly WorkbookAgentTimelineEntry[]
}): readonly WorkbookAgentTimelineEntry[] {
  return [...input.optimisticEntries, ...input.snapshotEntries].filter((entry) => !isHiddenTimelineEntry(entry))
}

export function getWorkbookAgentProgressAnchorIndex(input: {
  readonly activeResponseTurnId: string | null
  readonly showAssistantProgress: boolean
  readonly visibleEntries: readonly WorkbookAgentTimelineEntry[]
}): number {
  if (!input.showAssistantProgress || !input.activeResponseTurnId) {
    return -1
  }
  return input.visibleEntries.findLastIndex((entry) => entry.turnId === input.activeResponseTurnId)
}

export function isHiddenTimelineEntry(entry: Pick<WorkbookAgentTimelineEntry, 'kind' | 'text'>): boolean {
  return isAppliedExecutionSystemEntry(entry) || isLegacyGenericCommandExecutionSystemEntry(entry)
}

export function summarizeDisclosureText(text: string | null): string | null {
  if (!text) {
    return null
  }
  const normalized = renderMarkdownPlainText(text).trim().replaceAll(/\s+/g, ' ')
  if (normalized.length === 0) {
    return null
  }
  return normalized.length <= 88 ? normalized : `${normalized.slice(0, 85)}...`
}

function renderMarkdownPlainText(markdown: string): string {
  return markdown
    .replaceAll(/```[\s\S]*?```/g, ' ')
    .replaceAll(/`([^`]+)`/g, '$1')
    .replaceAll(/\[([^\]]+)\]\(([^)]+)\)/g, '$1')
    .replaceAll(/!\[([^\]]*)\]\(([^)]+)\)/g, '$1')
    .replaceAll(/^#{1,6}\s+/gm, '')
    .replaceAll(/^>\s?/gm, '')
    .replaceAll(/[*_~]+/g, ' ')
    .replaceAll(/\s+/g, ' ')
}

function isAppliedExecutionSystemEntry(entry: Pick<WorkbookAgentTimelineEntry, 'kind' | 'text'>): boolean {
  return (
    entry.kind === 'system' &&
    (entry.text?.startsWith('Applied workbook change set at revision r') === true ||
      entry.text?.startsWith('Applied automatically workbook change set at revision r') === true ||
      entry.text?.startsWith('Applied automatically selected workbook change set at revision r') === true ||
      entry.text?.startsWith('Applied selected workbook change set at revision r') === true)
  )
}

function isLegacyGenericCommandExecutionSystemEntry(entry: Pick<WorkbookAgentTimelineEntry, 'kind' | 'text'>): boolean {
  return entry.kind === 'system' && entry.text === 'Codex emitted commandExecution.'
}
