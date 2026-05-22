import { normalizeWorkbookAgentToolName } from '@bilig/agent-api'
import type { WorkbookAgentTimelineEntry } from '@bilig/contracts'

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function readString(value: unknown, fallback = ''): string {
  return typeof value === 'string' ? value : fallback
}

export function readNumber(value: unknown, fallback = 0): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback
}

export function isRangeRecord(value: unknown): value is {
  sheetName: string
  startAddress: string
  endAddress: string
} {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  )
}

export function formatRangeLabel(input: { sheetName: string; startAddress: string; endAddress: string }): string {
  return `${input.sheetName}!${input.startAddress}${input.startAddress === input.endAddress ? '' : `:${input.endAddress}`}`
}

export function humanizeKey(key: string): string {
  return key
    .replaceAll(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replaceAll(/[_-]+/g, ' ')
    .replace(/\b\w/g, (segment) => segment.toUpperCase())
}

export function formatCount(count: number, singular: string, plural = `${singular}s`): string {
  return `${count} ${count === 1 ? singular : plural}`
}

export function formatScalarValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number') {
    return Number.isInteger(value) ? String(value) : value.toFixed(2)
  }
  if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No'
  }
  if (value === null) {
    return 'None'
  }
  return null
}

export function formatStatusValue(value: unknown): string {
  return typeof value === 'string' && value.length > 0 ? humanizeKey(value) : 'Unknown'
}

export function firstStringItem(value: unknown): string | null {
  return Array.isArray(value) ? (value.find((item): item is string => typeof item === 'string' && item.trim().length > 0) ?? null) : null
}

export function readStringItems(value: unknown): readonly string[] {
  return Array.isArray(value) ? value.flatMap((item) => (typeof item === 'string' && item.trim().length > 0 ? [item] : [])) : []
}

export function mutationReceiptRecord(parsed: Record<string, unknown>): Record<string, unknown> | null {
  const receipt = parsed['mutationReceipt']
  return isRecord(receipt) ? receipt : null
}

export function mutationReceiptStatus(parsed: Record<string, unknown>): string | null {
  const receipt = mutationReceiptRecord(parsed)
  const status = typeof receipt?.['status'] === 'string' ? receipt['status'] : parsed['status']
  return typeof status === 'string' ? status : null
}

export function mutationReceiptWarning(parsed: Record<string, unknown>): string | null {
  const receipt = mutationReceiptRecord(parsed)
  return firstStringItem(receipt?.['warnings']) ?? firstStringItem(parsed['warnings'])
}

export function summarizeMutationReceipt(parsed: Record<string, unknown>): string | null {
  if (!mutationReceiptRecord(parsed)) {
    return null
  }
  const status = mutationReceiptStatus(parsed)
  if (status !== 'verification_incomplete') {
    return null
  }
  const warning = mutationReceiptWarning(parsed)
  return warning ? `Verification incomplete: ${warning}` : 'Verification incomplete'
}

export function verificationReportStatus(parsed: Record<string, unknown>): string | null {
  if (typeof parsed['verificationComplete'] !== 'boolean') {
    return null
  }
  return typeof parsed['status'] === 'string' ? parsed['status'] : parsed['verificationComplete'] ? 'verified' : 'verification_incomplete'
}

export function formatMissingCheckLabel(check: string): string {
  switch (check) {
    case 'formulaIssues':
      return 'Formula Issues'
    case 'formulaIssuesClean':
      return 'Formula Audit Clean'
    case 'invariants':
      return 'Invariants'
    case 'invariantsClean':
      return 'Invariant Audit Clean'
    case 'renderedReadback':
      return 'Rendered Readback'
    case 'targetRange':
      return 'Target Range'
    default:
      return humanizeKey(check)
  }
}

export function joinHumanList(items: readonly string[]): string {
  if (items.length === 0) {
    return ''
  }
  if (items.length === 1) {
    return items[0]!
  }
  return `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]!}`
}

export function summarizeVerificationReport(parsed: Record<string, unknown>): string | null {
  const status = verificationReportStatus(parsed)
  if (status !== 'verification_incomplete') {
    return null
  }
  const missingChecks = readStringItems(parsed['verificationMissingChecks']).map((check) => formatMissingCheckLabel(check).toLowerCase())
  return missingChecks.length > 0 ? `Verification incomplete: missing ${joinHumanList(missingChecks)} checks` : 'Verification incomplete'
}

export function renderReasonLabel(reason: string): string {
  switch (reason) {
    case 'sheet':
      return 'sheet'
    case 'address':
      return 'address'
    case 'formula':
      return 'formula'
    case 'input':
      return 'input'
    case 'value':
      return 'value'
    default:
      return reason
  }
}

export function summarizePlainText(text: string | null, maxLength = 88): string | null {
  if (!text) {
    return null
  }
  const normalized = text.trim().replaceAll(/\s+/g, ' ')
  if (normalized.length === 0) {
    return null
  }
  return normalized.length <= maxLength ? normalized : `${normalized.slice(0, maxLength - 3)}...`
}

export function safeParseToolOutput(outputText: string | null): unknown {
  if (!outputText) {
    return null
  }
  try {
    return JSON.parse(outputText) as unknown
  } catch {
    return null
  }
}

export function renderToolDisplayName(toolName: string | null): string {
  const normalizedToolName = toolName ? normalizeWorkbookAgentToolName(toolName) : null
  if (!normalizedToolName) {
    return 'Tool call'
  }
  if (normalizedToolName === 'command_execution') {
    return 'Command'
  }
  return normalizedToolName
    .split('_')
    .map((segment) => (segment.length === 0 ? segment : `${segment[0]!.toUpperCase()}${segment.slice(1)}`))
    .join(' ')
}

export function summarizeToolEntry(entry: WorkbookAgentTimelineEntry): string | null {
  if (entry.toolName === 'command_execution') {
    const parsedArguments = safeParseToolOutput(entry.argumentsText)
    const command = isRecord(parsedArguments) ? readString(parsedArguments['command']) : ''
    const exitCode = isRecord(parsedArguments) ? parsedArguments['exitCode'] : null
    const durationMs = isRecord(parsedArguments) ? parsedArguments['durationMs'] : null
    const segments = [
      command ? `$ ${command}` : null,
      typeof exitCode === 'number' ? `exit ${String(exitCode)}` : null,
      typeof durationMs === 'number' ? `${String(Math.round(durationMs))} ms` : null,
    ].filter((segment) => segment !== null)
    if (segments.length > 0) {
      return summarizePlainText(segments.join(' · '), 96)
    }
  }
  const parsed = safeParseToolOutput(entry.outputText)
  if (isRecord(parsed)) {
    const mutationSummary = summarizeMutationReceipt(parsed)
    if (mutationSummary) {
      return summarizePlainText(mutationSummary, 96)
    }
    const verificationSummary = summarizeVerificationReport(parsed)
    if (verificationSummary) {
      return summarizePlainText(verificationSummary, 96)
    }
    if (typeof parsed['summary'] === 'string') {
      return summarizePlainText(parsed['summary'], 96)
    }
    const workflowRun = isRecord(parsed['workflowRun']) ? parsed['workflowRun'] : null
    if (typeof workflowRun?.['summary'] === 'string') {
      return summarizePlainText(workflowRun['summary'], 96)
    }
    const selection = isRecord(parsed['selection']) ? parsed['selection'] : null
    if (typeof selection?.['sheetName'] === 'string' && typeof selection['address'] === 'string') {
      const selectionRange = isRecord(selection['range']) ? selection['range'] : null
      const startAddress = typeof selectionRange?.['startAddress'] === 'string' ? selectionRange['startAddress'] : selection['address']
      const endAddress = typeof selectionRange?.['endAddress'] === 'string' ? selectionRange['endAddress'] : selection['address']
      return `${selection['sheetName']}!${startAddress}${startAddress === endAddress ? '' : `:${endAddress}`}`
    }
    const range = isRecord(parsed['range']) ? parsed['range'] : null
    if (typeof range?.['sheetName'] === 'string' && typeof range['startAddress'] === 'string' && typeof range['endAddress'] === 'string') {
      const startAddress = range['startAddress']
      const endAddress = range['endAddress']
      return `${range['sheetName']}!${startAddress}${startAddress === endAddress ? '' : `:${endAddress}`}`
    }
    if (typeof parsed['sheetCount'] === 'number') {
      return `${String(parsed['sheetCount'])} ${parsed['sheetCount'] === 1 ? 'sheet' : 'sheets'}`
    }
    if (typeof parsed['changeCount'] === 'number') {
      return `${String(parsed['changeCount'])} ${parsed['changeCount'] === 1 ? 'change' : 'changes'}`
    }
    if (typeof parsed['tableCount'] === 'number') {
      return formatCount(parsed['tableCount'], 'table')
    }
  }
  const outputText = entry.outputText?.trim() ?? ''
  if (outputText.length > 0 && !outputText.startsWith('{') && !outputText.startsWith('[')) {
    return summarizePlainText(outputText, 96)
  }
  const argumentsText = entry.argumentsText?.trim() ?? ''
  if (argumentsText.length > 0 && !argumentsText.startsWith('{') && !argumentsText.startsWith('[')) {
    return summarizePlainText(argumentsText, 96)
  }
  return null
}

export function proofMatchedLabel(proof: Record<string, unknown> | null): string {
  if (!proof) {
    return 'Missing'
  }
  if (proof['requested'] === false) {
    return 'Not requested'
  }
  if (proof['matched'] === true) {
    return 'Matched'
  }
  if (proof['matched'] === false) {
    return 'Mismatch'
  }
  return 'Incomplete'
}

export function undoProofLabel(undo: Record<string, unknown> | null): string {
  if (!undo) {
    return 'Missing'
  }
  if (undo['available'] === true) {
    return 'Available'
  }
  return undo['lookupFailed'] === true ? 'Lookup failed' : 'Unavailable'
}

export function renderedProofLabel(value: unknown): string {
  const proofs = Array.isArray(value) ? value.flatMap((proof) => (isRecord(proof) ? [proof] : [])) : []
  if (proofs.length === 0) {
    return 'Missing'
  }
  if (proofs.every((proof) => proof['matched'] === true)) {
    return 'Matched'
  }
  if (proofs.some((proof) => proof['matched'] === false)) {
    return 'Mismatch'
  }
  return 'Incomplete'
}

export function formulaAuditLabel(value: unknown): string {
  if (!isRecord(value)) {
    return 'Missing'
  }
  const summary = isRecord(value['summary']) ? value['summary'] : null
  const count = typeof summary?.['actionableIssueCount'] === 'number' ? summary['actionableIssueCount'] : null
  if (count === null) {
    return 'Incomplete'
  }
  return count === 0 ? 'Clean' : `${String(count)} actionable ${count === 1 ? 'issue' : 'issues'}`
}

export function invariantAuditLabel(value: unknown): string {
  if (!isRecord(value)) {
    return 'Missing'
  }
  const summary = isRecord(value['summary']) ? value['summary'] : null
  if (summary?.['ok'] === true) {
    return 'Clean'
  }
  if (summary?.['ok'] === false) {
    return 'Problems found'
  }
  return 'Incomplete'
}
