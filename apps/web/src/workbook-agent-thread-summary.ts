export function formatWorkbookAgentThreadEntryCount(entryCount: number): string {
  return `${entryCount} ${entryCount === 1 ? 'item' : 'items'}`
}

export function summarizeWorkbookAgentThreadActivity(text: string | null, limit = 72): string | null {
  if (!text) {
    return null
  }
  const normalized = text.trim().replaceAll(/\s+/g, ' ')
  if (normalized.length === 0) {
    return null
  }
  return normalized.length <= limit ? normalized : `${normalized.slice(0, Math.max(0, limit - 3))}...`
}
