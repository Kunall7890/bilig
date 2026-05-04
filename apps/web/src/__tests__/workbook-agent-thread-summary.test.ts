import { describe, expect, it } from 'vitest'
import { formatWorkbookAgentThreadEntryCount, summarizeWorkbookAgentThreadActivity } from '../workbook-agent-thread-summary.js'

describe('workbook agent thread summary helpers', () => {
  it('formats thread entry counts with singular and plural labels', () => {
    expect(formatWorkbookAgentThreadEntryCount(1)).toBe('1 item')
    expect(formatWorkbookAgentThreadEntryCount(3)).toBe('3 items')
  })

  it('normalizes and truncates latest activity text', () => {
    expect(summarizeWorkbookAgentThreadActivity(null)).toBeNull()
    expect(summarizeWorkbookAgentThreadActivity('   ')).toBeNull()
    expect(summarizeWorkbookAgentThreadActivity('  Follow   up   on totals  ')).toBe('Follow up on totals')
    expect(summarizeWorkbookAgentThreadActivity('1234567890ABCDE', 10)).toBe('1234567...')
  })
})
