import { proxyActivities } from '@temporalio/workflow'
import type { TemporalWorkPaperActivities, TemporalWorkPaperQuoteInput, TemporalWorkPaperQuoteResult } from './types'

const { calculateWorkPaperQuoteActivity } = proxyActivities<TemporalWorkPaperActivities>({
  startToCloseTimeout: '30 seconds',
  retry: {
    maximumAttempts: 3,
  },
})

export async function quoteApprovalWorkflow(input: TemporalWorkPaperQuoteInput): Promise<TemporalWorkPaperQuoteResult> {
  return await calculateWorkPaperQuoteActivity(input)
}
