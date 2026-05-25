import { task } from '@trigger.dev/sdk'

import { calculateWorkPaperQuote, type TriggerDevWorkPaperQuoteInput } from './workpaper-quote.js'

export const calculateWorkPaperQuoteTask = task({
  id: 'bilig-workpaper-quote',
  retry: {
    maxAttempts: 3,
    minTimeoutInMs: 500,
    maxTimeoutInMs: 30_000,
    factor: 1.8,
    randomize: true,
  },
  run: async (payload: TriggerDevWorkPaperQuoteInput) => calculateWorkPaperQuote(payload),
})
