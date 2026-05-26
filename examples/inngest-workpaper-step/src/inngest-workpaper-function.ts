import { Inngest } from 'inngest'

import { calculateWorkPaperQuote, type InngestWorkPaperQuoteInput } from './workpaper-quote.js'

export const inngest = new Inngest({ id: 'bilig-workpaper-example' })

export const calculateWorkPaperQuoteFunction = inngest.createFunction(
  {
    id: 'bilig-workpaper-quote',
    retries: 3,
    triggers: [{ event: 'bilig/quote.requested' }],
  },
  async ({ event, step }) => {
    const result = await step.run('calculate-workpaper-quote', async () => calculateWorkPaperQuote(readInngestQuoteInput(event.data)))

    if (!result.proof.verified) {
      throw new Error(`WorkPaper proof failed: ${JSON.stringify(result.proof)}`)
    }

    return result
  },
)

function readInngestQuoteInput(value: unknown): InngestWorkPaperQuoteInput {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error('bilig/quote.requested data must be an object')
  }

  return {
    previousQuantity: readOptionalNumberProperty(value, 'previousQuantity'),
    quantity: readNumberProperty(value, 'quantity'),
    unitPrice: readNumberProperty(value, 'unitPrice'),
    discountRate: readOptionalNumberProperty(value, 'discountRate'),
    taxRate: readOptionalNumberProperty(value, 'taxRate'),
    unitCost: readOptionalNumberProperty(value, 'unitCost'),
  }
}

function readNumberProperty(value: object, property: keyof InngestWorkPaperQuoteInput): number {
  const propertyValue = Reflect.get(value, property)

  if (typeof propertyValue !== 'number' || !Number.isFinite(propertyValue)) {
    throw new Error(`bilig/quote.requested data.${property} must be a finite number`)
  }

  return propertyValue
}

function readOptionalNumberProperty(value: object, property: keyof InngestWorkPaperQuoteInput): number | undefined {
  const propertyValue = Reflect.get(value, property)

  if (propertyValue === undefined || propertyValue === null) {
    return undefined
  }

  if (typeof propertyValue !== 'number' || !Number.isFinite(propertyValue)) {
    throw new Error(`bilig/quote.requested data.${property} must be a finite number when present`)
  }

  return propertyValue
}
