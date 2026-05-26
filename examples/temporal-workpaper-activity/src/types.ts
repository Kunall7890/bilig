export interface TemporalWorkPaperQuoteInput {
  readonly previousQuantity: number
  readonly quantity: number
  readonly unitPrice: number
  readonly discountRate: number
  readonly taxRate: number
  readonly unitCost: number
  readonly output: string
}

export interface TemporalWorkPaperQuoteSummary {
  readonly subtotal: number
  readonly discountAmount: number
  readonly taxableAmount: number
  readonly taxAmount: number
  readonly total: number
  readonly marginAmount: number
}

export interface TemporalWorkPaperQuoteResult {
  readonly patch: {
    readonly subtotal: number
    readonly discount_amount: number
    readonly taxable_amount: number
    readonly tax_amount: number
    readonly total: number
    readonly margin_amount: number
  }
  readonly proof: {
    readonly editedCell: 'Inputs!B2'
    readonly before: TemporalWorkPaperQuoteSummary
    readonly after: TemporalWorkPaperQuoteSummary
    readonly afterRestore: TemporalWorkPaperQuoteSummary
    readonly persistedDocumentBytes: number
    readonly outputFile: string
    readonly verified: boolean
  }
  readonly temporalBoundary: {
    readonly workflowImportsWorkPaper: false
    readonly activityOwnsWorkPaper: true
    readonly payloadShape: 'serializable-patch-and-proof'
  }
  readonly limitations: readonly string[]
}

export interface TemporalWorkPaperActivities {
  readonly calculateWorkPaperQuoteActivity: (input: TemporalWorkPaperQuoteInput) => Promise<TemporalWorkPaperQuoteResult>
}
