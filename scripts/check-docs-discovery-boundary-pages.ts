import { requireGoogleSheetsQuerySortnDiscovery } from './check-docs-discovery-google-sheets.ts'
import { requireMicrosoftGraphExcelBoundaryDiscovery } from './check-docs-discovery-microsoft-graph.ts'

export function requireBoundaryPageDiscovery(input: {
  readonly googleSheetsQuerySortnNodeWorkpaperDoc: string
  readonly headlessReadme: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly microsoftGraphExcelRecalculationNode: string
  readonly readme: string
}): void {
  requireMicrosoftGraphExcelBoundaryDiscovery({
    doc: input.microsoftGraphExcelRecalculationNode,
    headlessReadme: input.headlessReadme,
    index: input.index,
    llms: input.llms,
    readme: input.readme,
  })

  requireGoogleSheetsQuerySortnDiscovery({
    doc: input.googleSheetsQuerySortnNodeWorkpaperDoc,
    headlessReadme: input.headlessReadme,
    index: input.index,
    llms: input.llms,
    llmsFull: input.llmsFull,
    readme: input.readme,
  })
}
