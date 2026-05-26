import type {
  SameCorpusCaptureCorpusVerification,
  UiResponsivenessSameCorpusProduct,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'

export interface PreflightProductResult {
  readonly product: Exclude<UiResponsivenessSameCorpusProduct, 'bilig'>
  readonly source: string
  readonly finalUrl: string
  readonly title: string
  readonly status: 'ready' | 'blocked'
  readonly blocker: string | null
  readonly corpusVerification: SameCorpusCaptureCorpusVerification | null
  readonly limitations: string[]
}

export interface SameCorpusPreflight {
  readonly mode: 'preflight'
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly requiredProductCount: 2
  readonly checkedProductCount: number
  readonly readyProductCount: number
  readonly blockedProductCount: number
  readonly allCheckedProductsReady: boolean
  readonly products: readonly PreflightProductResult[]
}
