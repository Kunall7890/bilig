import type { WorkbookAgentUiContext } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookVerificationMismatch } from './workbook-agent-rendered-readback.js'

export interface WorkbookAgentMutationProofContext {
  readonly documentId: string
  readonly session?: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
  readonly stageCommand?: unknown
}

export interface WorkbookAuthoritativeReadbackProof {
  readonly requested: boolean
  readonly matched: boolean | null
  readonly ranges: readonly unknown[]
  readonly mismatches: readonly WorkbookVerificationMismatch[]
  readonly incompleteReason: string | null
}

export interface WorkbookSemanticReadbackProof {
  readonly requested: boolean
  readonly matched: boolean | null
  readonly incompleteReason: string | null
}

export interface WorkbookMutationUndoProof {
  readonly available: boolean
  readonly token: string | null
  readonly reasonUnavailable: string | null
  readonly lookupFailed: boolean
}

export interface WorkbookMutationRecalculationProof {
  readonly requested: boolean
  readonly upToDate: boolean | null
  readonly appliedRevision: number | null
  readonly headRevision: number | null
  readonly calculatedRevision: number | null
  readonly lastMetrics: unknown
  readonly incompleteReason: string | null
}
