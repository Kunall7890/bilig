import type { FormulaFamilyStore } from '../formula/formula-family-store.js'
import type { EngineCellStateService } from './services/cell-state-service.js'
import type { EngineEventService } from './services/event-service.js'
import type { EngineFormulaBindingService } from './services/formula-binding-service.js'
import type { EngineFormulaEvaluationService } from './services/formula-evaluation-service.js'
import type { EngineFormulaGraphService } from './services/formula-graph-service.js'
import type { EngineFormulaInitializationService } from './services/formula-initialization-service.js'
import type { EngineHistoryService } from './services/history-service.js'
import type { EngineMaintenanceService } from './services/maintenance-service.js'
import type { EngineMutationSupportService } from './services/mutation-support-service.js'
import type { EngineMutationService } from './services/mutation-service.js'
import type { EngineOperationService } from './services/operation-service.js'
import type { EnginePivotService } from './services/pivot-service.js'
import type { EngineReadService } from './services/read-service.js'
import type { EngineRecalcService } from './services/recalc-service.js'
import type { EngineReplicaSyncService } from './services/replica-sync-service.js'
import type { EngineSelectionService } from './services/selection-service.js'
import type { EngineSnapshotService } from './services/snapshot-service.js'
import type { EngineStructureService } from './services/structure-service.js'
import type { EngineTraversalService } from './services/traversal-service.js'

export interface EngineServiceRuntime {
  readonly cellState: EngineCellStateService
  readonly maintenance: EngineMaintenanceService
  readonly traversal: EngineTraversalService
  readonly deferKernelSync: (cellIndices: readonly number[] | Uint32Array) => void
  readonly hasVolatileFormulas: () => boolean
  readonly hasRegionFormulaSubscriptions: () => boolean
  readonly events: EngineEventService
  readonly evaluation: EngineFormulaEvaluationService
  readonly selection: EngineSelectionService
  readonly binding: EngineFormulaBindingService
  readonly formulaFamilies: FormulaFamilyStore
  readonly formulaInitialization: EngineFormulaInitializationService
  readonly graph: EngineFormulaGraphService
  readonly history: EngineHistoryService
  readonly mutation: EngineMutationService
  readonly support: EngineMutationSupportService
  readonly operations: EngineOperationService
  readonly pivot: EnginePivotService
  readonly read: EngineReadService
  readonly recalc: EngineRecalcService
  readonly structure: EngineStructureService
  readonly snapshot: EngineSnapshotService
  readonly sync: EngineReplicaSyncService
}
