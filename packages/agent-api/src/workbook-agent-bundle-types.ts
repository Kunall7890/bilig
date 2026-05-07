import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
  WorkbookChartSnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookImageSnapshot,
  WorkbookNoteSnapshot,
  WorkbookPivotSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookShapeSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookTableSnapshot,
} from '@bilig/protocol'

export interface WorkbookAgentUiSelectionRef {
  sheetName: string
  address: string
  range?: {
    startAddress: string
    endAddress: string
  }
}

export interface WorkbookAgentViewportRef {
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

export interface WorkbookAgentContextRef {
  selection: WorkbookAgentUiSelectionRef
  viewport: WorkbookAgentViewportRef
}

export type WorkbookAgentWriteCellInput =
  | LiteralInput
  | {
      value: LiteralInput
    }
  | {
      formula: string
    }

export type WorkbookAgentCommand =
  | {
      kind: 'writeRange'
      sheetName: string
      startAddress: string
      values: WorkbookAgentWriteCellInput[][]
    }
  | {
      kind: 'setRangeFormulas'
      range: CellRangeRef
      formulas: string[][]
    }
  | {
      kind: 'clearRange'
      range: CellRangeRef
    }
  | {
      kind: 'formatRange'
      range: CellRangeRef
      patch?: CellStylePatch
      numberFormat?: CellNumberFormatInput
    }
  | {
      kind: 'fillRange'
      source: CellRangeRef
      target: CellRangeRef
    }
  | {
      kind: 'copyRange'
      source: CellRangeRef
      target: CellRangeRef
    }
  | {
      kind: 'moveRange'
      source: CellRangeRef
      target: CellRangeRef
    }
  | {
      kind: 'upsertDefinedName'
      name: string
      value: WorkbookDefinedNameValueSnapshot
    }
  | {
      kind: 'deleteDefinedName'
      name: string
    }
  | {
      kind: 'upsertTable'
      table: WorkbookTableSnapshot
    }
  | {
      kind: 'deleteTable'
      name: string
    }
  | {
      kind: 'upsertPivotTable'
      pivot: WorkbookPivotSnapshot
    }
  | {
      kind: 'deletePivotTable'
      sheetName: string
      address: string
    }
  | {
      kind: 'upsertChart'
      chart: WorkbookChartSnapshot
    }
  | {
      kind: 'deleteChart'
      id: string
    }
  | {
      kind: 'upsertImage'
      image: WorkbookImageSnapshot
    }
  | {
      kind: 'deleteImage'
      id: string
    }
  | {
      kind: 'upsertShape'
      shape: WorkbookShapeSnapshot
    }
  | {
      kind: 'deleteShape'
      id: string
    }
  | {
      kind: 'createSheet'
      name: string
    }
  | {
      kind: 'renameSheet'
      currentName: string
      nextName: string
    }
  | {
      kind: 'deleteSheet'
      name: string
    }
  | {
      kind: 'insertRows'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'deleteRows'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'insertColumns'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'deleteColumns'
      sheetName: string
      start: number
      count: number
    }
  | {
      kind: 'setFreezePane'
      sheetName: string
      rows: number
      cols: number
    }
  | {
      kind: 'setFilter'
      range: CellRangeRef
    }
  | {
      kind: 'clearFilter'
      range: CellRangeRef
    }
  | {
      kind: 'setSort'
      range: CellRangeRef
      keys: {
        keyAddress: string
        direction: 'asc' | 'desc'
      }[]
    }
  | {
      kind: 'clearSort'
      range: CellRangeRef
    }
  | {
      kind: 'setDataValidation'
      validation: WorkbookDataValidationSnapshot
    }
  | {
      kind: 'clearDataValidation'
      range: CellRangeRef
    }
  | {
      kind: 'upsertConditionalFormat'
      format: WorkbookConditionalFormatSnapshot
    }
  | {
      kind: 'deleteConditionalFormat'
      id: string
      range: CellRangeRef
    }
  | {
      kind: 'setSheetProtection'
      protection: WorkbookSheetProtectionSnapshot
    }
  | {
      kind: 'clearSheetProtection'
      sheetName: string
    }
  | {
      kind: 'upsertRangeProtection'
      protection: WorkbookRangeProtectionSnapshot
    }
  | {
      kind: 'deleteRangeProtection'
      id: string
      range: CellRangeRef
    }
  | {
      kind: 'upsertCommentThread'
      thread: WorkbookCommentThreadSnapshot
    }
  | {
      kind: 'deleteCommentThread'
      sheetName: string
      address: string
    }
  | {
      kind: 'upsertNote'
      note: WorkbookNoteSnapshot
    }
  | {
      kind: 'deleteNote'
      sheetName: string
      address: string
    }
  | {
      kind: 'updateRowMetadata'
      sheetName: string
      startRow: number
      count: number
      height?: number | null
      hidden?: boolean | null
    }
  | {
      kind: 'updateColumnMetadata'
      sheetName: string
      startCol: number
      count: number
      width?: number | null
      hidden?: boolean | null
    }

export type WorkbookAgentRiskClass = 'low' | 'medium' | 'high'
export type WorkbookAgentBundleScope = 'selection' | 'sheet' | 'workbook'
export type WorkbookAgentAppliedBy = 'user' | 'auto'
export type WorkbookAgentAcceptedScope = 'full' | 'partial'
export type WorkbookAgentSharedReviewStatus = 'pending' | 'approved' | 'rejected'
export type WorkbookAgentPreviewRangeRole = 'target' | 'source'
export type WorkbookAgentPreviewChangeKind = 'input' | 'formula' | 'style' | 'numberFormat'

export interface WorkbookAgentPreviewRange {
  sheetName: string
  startAddress: string
  endAddress: string
  role: WorkbookAgentPreviewRangeRole
}

export interface WorkbookAgentPreviewCellDiff {
  sheetName: string
  address: string
  beforeInput: LiteralInput | null
  beforeFormula: string | null
  afterInput: LiteralInput | null
  afterFormula: string | null
  changeKinds: WorkbookAgentPreviewChangeKind[]
}

export interface WorkbookAgentPreviewEffectSummary {
  displayedCellDiffCount: number
  truncatedCellDiffs: boolean
  inputChangeCount: number
  formulaChangeCount: number
  styleChangeCount: number
  numberFormatChangeCount: number
  structuralChangeCount: number
}

export interface WorkbookAgentPreviewSummary {
  ranges: WorkbookAgentPreviewRange[]
  structuralChanges: string[]
  cellDiffs: WorkbookAgentPreviewCellDiff[]
  effectSummary: WorkbookAgentPreviewEffectSummary
}

export interface WorkbookAgentCommandBundle {
  id: string
  documentId: string
  threadId: string
  turnId: string
  goalText: string
  summary: string
  scope: WorkbookAgentBundleScope
  riskClass: WorkbookAgentRiskClass
  baseRevision: number
  createdAtUnixMs: number
  context: WorkbookAgentContextRef | null
  commands: WorkbookAgentCommand[]
  affectedRanges: WorkbookAgentPreviewRange[]
  estimatedAffectedCells: number | null
  sharedReview?: WorkbookAgentSharedReviewState | null
}

export interface WorkbookAgentSharedReviewState {
  ownerUserId: string
  status: WorkbookAgentSharedReviewStatus
  decidedByUserId: string | null
  decidedAtUnixMs: number | null
  recommendations: WorkbookAgentSharedReviewRecommendation[]
}

export interface WorkbookAgentSharedReviewRecommendation {
  userId: string
  decision: Extract<WorkbookAgentSharedReviewStatus, 'approved' | 'rejected'>
  decidedAtUnixMs: number
}

export interface WorkbookAgentExecutionRecord {
  id: string
  bundleId: string
  documentId: string
  threadId: string
  turnId: string
  actorUserId: string
  goalText: string
  planText: string | null
  summary: string
  scope: WorkbookAgentBundleScope
  riskClass: WorkbookAgentRiskClass
  acceptedScope: WorkbookAgentAcceptedScope
  appliedBy: WorkbookAgentAppliedBy
  baseRevision: number
  appliedRevision: number
  createdAtUnixMs: number
  appliedAtUnixMs: number
  context: WorkbookAgentContextRef | null
  commands: WorkbookAgentCommand[]
  preview: WorkbookAgentPreviewSummary | null
}
