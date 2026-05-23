import type {
  WorkbookCommentThreadSnapshot,
  WorkbookLegacyCommentVmlSnapshot,
  WorkbookSheetThreadedCommentArtifactsSnapshot,
  WorkbookThreadedCommentArtifactsSnapshot,
} from '@bilig/protocol'

import type {
  WorkbookCommentThreadRecord,
  WorkbookSheetLegacyCommentVmlRecord,
  WorkbookSheetThreadedCommentArtifactsRecord,
  WorkbookThreadedCommentArtifactsRecord,
} from './workbook-metadata-types.js'
import { runWorkbookMetadataEffect } from './workbook-metadata-service.js'
import type { WorkbookMetadataService } from './workbook-metadata-service-contract.js'

export abstract class WorkbookStoreCommentAccessors {
  protected abstract readonly metadataService: WorkbookMetadataService

  setThreadedCommentArtifacts(artifacts: WorkbookThreadedCommentArtifactsSnapshot): WorkbookThreadedCommentArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setThreadedCommentArtifacts(artifacts))
  }

  getThreadedCommentArtifacts(): WorkbookThreadedCommentArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getThreadedCommentArtifacts())
  }

  clearThreadedCommentArtifacts(): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearThreadedCommentArtifacts())
  }

  setSheetThreadedCommentArtifacts(
    sheetName: string,
    artifacts: WorkbookSheetThreadedCommentArtifactsSnapshot,
  ): WorkbookSheetThreadedCommentArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSheetThreadedCommentArtifacts(sheetName, artifacts))
  }

  getSheetThreadedCommentArtifacts(sheetName: string): WorkbookSheetThreadedCommentArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSheetThreadedCommentArtifacts(sheetName))
  }

  deleteSheetThreadedCommentArtifacts(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSheetThreadedCommentArtifacts(sheetName))
  }

  setSheetLegacyCommentVml(sheetName: string, legacyCommentVml: WorkbookLegacyCommentVmlSnapshot): WorkbookSheetLegacyCommentVmlRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSheetLegacyCommentVml(sheetName, legacyCommentVml))
  }

  getSheetLegacyCommentVml(sheetName: string): WorkbookSheetLegacyCommentVmlRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSheetLegacyCommentVml(sheetName))
  }

  deleteSheetLegacyCommentVml(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSheetLegacyCommentVml(sheetName))
  }

  setCommentThread(record: WorkbookCommentThreadSnapshot): WorkbookCommentThreadRecord {
    return runWorkbookMetadataEffect(this.metadataService.setCommentThread(record))
  }

  getCommentThread(sheetName: string, address: string): WorkbookCommentThreadRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getCommentThread(sheetName, address))
  }

  deleteCommentThread(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteCommentThread(sheetName, address))
  }

  listCommentThreads(sheetName: string): WorkbookCommentThreadRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listCommentThreads(sheetName))
  }
}
