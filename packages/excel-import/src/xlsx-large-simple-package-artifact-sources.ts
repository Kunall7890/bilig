import type { ImportedLegacyCommentVmlSheetSource } from './xlsx-comment-vml.js'
import type { ImportedWorkbookControlArtifactSheetSource } from './xlsx-control-artifacts.js'
import type { LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'
import type { ImportedWorkbookSlicerConnectionSheetSource } from './xlsx-slicer-connection-artifacts.js'

interface LargeSimplePackageArtifactWorksheet {
  readonly name: string
  readonly order: number
  readonly metadataScan: LargeSimpleWorksheetScannedMetadata | undefined
}

interface LargeSimplePackageArtifactWorksheetEntry {
  readonly path: string
}

export function largeSimpleSlicerConnectionSheetSources(
  scannedWorksheets: readonly (LargeSimplePackageArtifactWorksheet | undefined)[],
  worksheetEntries: readonly LargeSimplePackageArtifactWorksheetEntry[],
): ImportedWorkbookSlicerConnectionSheetSource[] {
  return scannedWorksheets.flatMap((scanned) => {
    const sheetPath = scanned ? worksheetEntries[scanned.order]?.path : undefined
    return scanned && sheetPath
      ? [
          {
            sheetName: scanned.name,
            sheetPath,
            ...(scanned.metadataScan?.sheetSlicerListExtXml ? { sheetSlicerListExtXml: scanned.metadataScan.sheetSlicerListExtXml } : {}),
          },
        ]
      : []
  })
}

export function largeSimpleControlArtifactSheetSources(
  scannedWorksheets: readonly (LargeSimplePackageArtifactWorksheet | undefined)[],
  worksheetEntries: readonly LargeSimplePackageArtifactWorksheetEntry[],
): ImportedWorkbookControlArtifactSheetSource[] {
  return scannedWorksheets.flatMap((scanned) => {
    const controlArtifacts = scanned?.metadataScan?.controlArtifacts
    const legacyDrawingRelationshipId = scanned?.metadataScan?.legacyDrawingRelationshipId
    const sheetPath = scanned ? worksheetEntries[scanned.order]?.path : undefined
    return scanned && sheetPath && (controlArtifacts || legacyDrawingRelationshipId)
      ? [
          {
            sheetName: scanned.name,
            sheetPath,
            ...(controlArtifacts ? { ...controlArtifacts } : {}),
            ...(legacyDrawingRelationshipId ? { legacyDrawingRelationshipId } : {}),
          },
        ]
      : []
  })
}

export function largeSimpleLegacyCommentVmlSheetSources(
  scannedWorksheets: readonly (LargeSimplePackageArtifactWorksheet | undefined)[],
  worksheetEntries: readonly LargeSimplePackageArtifactWorksheetEntry[],
): ImportedLegacyCommentVmlSheetSource[] {
  return scannedWorksheets.flatMap((scanned) => {
    const legacyDrawingRelationshipId = scanned?.metadataScan?.legacyDrawingRelationshipId
    const sheetPath = scanned ? worksheetEntries[scanned.order]?.path : undefined
    return scanned && sheetPath && legacyDrawingRelationshipId ? [{ sheetName: scanned.name, sheetPath, legacyDrawingRelationshipId }] : []
  })
}
