import type { ImportedLegacyCommentVmlSheetSource } from './xlsx-comment-vml.js'
import type { ImportedWorkbookControlArtifactSheetSource } from './xlsx-control-artifacts.js'
import type { LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'
import { parseRelationships } from './xlsx-pivot-artifacts.js'
import type { ImportedWorkbookSlicerConnectionSheetSource } from './xlsx-slicer-connection-artifacts.js'
import { getZipText, normalizeZipPath, type XlsxZipEntries } from './xlsx-zip.js'

const queryTableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable'
const slicerRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicer'

interface LargeSimplePackageArtifactWorksheet {
  readonly name: string
  readonly order: number
  readonly metadataScan: LargeSimpleWorksheetScannedMetadata | undefined
}

interface LargeSimplePackageArtifactWorksheetEntry {
  readonly name?: string
  readonly path: string
}

function relationshipPartPath(partPath: string): string {
  const normalized = normalizeZipPath(partPath)
  const slashIndex = normalized.lastIndexOf('/')
  const directory = slashIndex >= 0 ? normalized.slice(0, slashIndex) : ''
  const fileName = slashIndex >= 0 ? normalized.slice(slashIndex + 1) : normalized
  return directory.length > 0 ? `${directory}/_rels/${fileName}.rels` : `_rels/${fileName}.rels`
}

export function largeSimpleSlicerConnectionRelationshipSheetNames(
  zip: XlsxZipEntries,
  worksheetEntries: readonly LargeSimplePackageArtifactWorksheetEntry[],
): ReadonlySet<string> {
  const sheetNames = new Set<string>()
  for (const worksheetEntry of worksheetEntries) {
    const sheetName = worksheetEntry.name
    if (!sheetName) {
      continue
    }
    const relationshipsXml = getZipText(zip, relationshipPartPath(worksheetEntry.path))
    if (
      parseRelationships(relationshipsXml).some(
        (relationship) => relationship.type === queryTableRelationshipType || relationship.type === slicerRelationshipType,
      )
    ) {
      sheetNames.add(sheetName)
    }
  }
  return sheetNames
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
            readSheetXmlForSlicerList: false,
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
