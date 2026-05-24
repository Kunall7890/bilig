import type { Unzipped } from 'fflate'
import { isDataModelPackagePartPath } from './xlsx-data-model-artifacts.js'

export function shouldBypassLargeSimpleByteThresholdForPackageArtifacts(workbookZip: Unzipped): boolean {
  return Object.keys(workbookZip).some(isDataModelPackagePartPath)
}

export function hasFullImporterOnlyPackageMetadata(workbookZip: Unzipped): boolean {
  return Object.keys(workbookZip).some((path) => path.startsWith('xl/comments') || path.startsWith('xl/threadedComments/'))
}
