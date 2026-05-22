export const workbookPath = 'xl/workbook.xml'
export const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
export const sharedStringsPath = 'xl/sharedStrings.xml'
export const stylesPath = 'xl/styles.xml'

const unsupportedPackagePathPattern = /^xl\/(?:ctrlProps|threadedComments|vbaProject\.bin)/u

export interface LargeSimplePackageFlags {
  readonly hasSharedStrings: boolean
  readonly hasStyles: boolean
  readonly hasCalcChain: boolean
  readonly hasDrawingParts: boolean
  readonly hasChartParts: boolean
  readonly hasPivotParts: boolean
  readonly hasExternalLinkParts: boolean
  readonly hasLegacyCommentParts: boolean
  readonly hasDataModelParts: boolean
  readonly hasSlicerConnectionParts: boolean
}

export function hasUnsupportedLargeSimplePackagePath(packagePaths: readonly string[]): boolean {
  return packagePaths.some((path) => unsupportedPackagePathPattern.test(path))
}

export function readLargeSimplePackageFlags(packagePaths: readonly string[]): LargeSimplePackageFlags {
  return {
    hasSharedStrings: packagePaths.includes(sharedStringsPath),
    hasStyles: packagePaths.includes(stylesPath),
    hasCalcChain: packagePaths.includes('xl/calcChain.xml'),
    hasDrawingParts: packagePaths.some((path) => path.startsWith('xl/drawings/') || path.startsWith('xl/media/')),
    hasChartParts: packagePaths.some((path) => path.startsWith('xl/charts/') || path.startsWith('xl/chartSheets/')),
    hasPivotParts: packagePaths.some((path) => path.startsWith('xl/pivotTables/') || path.startsWith('xl/pivotCache/')),
    hasExternalLinkParts: packagePaths.some((path) => path.startsWith('xl/externalLinks/')),
    hasLegacyCommentParts: packagePaths.some((path) => path.startsWith('xl/comments') || path.endsWith('.vml')),
    hasDataModelParts: packagePaths.some(
      (path) => path.startsWith('xl/model/') || path.startsWith('xl/customData/') || path.startsWith('customXml/'),
    ),
    hasSlicerConnectionParts: packagePaths.some(
      (path) => path === 'xl/connections.xml' || path.startsWith('xl/slicerCaches/') || path.startsWith('xl/slicers/'),
    ),
  }
}
