import { readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

const testDir = dirname(fileURLToPath(import.meta.url))
const repoRoot = join(testDir, '..', '..', '..', '..')

interface PackageManifest {
  readonly scripts?: Record<string, string>
}

const existingDesktopExcelOracleFiles = [
  'macos-desktop-excel-array-formula-structural-oracle.test.ts',
  'macos-desktop-excel-autofilter-oracle.test.ts',
  'macos-desktop-excel-calc-chain-reorder-oracle.test.ts',
  'macos-desktop-excel-chart-deleted-sheet-oracle.test.ts',
  'macos-desktop-excel-chart-drawing-anchor-oracle.test.ts',
  'macos-desktop-excel-cell-metadata-oracle.test.ts',
  'macos-desktop-excel-conditional-format-artifacts-oracle.test.ts',
  'macos-desktop-excel-control-artifacts-oracle.test.ts',
  'macos-desktop-excel-data-table-structural-oracle.test.ts',
  'macos-desktop-excel-defined-name-structural-oracle.test.ts',
  'macos-desktop-excel-drawing-artifacts-oracle.test.ts',
  'macos-desktop-excel-external-link-cache.test.ts',
  'macos-desktop-excel-hyperlink-structural-oracle.test.ts',
  'macos-desktop-excel-ignored-errors-oracle.test.ts',
  'macos-desktop-excel-preserved-package-metadata-oracle.test.ts',
  'macos-desktop-excel-precision-as-displayed-oracle.test.ts',
  'macos-desktop-excel-pivot-oracle.test.ts',
  'macos-desktop-excel-print-page-setup-oracle.test.ts',
  'macos-desktop-excel-rich-text-oracle.test.ts',
  'macos-desktop-excel-sheet-move-metadata-topology-oracle.test.ts',
  'macos-desktop-excel-sheet-properties-oracle.test.ts',
  'macos-desktop-excel-slicer-connection-delete-oracle.test.ts',
  'macos-desktop-excel-sort-oracle.test.ts',
  'macos-desktop-excel-sparklines-oracle.test.ts',
  'macos-desktop-excel-structured-reference-syntax.test.ts',
  'macos-desktop-excel-table-header-canonicalization.test.ts',
  'macos-desktop-excel-threaded-comment-structural-oracle.test.ts',
  'macos-desktop-excel-xlsx-oracle.test.ts',
] as const

const corpusDesktopExcelOracleFiles = [
  'macos-desktop-excel-array-formula-structural-oracle.test.ts',
  'macos-desktop-excel-chart-drawing-anchor-oracle.test.ts',
  'macos-desktop-excel-data-table-structural-oracle.test.ts',
  'macos-desktop-excel-defined-name-structural-oracle.test.ts',
  'macos-desktop-excel-external-link-cache.test.ts',
  'macos-desktop-excel-hyperlink-structural-oracle.test.ts',
  'macos-desktop-excel-pivot-oracle.test.ts',
  'macos-desktop-excel-sort-oracle.test.ts',
  'macos-desktop-excel-structured-reference-syntax.test.ts',
  'macos-desktop-excel-table-header-canonicalization.test.ts',
  'macos-desktop-excel-threaded-comment-structural-oracle.test.ts',
  'macos-desktop-excel-xlsx-oracle.test.ts',
] as const

describe('macOS Desktop Excel oracle inventory', () => {
  it('keeps high-value Desktop Excel oracle files under package ownership', () => {
    const present = new Set(readdirSync(testDir).filter((fileName) => /^macos-desktop-excel-.*\.test\.ts$/u.test(fileName)))

    for (const fileName of existingDesktopExcelOracleFiles) {
      expect(present.has(fileName), `${fileName} must stay package-owned under packages/headless/src/__tests__`).toBe(true)
    }
  })

  it('keeps the sort oracle anchored to live Desktop Excel table sort semantics', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-sort-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain("kind: 'applyTableSort'")
    expect(source).toContain('headless.sortTable')
    expect(source).toContain('matches Desktop Excel table-body sort row-bundle semantics')
  })

  it('keeps the conditional-format artifact oracle anchored to Desktop Excel advanced visual rules', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-conditional-format-artifacts-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('type="dataBar"')
    expect(source).toContain('type="colorScale"')
    expect(source).toContain('type="iconSet"')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('preserves Desktop Excel advanced visual conditional-format rules after a headless edit')
    expect(source).toContain('matches Desktop Excel cross-sheet conditional-format artifact formulas after target sheet row inserts')
    expect(source).toContain('matches Desktop Excel x14 conditional-format artifact ranges after owner sheet row inserts')
  })

  it('keeps the chart drawing anchor oracle anchored to Desktop Excel structural geometry', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-chart-drawing-anchor-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('editAs="oneCell"')
    expect(source).toContain('matches Desktop Excel chart drawing anchors after structural row inserts')
  })

  it('keeps the chart deleted sheet oracle anchored to Desktop Excel raw chart formulas', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-chart-deleted-sheet-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain("kind: 'deleteSheet'")
    expect(source).toContain('drawingArtifacts')
    expect(source).toContain('matches Desktop Excel raw chart formula invalidation after deleting a referenced sheet')
  })

  it('keeps the control artifact oracle anchored to Desktop Excel form-control geometry', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-control-artifacts-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('controlArtifacts')
    expect(source).toContain('macro="[0]!WriteHelloWorld"')
    expect(source).toContain('matches Desktop Excel worksheet form control anchors after structural row inserts')
  })

  it('keeps the threaded comment structural oracle anchored to Desktop Excel package refs', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-threaded-comment-structural-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('threadedCommentArtifacts')
    expect(source).toContain('legacyCommentVml')
    expect(source).toContain('matches Desktop Excel threaded comment refs after structural row inserts')
    expect(source).toContain('matches Desktop Excel threaded comment package cleanup after deleting a commented sheet')
  })

  it('keeps the hyperlink structural oracle anchored to Desktop Excel metadata movement', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-hyperlink-structural-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('metadata?.hyperlinks')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('matches Desktop Excel hyperlink metadata after structural row inserts')
  })

  it('keeps the external-link oracle anchored to Desktop Excel package artifacts', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-external-link-cache.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain("updateLinks: 'external'")
    expect(source).toContain('externalLinkArtifacts')
    expect(source).toContain('externalLinkPackageSummary')
    expect(source).toContain('round-trips cached external ranges through Desktop Excel and Bilig recalc')
  })

  it('keeps the sparkline oracle anchored to Desktop Excel extension XML refs', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-sparklines-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('sparklineExtensionUri')
    expect(source).toContain('metadata?.sparklines')
    expect(source).toContain('matches Desktop Excel sparkline source and output refs after structural row inserts')
    expect(source).toContain('matches Desktop Excel cross-sheet sparkline source refs after source sheet row inserts')
  })

  it('keeps the ignored-errors oracle anchored to Desktop Excel warning-suppression refs', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-ignored-errors-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('ignoredErrors')
    expect(source).toContain('numberStoredAsText="1"')
    expect(source).toContain('matches Desktop Excel ignoredErrors sqref movement after structural row inserts')
  })

  it('keeps the print page setup oracle anchored to Desktop Excel manual page breaks', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-print-page-setup-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('rowBreaksXml')
    expect(source).toContain('colBreaksXml')
    expect(source).toContain('matches Desktop Excel manual page-break ids after structural inserts')
  })

  it('keeps the preserved package metadata oracle anchored to Desktop Excel workbook view state', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-preserved-package-metadata-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('viewState')
    expect(source).toContain('styleArtifacts')
    expect(source).toContain('chartArtifacts')
    expect(source).toContain('preserves Desktop Excel workbook and sheet view state after a headless edit')
    expect(source).toContain('matches Desktop Excel workbook view tab indexes after deleting a prior sheet')
    expect(source).toContain('matches Desktop Excel workbook view tab indexes after moving a sheet tab')
    expect(source).toContain('matches Desktop Excel calc-chain sheet ids after deleting a prior sheet')
    expect(source).toContain('matches Desktop Excel preserved style artifacts and view refs after structural row inserts')
    expect(source).toContain('matches Desktop Excel preserved pivot cache source sheet after source sheet rename')
    expect(source).toContain('matches Desktop Excel raw worksheet chart package formulas after structural source row inserts')
    expect(source).toContain('matches Desktop Excel raw worksheet chart package formulas after source sheet rename')
  })

  it('keeps the calc-chain reorder oracle anchored to Desktop Excel sheet moves', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-calc-chain-reorder-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('moveSheet')
    expect(source).toContain('xl/calcChain.xml')
    expect(source).toContain('matches Desktop Excel calc-chain sheet ids after moving a sheet tab')
  })

  it('keeps the sheet-move metadata topology oracle anchored to Desktop Excel worksheet ownership', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-sheet-move-metadata-topology-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('moveSheet')
    expect(source).toContain('metadataCodeNames')
    expect(source).toContain('matches Desktop Excel worksheet metadata ownership after moving a sheet tab')
  })

  it('keeps the slicer connection delete oracle anchored to Desktop Excel package cleanup', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-slicer-connection-delete-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain("kind: 'deleteSheet'")
    expect(source).toContain('slicerConnectionTopology')
    expect(source).toContain('matches Desktop Excel slicer package cleanup after deleting the slicer sheet')
  })

  it('keeps the rich text oracle anchored to Desktop Excel cell run preservation', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-rich-text-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('richTextArtifacts')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('matches Desktop Excel rich text movement after structural row inserts')
  })

  it('keeps the cell metadata oracle anchored to Desktop Excel dynamic-array refs', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-cell-metadata-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('cellMetadataRefs')
    expect(source).toContain('XLDAPR')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('matches Desktop Excel cell metadata refs after structural row inserts')
  })

  it('keeps the worksheet properties oracle anchored to Desktop Excel sheetPr metadata', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-sheet-properties-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('sheetPr')
    expect(source).toContain('codeName="Sheet8"')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('matches Desktop Excel worksheet sheetPr properties after structural row inserts')
  })

  it('keeps the pivot oracle anchored to Desktop Excel source-backed pivot refresh semantics', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-pivot-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('refreshWorkbook: true')
    expect(source).toContain('invalid="1"')
    expect(source).toContain('GETPIVOTDATA("Sum of Sales",$B$2,"Region","East")')
    expect(source).toContain('cachedRecords')
    expect(source).toContain('matches Desktop Excel GETPIVOTDATA after headless source edits and XLSX export')
    expect(source).toContain('matches Desktop Excel pivot package cleanup after deleting a pivot sheet')
  })

  it('keeps the drawing artifact oracle anchored to Desktop Excel package preservation', () => {
    const source = readFileSync(join(testDir, 'macos-desktop-excel-drawing-artifacts-oracle.test.ts'), 'utf8')

    expect(source).toContain("BILIG_EXCEL_ORACLE_RUN === '1'")
    expect(source).toContain('runMacosExcelInspectionOracle')
    expect(source).toContain('runMacosExcelStructuralOperationOracle')
    expect(source).toContain('drawingArtifacts')
    expect(source).toContain('WorkPaper.buildFromSnapshot')
    expect(source).toContain('preserves Desktop Excel drawing package parts after a headless edit')
    expect(source).toContain('matches Desktop Excel raw DrawingML anchors after structural row inserts')
  })

  it('keeps the sort oracle in both package and corpus gates', () => {
    const headlessPackageJson = readPackageManifest(join(repoRoot, 'packages/headless/package.json'))
    const rootPackageJson = readPackageManifest(join(repoRoot, 'package.json'))

    expect(headlessPackageJson.scripts?.['test:excel-oracle']).toContain('src/__tests__/desktop-excel-oracle-inventory.test.ts')
    expect(headlessPackageJson.scripts?.['test:excel-oracle']).toContain('src/__tests__/macos-desktop-excel-*.test.ts')
    expect(headlessPackageJson.scripts?.['test:excel-oracle:live']).toContain('src/__tests__/desktop-excel-oracle-inventory.test.ts')
    expect(headlessPackageJson.scripts?.['test:excel-oracle:live']).toContain('src/__tests__/macos-desktop-excel-*.test.ts')
    expect(rootPackageJson.scripts?.['test:correctness:corpus']).toContain(
      'packages/headless/src/__tests__/desktop-excel-oracle-inventory.test.ts',
    )
    for (const fileName of corpusDesktopExcelOracleFiles) {
      expect(rootPackageJson.scripts?.['test:correctness:corpus']).toContain(`packages/headless/src/__tests__/${fileName}`)
    }
  })
})

function readPackageManifest(path: string): PackageManifest {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isRecord(parsed)) {
    return {}
  }
  const scripts = parsed['scripts']
  return isRecord(scripts) ? { scripts: stringRecord(scripts) } : {}
}

function stringRecord(record: Record<string, unknown>): Record<string, string> {
  return Object.fromEntries(Object.entries(record).filter((entry): entry is [string, string] => typeof entry[1] === 'string'))
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
