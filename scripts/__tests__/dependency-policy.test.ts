import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join, resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)
const packageManifestDirs = ['.', 'packages', 'apps', 'examples'] as const
const forbiddenDependencySources = Object.freeze(['cdn.sheetjs.com'])
const excelImportRuntimeXlsxImportAllowlist = new Set([
  'packages/excel-import/src/xlsx-export.ts',
  'packages/excel-import/src/xlsx-sheetjs-import.ts',
])
const liveScorecardFixtureScripts = [
  'scripts/gen-google-sheets-live-calculation-scorecard.ts',
  'scripts/gen-google-sheets-live-recalculation-scorecard.ts',
  'scripts/gen-google-sheets-live-structural-scorecard.ts',
  'scripts/gen-microsoft-excel-live-calculation-scorecard.ts',
  'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts',
  'scripts/gen-microsoft-excel-live-structural-scorecard.ts',
] as const
const nativeXlsxFixtureScripts = [
  ...liveScorecardFixtureScripts,
  'e2e/tests/web-shell-import.pw.ts',
  'scripts/gen-import-export-fidelity-scorecard.ts',
  'scripts/gen-security-posture-scorecard.ts',
  'scripts/gen-workpaper-xlsx-corpus-fixtures.ts',
] as const
const nativeXlsxCorpusProofScripts = ['scripts/check-workpaper-xlsx-corpus.ts', 'scripts/workpaper-xlsx-volatile-dependencies.ts'] as const
const nativeXlsxCorpusProofTests = [
  'scripts/__tests__/workpaper-xlsx-corpus-no-formula.test.ts',
  'scripts/__tests__/workpaper-xlsx-corpus.test.ts',
] as const
const nativeXlsxPublicWorkbookCorpusTests = [
  'scripts/__tests__/public-workbook-corpus-links.test.ts',
  'scripts/__tests__/public-workbook-corpus.test.ts',
  'scripts/__tests__/public-workbook-corpus-workbook.test.ts',
] as const
const nativeXlsxExampleScripts = [
  'examples/recalc-bridge-workflows/smoke.mjs',
  'examples/recalc-bridge-workflows/stackoverflow-sheetjs-63085785.mjs',
] as const
const nativeXlsxAppFixtureTests = ['apps/bilig/src/workbook-runtime/workbook-session-shared.test.ts'] as const
const nativeXlsxExcelFixtureTests = ['packages/excel-fixtures/src/__tests__/macos-excel-oracle.test.ts'] as const
const nativeXlsxHeadlessFixtureTests = [
  'packages/headless/src/__tests__/macos-desktop-excel-chart-deleted-sheet-oracle.test.ts',
  'packages/headless/src/__tests__/macos-desktop-excel-conditional-format-artifacts-oracle.test.ts',
  'packages/headless/src/__tests__/macos-desktop-excel-external-link-cache.test.ts',
  'packages/headless/src/__tests__/macos-desktop-excel-hyperlink-structural-oracle.test.ts',
  'packages/headless/src/__tests__/macos-desktop-excel-preserved-package-metadata-oracle.test.ts',
  'packages/headless/src/__tests__/macos-desktop-excel-protected-ranges-oracle.test.ts',
  'packages/headless/src/__tests__/macos-desktop-excel-threaded-comment-structural-oracle.test.ts',
  'packages/headless/src/__tests__/macos-desktop-excel-workbook-protection-oracle.test.ts',
  'packages/headless/src/__tests__/work-paper-source-preserving-xlsx-export.test.ts',
] as const
const nativeXlsxExcelImportFixtureTests = [
  'packages/excel-import/src/__tests__/excel-import.test.ts',
  'packages/excel-import/src/__tests__/excel-import-array-formulas.test.ts',
  'packages/excel-import/src/__tests__/excel-import.fuzz.test.ts',
  'packages/excel-import/src/__tests__/xlsx-alignment-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-byte-source-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-calculation-properties-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-cell-metadata-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-cell-rich-text-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-chart-artifacts-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-conditional-format-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-data-validations-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-date1904-system.test.ts',
  'packages/excel-import/src/__tests__/xlsx-defined-names-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-drawing-artifacts-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-error-cell-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-external-defined-names.test.ts',
  'packages/excel-import/src/__tests__/xlsx-export-large-simple.test.ts',
  'packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-formula-cache-text-normalization.test.ts',
  'packages/excel-import/src/__tests__/xlsx-hyperlink-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-large-simple-lazy-package-artifacts.test.ts',
  'packages/excel-import/src/__tests__/xlsx-large-simple-zip-release.test.ts',
  'packages/excel-import/src/__tests__/xlsx-legacy-comment-rich-text-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-legacy-comment-vml-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-ms-oi29500-calc-chain.test.ts',
  'packages/excel-import/src/__tests__/xlsx-ms-oi29500-external-data.test.ts',
  'packages/excel-import/src/__tests__/xlsx-ms-oi29500-formula-context.test.ts',
  'packages/excel-import/src/__tests__/xlsx-ms-oi29500-pivot-semantics.test.ts',
  'packages/excel-import/src/__tests__/xlsx-print-page-setup-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-printer-settings-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-protected-range-attributes-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/excel-import-sparse-range.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-protection-attributes-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-axis-visibility-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-workbook-protection-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-properties-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-name-whitespace.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-visibility-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-threaded-comments-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-view-state-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-workbook-sheet-paths.test.ts',
  'packages/excel-import/src/__tests__/xlsx-fallback-lazy-artifacts.test.ts',
  'packages/excel-import/src/__tests__/xlsx-worksheet-relationship-path-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-worksheet-dimensions-roundtrip.test.ts',
] as const
const nativeXlsxFormulaRecalcPackages = ['packages/xlsx-formula-recalc', 'packages/bilig-xlsx-formula-recalc'] as const
const nativeXlsxFormulaRecalcReadmes = ['packages/xlsx-formula-recalc/README.md', 'packages/bilig-xlsx-formula-recalc/README.md'] as const
const publishedNativeXlsxRuntimePackages = [
  'packages/xlsx',
  'packages/excel-import',
  'packages/headless',
  'packages/bilig',
  'packages/workpaper',
  'packages/xlsx-formula-recalc',
  'packages/bilig-xlsx-formula-recalc',
  'packages/exceljs-formula-recalc',
  'packages/bilig-exceljs-formula-recalc',
  'packages/sheetjs-formula-recalc',
  'packages/bilig-sheetjs-formula-recalc',
] as const
const fileBackedXlsxFormulaRecalcCliEntrypoints = [
  'packages/xlsx-formula-recalc/src/cli.ts',
  'packages/xlsx-formula-recalc/src/cache-doctor-cli.ts',
  'packages/xlsx-formula-recalc/src/sheetjs-cli.ts',
  'packages/bilig-xlsx-formula-recalc/src/cli.ts',
  'packages/bilig-xlsx-formula-recalc/src/cache-doctor-cli.ts',
  'packages/bilig-xlsx-formula-recalc/src/sheetjs-cli.ts',
  'packages/sheetjs-formula-recalc/src/cli.ts',
  'packages/exceljs-formula-recalc/src/cli.ts',
  'packages/bilig-sheetjs-formula-recalc/src/cli.ts',
  'packages/bilig-exceljs-formula-recalc/src/cli.ts',
] as const
const xlsxOwnedStreamingNativeSources = [
  'packages/xlsx/src/streaming-native-cell-arena.ts',
  'packages/xlsx/src/streaming-native-external-cache.ts',
  'packages/xlsx/src/streaming-native-inspect.ts',
  'packages/xlsx/src/streaming-native-recalc.ts',
  'packages/xlsx/src/streaming-native-row-chain-wasm.ts',
  'packages/xlsx/src/streaming-native-workbook-core.ts',
] as const
const legacyFormulaRecalcStreamingNativeSources = [
  'packages/xlsx-formula-recalc/src/streaming-native-inspect.ts',
  'packages/xlsx-formula-recalc/src/streaming-native-recalc.ts',
  'packages/xlsx-formula-recalc/src/streaming-native-row-chain-wasm.ts',
] as const
const nativeXlsxFormulaRecalcPathBoundarySources = [
  'packages/xlsx-formula-recalc/src/index.ts',
  'packages/xlsx-formula-recalc/src/cli-api.ts',
  'packages/xlsx-formula-recalc/src/evaluator-bin.ts',
  'packages/xlsx-formula-recalc/src/evaluator-cli.ts',
  'packages/xlsx-formula-recalc/src/file-recalc.ts',
  'packages/xlsx-formula-recalc/src/types.ts',
  'packages/xlsx-formula-recalc/src/workbook-compatibility-report.ts',
] as const
const nativeXlsxLargeFileModeBoundarySources = [
  'packages/xlsx-formula-recalc/src/file-recalc.ts',
  'packages/xlsx/src/streaming-native-inspect.ts',
] as const

type JsonRecord = { readonly [key: string]: unknown }

function packageManifestPaths(): string[] {
  return packageManifestDirs.flatMap((dir) => {
    if (dir === '.') {
      return ['package.json']
    }

    return readdirSync(join(repoRoot, dir), { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => join(dir, entry.name, 'package.json'))
      .filter((path) => {
        try {
          readFileSync(join(repoRoot, path), 'utf8')
          return true
        } catch {
          return false
        }
      })
  })
}

function packageDependencySourceViolations(path: string): string[] {
  const source = readFileSync(join(repoRoot, path), 'utf8')
  return forbiddenDependencySources.filter((forbidden) => source.includes(forbidden)).map((forbidden) => `${path}: ${forbidden}`)
}

function sourceFiles(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name)
    if (entry.isDirectory()) {
      return entry.name === '__tests__' ? [] : sourceFiles(path)
    }
    return entry.isFile() && entry.name.endsWith('.ts') ? [path] : []
  })
}

function allTypeScriptFiles(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name)
    if (entry.isDirectory()) {
      return allTypeScriptFiles(path)
    }
    return entry.isFile() && entry.name.endsWith('.ts') ? [path] : []
  })
}

function relativePath(path: string): string {
  return path.slice(repoRoot.length + 1)
}

function runtimeXlsxImportViolations(path: string): string[] {
  const source = readFileSync(path, 'utf8')
  if (!/(?:^|\n)\s*import\s+\*\s+as\s+\w+\s+from\s+['"]xlsx['"]|require\(['"]xlsx['"]\)|import\(['"]xlsx['"]\)/u.test(source)) {
    return []
  }
  const relative = relativePath(path)
  return excelImportRuntimeXlsxImportAllowlist.has(relative) ? [] : [relative]
}

function hasRuntimeXlsxImport(source: string): boolean {
  return /(?:^|\n)\s*import\s+\*\s+as\s+\w+\s+from\s+['"]xlsx['"]|require\(['"]xlsx['"]\)|import\(['"]xlsx['"]\)/u.test(source)
}

function packageDependencyViolations(path: string, forbiddenDependencies: readonly string[]): string[] {
  const parsed: unknown = JSON.parse(readFileSync(join(repoRoot, path, 'package.json'), 'utf8'))
  const dependencies =
    typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed) && 'dependencies' in parsed ? parsed.dependencies : undefined
  if (typeof dependencies !== 'object' || dependencies === null || Array.isArray(dependencies)) {
    return []
  }
  return forbiddenDependencies.filter((dependency) => dependency in dependencies).map((dependency) => `${path}: ${dependency}`)
}

function isJsonRecord(value: unknown): value is JsonRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function packageManifest(path: string): JsonRecord {
  const parsed: unknown = JSON.parse(readFileSync(join(repoRoot, path, 'package.json'), 'utf8'))
  if (!isJsonRecord(parsed)) {
    throw new Error(`Invalid package manifest: ${path}`)
  }
  return parsed
}

function objectField(source: JsonRecord, field: string): JsonRecord {
  const value = source[field]
  return isJsonRecord(value) ? value : {}
}

function stringField(source: JsonRecord, field: string): string {
  const value = source[field]
  return typeof value === 'string' ? value : ''
}

function sourceImportViolations(path: string, forbiddenImports: readonly string[]): string[] {
  return sourceFiles(join(repoRoot, path, 'src')).flatMap((sourceFile) => {
    const source = readFileSync(sourceFile, 'utf8')
    return forbiddenImports
      .filter((specifier) => new RegExp(`from\\s+['"]${specifier.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')}['"]`, 'u').test(source))
      .map((specifier) => `${relativePath(sourceFile)}: ${specifier}`)
  })
}

function sourceSpecifierViolations(path: string, forbiddenSpecifiers: readonly string[]): string[] {
  const source = readFileSync(join(repoRoot, path), 'utf8')
  return forbiddenSpecifiers
    .filter((specifier) => new RegExp(`from\\s+['"]${specifier.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')}(?:/[^'"]*)?['"]`, 'u').test(source))
    .map((specifier) => `${path}: ${specifier}`)
}

function sourceContains(source: string, needle: string): boolean {
  return source.indexOf(needle) >= 0
}

describe('repository dependency policy', () => {
  it('does not pin package dependencies to the SheetJS CDN tarball', () => {
    const violations = packageManifestPaths().flatMap(packageDependencySourceViolations)

    expect(violations).toEqual([])
  })

  it('does not resolve lockfile entries from the SheetJS CDN tarball', () => {
    const lockfile = readFileSync(join(repoRoot, 'pnpm-lock.yaml'), 'utf8')
    const violations = forbiddenDependencySources.filter((forbidden) => lockfile.includes(forbidden))

    expect(violations).toEqual([])
  })

  it('keeps @bilig/excel-import runtime SheetJS imports isolated to the parser and writer boundary', () => {
    const violations = sourceFiles(join(repoRoot, 'packages/excel-import/src')).flatMap(runtimeXlsxImportViolations)

    expect(violations).toEqual([])
  })

  it('keeps native XLSX fixture generation on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxFixtureScripts.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps WorkPaper XLSX corpus proof readback on @bilig/xlsx instead of SheetJS', () => {
    const violations = [...nativeXlsxCorpusProofScripts, ...nativeXlsxCorpusProofTests].filter((path) =>
      hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')),
    )

    expect(violations).toEqual([])
  })

  it('keeps WorkPaper XLSX corpus materialization small-workbook only by default', () => {
    const corpus = readFileSync(join(repoRoot, 'scripts/check-workpaper-xlsx-corpus.ts'), 'utf8')
    const corpusCli = readFileSync(join(repoRoot, 'scripts/workpaper-xlsx-corpus-cli.ts'), 'utf8')
    const recentComplex = readFileSync(join(repoRoot, 'scripts/public-workbook-corpus-recent-complex.ts'), 'utf8')

    expect(corpus).toContain('const defaultMaxFileBytes = 1_000_000')
    expect(corpus).toContain('options.allowLargeWorkPaperMaterialization !== true')
    expect(corpus).toContain('--allow-large-workpaper-materialization')
    expect(corpusCli).toContain("const allowLargeWorkPaperMaterializationFlag = '--allow-large-workpaper-materialization'")
    expect(corpusCli).toContain('Default: 1000000.')
    expect(recentComplex).toContain('const defaultWorkPaperHeadlessMaxFileBytes = 1_000_000')
    expect(recentComplex).toContain("readFlagArg('--allow-large-workpaper-materialization')")
    expect(recentComplex).not.toContain("maxFileBytes: readNumberArg('--max-file-bytes', 50 * 1024 * 1024)")
  })

  it('keeps public workbook corpus test fixtures on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxPublicWorkbookCorpusTests.filter((path) =>
      hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')),
    )

    expect(violations).toEqual([])
  })

  it('keeps recalculation bridge examples on @bilig/xlsx instead of SheetJS fixture edits', () => {
    const violations = nativeXlsxExampleScripts.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps Excel oracle fixture builders on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxExcelFixtureTests.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps app workbook import fixture builders on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxAppFixtureTests.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps source-preserving headless XLSX fixture builders on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxHeadlessFixtureTests.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps simple excel-import fixture builders on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxExcelImportFixtureTests.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps excel-import tests free of static SheetJS imports', () => {
    const violations = allTypeScriptFiles(join(repoRoot, 'packages/excel-import/src/__tests__'))
      .map(relativePath)
      .filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps XLSX formula recalculation on @bilig/xlsx instead of direct ZIP or SheetJS dependencies', () => {
    const dependencyViolations = nativeXlsxFormulaRecalcPackages.flatMap((path) => packageDependencyViolations(path, ['xlsx', 'fflate']))
    const importViolations = nativeXlsxFormulaRecalcPackages.flatMap((path) => sourceImportViolations(path, ['xlsx', 'fflate']))

    expect([...dependencyViolations, ...importViolations]).toEqual([])
  })

  it('keeps streaming-native XLSX recalculation implementation owned by @bilig/xlsx', () => {
    const missingNativeSources = xlsxOwnedStreamingNativeSources.filter((path) => !existsSync(join(repoRoot, path)))
    const staleRecalcSources = legacyFormulaRecalcStreamingNativeSources.filter((path) => existsSync(join(repoRoot, path)))
    const recalcIndex = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/index.ts'), 'utf8')
    const nativeRecalc = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-recalc.ts'), 'utf8')

    expect(missingNativeSources).toEqual([])
    expect(staleRecalcSources).toEqual([])
    expect(recalcIndex).toContain("from '@bilig/xlsx'")
    expect(recalcIndex).not.toContain("from './streaming-native-recalc.js'")
    expect(recalcIndex).not.toContain("from './streaming-native-inspect.js'")
    expect(nativeRecalc).not.toContain("'workpaper'")
  })

  it('keeps public cache-inspection API docs on file-backed streaming-native inspection', () => {
    const recalcIndex = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/index.ts'), 'utf8')
    const readmeViolations = nativeXlsxFormulaRecalcReadmes.filter((path) => {
      const source = readFileSync(join(repoRoot, path), 'utf8')
      return !source.includes('inspectXlsxCacheFile') || source.includes('inspectXlsxCache(await readFile')
    })

    expect(recalcIndex).toContain('inspectXlsxCacheFileStreamingNative as inspectXlsxCacheFile')
    expect(recalcIndex).not.toContain('inspectXlsxCache } from')
    expect(readmeViolations).toEqual([])
  })

  it('keeps workbook compatibility file reports on file-backed native scans', () => {
    const reportSource = readFileSync(join(repoRoot, 'packages/xlsx/src/workbook-compatibility-report.ts'), 'utf8')
    const cliStart = reportSource.indexOf('export function runWorkbookCompatibilityReportCli')
    const cliEnd = reportSource.indexOf('export function buildWorkbookCompatibilityDemoBytes')
    const cliSource = reportSource.slice(cliStart, cliEnd)

    expect(cliStart).toBeGreaterThan(-1)
    expect(cliEnd).toBeGreaterThan(cliStart)
    expect(reportSource).toContain('const workbookCompatibilityReportBytesApiLimit = 1_000_000')
    expect(reportSource).toContain("assertWorkbookCompatibilityReportBytesApiWithinLimit(bytes, 'buildWorkbookCompatibilityReport')")
    expect(reportSource).toContain('function readExternalWorkbookReferenceInputs(')
    expect(reportSource).toContain('const emptyExternalWorkbookReferenceBytes = new Uint8Array(0)')
    expect(reportSource).toContain('statSync(workbook.path)')
    expect(reportSource).not.toContain('bytes: readFileSync(workbook.path)')
    expect(reportSource).not.toContain('assertXlsxExternalWorkbookByteInputWithinLimit(byteLength, workbook.path)')
    expect(reportSource).toContain('export function buildWorkbookCompatibilityReportFromFile')
    expect(reportSource).toContain('openStreamingNativeWorkbookCore')
    expect(reportSource).toContain('scanStreamingNativeWorkbookCellStats(core)')
    expect(reportSource).toContain('scanStreamingNativeWorkbookPackageParts(core)')
    expect(reportSource).toContain('readXlsxFormulaCacheCellsFromWorkbookCore')
    expect(reportSource).toContain('const defaultFileInspectLimit: XlsxCacheInspectionLimit = 2000')
    expect(cliSource).toContain('buildWorkbookCompatibilityReportFromFile')
    expect(cliSource).not.toContain('readFileSync(requireInputPath(options))')
    expect(cliSource).not.toContain('readXlsxWorkbookCells')
  })

  it('keeps large XLSX file-mode diagnostics bounded by default', () => {
    const cliApi = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/cli-api.ts'), 'utf8')
    const formulaCacheReader = readFileSync(join(repoRoot, 'packages/xlsx/src/formula-cache-reader.ts'), 'utf8')
    const nativeInspect = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-inspect.ts'), 'utf8')
    const reportSource = readFileSync(join(repoRoot, 'packages/xlsx/src/workbook-compatibility-report.ts'), 'utf8')
    const nativePublicCorpus = readFileSync(join(repoRoot, 'scripts/xlsx-native-recalc-public-corpus.ts'), 'utf8')

    expect(cliApi).toContain('const defaultInspectFormulaLimit = 2000')
    expect(cliApi).not.toContain("const defaultInspectFormulaLimit = 'all'")
    expect(formulaCacheReader).toContain('export const defaultXlsxFormulaCacheInspectionLimit: XlsxFormulaCacheInspectionLimit = 2000')
    expect(formulaCacheReader).toContain('options.inspectLimit ?? defaultXlsxFormulaCacheInspectionLimit')
    expect(formulaCacheReader).not.toContain("options.inspectLimit ?? 'all'")
    expect(nativeInspect).toContain('defaultXlsxFormulaCacheInspectionLimit')
    expect(nativeInspect).toContain('export const defaultStreamingNativeXlsxCacheInspectionLimit: StreamingNativeXlsxCacheInspectionLimit')
    expect(nativeInspect).not.toContain("options.inspectLimit ?? 'all'")
    expect(reportSource).toContain('const defaultFileInspectLimit: XlsxCacheInspectionLimit = 2000')
    expect(nativePublicCorpus).toContain("readXlsxFormulaCacheCellsFromFile(inputPath, { inspectLimit: 'all' })")
  })

  it('keeps file-backed native inspection and reports on the unified workbook core', () => {
    const formulaCacheReader = readFileSync(join(repoRoot, 'packages/xlsx/src/formula-cache-reader.ts'), 'utf8')
    const reportSource = readFileSync(join(repoRoot, 'packages/xlsx/src/workbook-compatibility-report.ts'), 'utf8')
    const nativeRecalc = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-recalc.ts'), 'utf8')
    const nativeCore = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-workbook-core.ts'), 'utf8')

    expect(formulaCacheReader).toContain("from './streaming-native-workbook-core.js'")
    expect(formulaCacheReader).toContain('readXlsxFormulaCacheCellsFromWorkbookCore')
    expect(reportSource).toContain("from './streaming-native-workbook-core.js'")
    expect(reportSource).toContain('scanStreamingNativeWorkbookCellStats')
    expect(reportSource).toContain('scanStreamingNativeWorkbookPackageParts')
    expect(reportSource).toContain('readXlsxFormulaCacheCellsFromWorkbookCore')
    expect(reportSource).not.toContain('readXlsxFormulaCacheCellsFromFile(inputPath')
    expect(reportSource).not.toContain('function scanWorkbookCellStats')
    expect(reportSource).not.toContain('function scanWorkbookPackageParts')
    expect(reportSource).not.toContain('forEachInflatedXlsxZipEntryChunk')
    expect(nativeRecalc).toContain("from './streaming-native-workbook-core.js'")
    expect(nativeRecalc).not.toContain("from './file-source.js'")
    expect(nativeRecalc).not.toContain('readXlsxZipEntriesLazyFromByteSource')
    expect(nativeRecalc).not.toContain('workbookSheetPathEntriesForSource')
    expect(nativeCore).toContain('createFileXlsxSourceReader')
    expect(nativeCore).toContain('readXlsxZipEntriesLazyFromByteSource')
    expect(nativeCore).toContain('workbookSheetPathEntriesForSource')
    expect(nativeCore).toContain('export function scanStreamingNativeWorkbookCellStats')
    expect(nativeCore).toContain('export function scanStreamingNativeWorkbookPackageParts')
  })

  it('keeps native recalc row values in the typed columnar arena', () => {
    const recalcSource = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-recalc.ts'), 'utf8')
    const rowChainWasm = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-row-chain-wasm.ts'), 'utf8')
    const lookupWasm = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-lookup-wasm.ts'), 'utf8')
    const cellArena = readFileSync(join(repoRoot, 'packages/xlsx/src/streaming-native-cell-arena.ts'), 'utf8')

    expect(recalcSource).toContain('StreamingNativeSheetCellArena')
    expect(recalcSource).not.toContain('rows: Map<number, Map<number, PendingCellValue>>')
    expect(recalcSource).not.toContain('new Map<number, PendingCellValue>()')
    expect(rowChainWasm).not.toContain('Map<number, PendingCellValue>')
    expect(lookupWasm).not.toContain('Map<number, PendingCellValue>')
    expect(cellArena).toContain('Int32Array')
    expect(cellArena).toContain('Uint8Array')
    expect(cellArena).toContain('Float64Array')
    expect(cellArena).toContain('valueTags')
    expect(cellArena).toContain('stringIds')
  })

  it('keeps source-preserving file output on the streaming ZIP writer', () => {
    const sourcePreservingPatches = readFileSync(join(repoRoot, 'packages/xlsx/src/source-preserving-literal-patches.ts'), 'utf8')
    const byteExportStart = sourcePreservingPatches.indexOf('export function exportXlsxSourceLiteralPatches(')
    const syncFileStart = sourcePreservingPatches.indexOf('export function exportXlsxSourceLiteralPatchesToFile(')
    const asyncFileStart = sourcePreservingPatches.indexOf('export function exportXlsxSourceLiteralPatchesToFileAsync(')
    const streamingFileStart = sourcePreservingPatches.indexOf('function exportXlsxSourceLiteralPatchesToFileStreaming(')
    const readBytesGuardIndex = sourcePreservingPatches.indexOf(
      "assertSourcePreservingLiteralPatchBytesApiWithinLimit(source, 'source-preserving readBytes fallback')",
    )
    const readBytesFallbackIndex = sourcePreservingPatches.indexOf('return readXlsxZipEntries(source.readBytes())')
    const byteExportSource = sourcePreservingPatches.slice(byteExportStart, syncFileStart)
    const syncFileSource = sourcePreservingPatches.slice(syncFileStart, asyncFileStart)
    const streamingFileSource = sourcePreservingPatches.slice(streamingFileStart)

    expect(byteExportStart).toBeGreaterThan(-1)
    expect(syncFileStart).toBeGreaterThan(-1)
    expect(asyncFileStart).toBeGreaterThan(syncFileStart)
    expect(streamingFileStart).toBeGreaterThan(asyncFileStart)
    expect(sourcePreservingPatches).toContain('const sourcePreservingLiteralPatchBytesApiLimit = 1_000_000')
    expect(byteExportSource).toContain("assertSourcePreservingLiteralPatchBytesApiWithinLimit(source, 'exportXlsxSourceLiteralPatches')")
    expect(readBytesGuardIndex).toBeGreaterThan(-1)
    expect(readBytesFallbackIndex).toBeGreaterThan(readBytesGuardIndex)
    expect(syncFileSource).toContain('exportXlsxSourceLiteralPatchesToFileStreaming(input)')
    expect(syncFileSource).not.toContain('exportXlsxSourceLiteralPatches(input)')
    expect(syncFileSource).not.toContain('writeAllSync(fd, exported)')
    expect(streamingFileSource).toContain('zipSourcePreservingEntriesToFile')
    expect(streamingFileSource).toContain('tryPrepareStreamingPatchedWorksheetEntryFile')
  })

  it('keeps file-backed XLSX source readBytes small-workbook only', () => {
    const fileSource = readFileSync(join(repoRoot, 'packages/xlsx/src/file-source.ts'), 'utf8')
    const readBytesStart = fileSource.indexOf('readBytes()')
    const readBytesSource = fileSource.slice(readBytesStart)
    const guardIndex = readBytesSource.indexOf('assertReadBytesWithinLimit()')
    const materializeIndex = readBytesSource.indexOf('return readFileSync(path)')

    expect(fileSource).toContain('export const defaultFileXlsxSourceReadBytesLimit = 1_000_000')
    expect(fileSource).toContain('Use readRange/readRangeInto or a file-backed native XLSX API for large workbooks.')
    expect(readBytesStart).toBeGreaterThan(-1)
    expect(guardIndex).toBeGreaterThan(-1)
    expect(materializeIndex).toBeGreaterThan(guardIndex)
  })

  it('keeps excel-import source-preserving output as a shim over @bilig/xlsx', () => {
    const sourcePreservingExport = readFileSync(join(repoRoot, 'packages/excel-import/src/xlsx-source-preserving-export.ts'), 'utf8')

    expect(sourcePreservingExport).toContain("from '@bilig/xlsx'")
    expect(sourcePreservingExport).toContain('exportBiligXlsxSourceLiteralPatches')
    expect(sourcePreservingExport).toContain('forceWorkbookRecalculation: calculationTextPatches.length === 0')
    expect(sourcePreservingExport).not.toContain('tryWriteStreamingPatchedWorksheetEntry')
    expect(sourcePreservingExport).not.toContain('tryWriteNativeStreamingPatchedWorksheetEntry')
    expect(sourcePreservingExport).not.toContain('zipSourcePreservingEntriesToFile')
    expect(sourcePreservingExport).not.toContain('unzipSync(source.readBytes())')
  })

  it('keeps large XLSX import preflight on @bilig/xlsx workbook path resolution', () => {
    const largeSimpleImport = readFileSync(join(repoRoot, 'packages/excel-import/src/xlsx-large-simple-import.ts'), 'utf8')
    const largeSimpleInspect = readFileSync(join(repoRoot, 'packages/excel-import/src/xlsx-large-simple-headless-inspect.ts'), 'utf8')
    const largeSimpleWorkbookMetadata = readFileSync(
      join(repoRoot, 'packages/excel-import/src/xlsx-large-simple-workbook-metadata.ts'),
      'utf8',
    )

    expect(largeSimpleImport).toContain("from '@bilig/xlsx'")
    expect(largeSimpleImport).toContain('workbookSheetPathEntriesForSource(zip)')
    expect(largeSimpleImport).not.toContain('readWorksheetPathsByRelationshipId')
    expect(largeSimpleInspect).toContain("from '@bilig/xlsx'")
    expect(largeSimpleInspect).toContain('workbookSheetPathEntriesForSource(zip)')
    expect(largeSimpleInspect).not.toContain('readWorksheetPathsByRelationshipId')
    expect(largeSimpleWorkbookMetadata).not.toContain('readWorksheetPathsByRelationshipId')
    expect(largeSimpleWorkbookMetadata).not.toContain('worksheetRelationshipType')
  })

  it('keeps SheetJS fallback import bounded before source-byte materialization', () => {
    const limitsSource = readFileSync(join(repoRoot, 'packages/excel-import/src/xlsx-import-limits.ts'), 'utf8')
    const publicImportSource = readFileSync(join(repoRoot, 'packages/excel-import/src/index.ts'), 'utf8')
    const byteSourceImportSource = readFileSync(join(repoRoot, 'packages/excel-import/src/xlsx-byte-source-import.ts'), 'utf8')
    const sourceBytesSource = readFileSync(join(repoRoot, 'packages/excel-import/src/xlsx-source-bytes.ts'), 'utf8')
    const publicImportStart = publicImportSource.indexOf('export function importXlsx(')
    const publicImportEnd = publicImportSource.indexOf('export function importXlsm(')
    const publicImport = publicImportSource.slice(publicImportStart, publicImportEnd)
    const publicGuardIndex = publicImport.indexOf(
      'assertXlsxSheetJsFallbackWithinMaterializationLimits(inspection, options, sourceByteLength)',
    )
    const publicReaderMaterializationIndex = publicImport.indexOf('spooledUntouchedExportSource.readBytes()')
    const publicZipSourceMaterializationIndex = publicImport.indexOf('readLazyXlsxZipSource(workbookZip)')
    const byteSourceGuardIndex = byteSourceImportSource.indexOf(
      'assertXlsxSheetJsFallbackWithinMaterializationLimits(inspection, options, sourceByteLength)',
    )
    const byteSourceFallbackIndex = byteSourceImportSource.indexOf(
      'return importXlsxFromMaterializedSource(source, fileName, options)',
      byteSourceGuardIndex,
    )
    const byteSourceReadIndex = byteSourceImportSource.indexOf('const data = readAllSourceBytes(source)')

    expect(publicImportStart).toBeGreaterThan(-1)
    expect(publicImportEnd).toBeGreaterThan(publicImportStart)
    expect(limitsSource).toContain('maxMaterializedSourceBytes: denseSheetJsByteThreshold')
    expect(limitsSource).toContain("reason: 'source-byte-count'")
    expect(limitsSource).toContain('assertXlsxSourceWithinMaterializationLimits(sourceByteLength, limits)')
    expect(limitsSource).toContain('allowLegacyLargeSheetJsFallback')
    expect(limitsSource).toContain('shouldAllowLegacyLargeSheetJsFallback(options)')
    expect(limitsSource).toContain('options.limits === false && options.allowLegacyLargeSheetJsFallback === true')
    expect(publicImport).toContain('shouldAllowLegacyLargeSheetJsFallback(options)')
    expect(byteSourceImportSource).toContain("from '@bilig/xlsx/zip-reader'")
    expect(byteSourceImportSource).not.toContain("from './xlsx-zip.js'")
    expect(publicGuardIndex).toBeGreaterThan(-1)
    expect(publicReaderMaterializationIndex).toBeGreaterThan(publicGuardIndex)
    expect(publicZipSourceMaterializationIndex).toBeGreaterThan(publicGuardIndex)
    expect(byteSourceGuardIndex).toBeGreaterThan(-1)
    expect(byteSourceFallbackIndex).toBeGreaterThan(byteSourceGuardIndex)
    expect(byteSourceReadIndex).toBeGreaterThan(byteSourceGuardIndex)
    expect(sourceBytesSource).toContain('export const defaultImportedXlsxSourceReadBytesLimit = 1_000_000')
    expect(sourceBytesSource).toContain('this.assertReadBytesWithinLimit()')
    expect(sourceBytesSource).toContain('Use readRange/readRangeInto or a file-backed native XLSX API for large workbooks.')
  })

  it('keeps unchanged source export file-backed for imported XLSX source readers', () => {
    const exportSource = readFileSync(join(repoRoot, 'packages/excel-import/src/xlsx-export.ts'), 'utf8')
    const exportToFileStart = exportSource.indexOf('export function exportXlsxToFile(')
    const exportToFileEnd = exportSource.indexOf('const importedSourceCopyChunkSize')
    const exportToFileSource = exportSource.slice(exportToFileStart, exportToFileEnd)

    expect(exportSource).toContain('const sourcePreservingExportFallbackBytesLimit = 1_000_000')
    expect(exportSource).toContain("assertSourcePreservingExportFallbackWithinSmallWorkbookLimit(importedSource, 'exportXlsx')")
    expect(exportSource).toContain("assertSourcePreservingExportFallbackWithinSmallWorkbookLimit(importedSource, 'exportXlsxToFile')")
    expect(exportSource).toContain('cannot fall back to full XLSX snapshot export for a large imported source')
    expect(exportToFileStart).toBeGreaterThan(-1)
    expect(exportToFileEnd).toBeGreaterThan(exportToFileStart)
    expect(exportToFileSource).toContain('tryCopyImportedXlsxSourceToFile(importedSource, outputPath)')
    expect(exportToFileSource).toContain('if (patches.length > 0)')
    expect(exportToFileSource).toContain('} else {')
    expect(exportSource).toContain('const importedSourceCopyChunkSize = 1024 * 1024')
    expect(exportSource).toContain('const chunk = source.readRange(offset, end)')
    expect(exportSource).toContain('writeSync(fd, chunk, chunkOffset, chunk.byteLength - chunkOffset)')
  })

  it('keeps the native recalc public corpus script as a hard 50-workbook gate', () => {
    const manifest = packageManifest('.')
    const scripts = objectField(manifest, 'scripts')
    const script = stringField(scripts, 'xlsx-native-recalc:public-corpus')

    expect(script).toContain('bun scripts/xlsx-native-recalc-public-corpus.ts')
    expect(script).toContain('--limit 50')
    expect(script).toContain('--max-rss-mb 350')
    expect(script).toContain('--require-formula-workbook-count 50')
    expect(script).toContain('--require-passed-formula-workbook-count 50')
    expect(script).toContain('--require-passed')
    expect(script).toContain('--corpus .cache/public-workbook-corpus/manifest.json .cache/public-workbook-corpus')
    expect(script).toContain('--corpus .cache/public-workbook-corpus-financial/manifest.json .cache/public-workbook-corpus-financial')
    expect(script).not.toContain('--dry-run')
  })

  it('keeps the issue 442 memory gate as an exact required OCHA gate', () => {
    const manifest = packageManifest('.')
    const scripts = objectField(manifest, 'scripts')
    const script = stringField(scripts, 'xlsx-native-recalc:issue-442-gate')

    expect(script).toContain('bun scripts/xlsx-native-recalc-memory-gate.ts')
    expect(script).toContain('--issue-442-path .cache/issue-442/ocha-operational-partners-presence-jan-sep-2024.xlsx')
    expect(script).toContain('--require-issue-442')
    expect(script).toContain('--issue-442-only')
    expect(script).not.toContain('--synthetic-only')
  })

  it('keeps large XLSX file-mode wrappers off materialized workbook and hidden fallback paths', () => {
    const reportSource = readFileSync(join(repoRoot, 'packages/xlsx/src/workbook-compatibility-report.ts'), 'utf8')
    const reportFileStart = reportSource.indexOf('export function buildWorkbookCompatibilityReportFromFile')
    const reportFileEnd = reportSource.indexOf('function buildWorkbookCompatibilityReportFromScans')
    const reportFileSource = reportSource.slice(reportFileStart, reportFileEnd)
    const violations = nativeXlsxLargeFileModeBoundarySources.flatMap((path) => {
      const source = readFileSync(join(repoRoot, path), 'utf8')
      return [
        ['readFileSync(requireInputPath(options))', 'input XLSX readFileSync materialization'],
        ['source.readBytes()', 'source readBytes fallback'],
        ['readXlsxWorkbookCells(', 'full workbook cell materialization'],
        ['WorkPaper.buildFromSnapshot', 'WorkPaper snapshot materialization'],
        ["import('./legacy-workpaper.js')", 'legacy WorkPaper dynamic fallback'],
        ["from 'xlsx'", 'SheetJS static import'],
        ['require("xlsx")', 'SheetJS runtime require'],
        ["require('xlsx')", 'SheetJS runtime require'],
      ]
        .filter(([needle]) => sourceContains(source, needle))
        .map(([_needle, reason]) => `${path}: ${reason}`)
    })
    const reportFileViolations = [
      ['readFileSync(requireInputPath(options))', 'input XLSX readFileSync materialization'],
      ['source.readBytes()', 'source readBytes fallback'],
      ['readXlsxWorkbookCells(', 'full workbook cell materialization'],
      ['WorkPaper.buildFromSnapshot', 'WorkPaper snapshot materialization'],
      ["import('./legacy-workpaper.js')", 'legacy WorkPaper dynamic fallback'],
      ["from 'xlsx'", 'SheetJS static import'],
      ['require("xlsx")', 'SheetJS runtime require'],
      ["require('xlsx')", 'SheetJS runtime require'],
    ]
      .filter(([needle]) => sourceContains(reportFileSource, needle))
      .map(([_needle, reason]) => `packages/xlsx/src/workbook-compatibility-report.ts buildWorkbookCompatibilityReportFromFile: ${reason}`)

    expect(reportFileStart).toBeGreaterThan(-1)
    expect(reportFileEnd).toBeGreaterThan(reportFileStart)
    expect(reportFileSource).toContain('createWorkbookCompatibilityRssRecorder(options.maxRssBytes)')
    expect(reportFileSource).toContain("rss.recordPhase('file-api:open-core')")
    expect(reportSource).toContain('function buildWorkbookCompatibilityNativeDiagnostics')
    expect(reportSource).toContain("engineMode: 'streaming-native'")
    expect(reportSource).toContain('fallbackUsed: false')
    expect([...violations, ...reportFileViolations]).toEqual([])
  })

  it('keeps file-backed xlsx-recalc CLI on @bilig/xlsx without primary WorkPaper fallback', () => {
    const cliApi = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/cli-api.ts'), 'utf8')
    const fileRecalc = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/file-recalc.ts'), 'utf8')
    const scopedIndex = readFileSync(join(repoRoot, 'packages/bilig-xlsx-formula-recalc/src/index.ts'), 'utf8')
    const sheetjsAdapter = readFileSync(join(repoRoot, 'packages/sheetjs-formula-recalc/src/index.ts'), 'utf8')
    const exceljsAdapter = readFileSync(join(repoRoot, 'packages/exceljs-formula-recalc/src/index.ts'), 'utf8')
    const sheetjsLegacyHelperExport = sheetjsAdapter.split('\n').find((line) => line.includes("from 'bilig-workpaper/xlsx'")) ?? ''
    const exceljsLegacyHelperExport = exceljsAdapter.split('\n').find((line) => line.includes("from 'bilig-workpaper/xlsx'")) ?? ''
    const unscopedManifest = packageManifest('packages/xlsx-formula-recalc')
    const scopedManifest = packageManifest('packages/bilig-xlsx-formula-recalc')
    const externalWorkbookGuardIndex = cliApi.indexOf('assertXlsxExternalWorkbookByteInputWithinLimit(byteLength, workbook.path)')
    const externalWorkbookReadIndex = cliApi.indexOf('bytes: readFileSync(workbook.path)')

    expect(cliApi).toContain("from './file-recalc.js'")
    expect(cliApi).toContain("from '@bilig/xlsx'")
    expect(cliApi).toContain('File-mode XLSX byte input is disabled')
    expect(cliApi).not.toContain('return readFileSync(requireInputPath(options))')
    expect(externalWorkbookGuardIndex).toBeGreaterThan(-1)
    expect(externalWorkbookReadIndex).toBeGreaterThan(externalWorkbookGuardIndex)
    expect(cliApi).not.toMatch(/from\s+['"]\.\/index\.js['"]/u)
    expect(cliApi).not.toContain("import('./legacy-workpaper.js')")
    expect(fileRecalc).toContain("from '@bilig/xlsx'")
    expect(fileRecalc).not.toContain("import('./legacy-workpaper.js')")
    expect(scopedIndex).not.toContain('legacy-workpaper')
    expect(objectField(unscopedManifest, 'exports')).not.toHaveProperty('./legacy-workpaper')
    expect(objectField(scopedManifest, 'exports')).not.toHaveProperty('./legacy-workpaper')
    expect(sheetjsAdapter).toContain("from 'xlsx-formula-recalc'")
    expect(sheetjsLegacyHelperExport).not.toContain('recalculateXlsx')
    expect(sheetjsAdapter).not.toContain('xlsx-formula-recalc/legacy-workpaper')
    expect(exceljsAdapter).toContain("from 'xlsx-formula-recalc'")
    expect(exceljsLegacyHelperExport).not.toContain('recalculateXlsx')
    expect(exceljsAdapter).not.toContain('xlsx-formula-recalc/legacy-workpaper')
  })

  it('keeps primary xlsx-formula-recalc option types native-only', () => {
    const primaryTypes = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/types.ts'), 'utf8')
    const legacyWorkPaper = readFileSync(join(repoRoot, 'packages/bilig/src/xlsx-recalc.ts'), 'utf8')

    expect(primaryTypes).toContain("export type XlsxFormulaRecalcEngineMode = 'streaming-native'")
    expect(primaryTypes).toContain("export type XlsxFormulaRecalcFallbackPolicy = 'error'")
    expect(primaryTypes).not.toContain("'workpaper'")
    expect(primaryTypes).not.toContain('XlsxFormulaRecalcWorkPaperConfig')
    expect(primaryTypes).not.toContain('readonly config?:')
    expect(existsSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/legacy-workpaper.ts'))).toBe(false)
    expect(legacyWorkPaper).toContain("export type XlsxFormulaRecalcWorkPaperEngine = 'auto' | 'workpaper'")
    expect(legacyWorkPaper).toContain('export interface XlsxFormulaRecalcWorkPaperConfig')
  })

  it('keeps bilig-workpaper/xlsx public file-to-file recalc on @bilig/xlsx streaming-native', () => {
    const legacyWorkPaper = readFileSync(join(repoRoot, 'packages/bilig/src/xlsx-recalc.ts'), 'utf8')
    const publicXlsxBarrel = readFileSync(join(repoRoot, 'packages/bilig/src/xlsx.ts'), 'utf8')
    const fileToFileStart = legacyWorkPaper.indexOf('export async function recalculateXlsxFileToFile')
    const fileToFileEnd = legacyWorkPaper.indexOf('export function recalculateXlsxToFile')
    const fileToFileSource = legacyWorkPaper.slice(fileToFileStart, fileToFileEnd)

    expect(fileToFileStart).toBeGreaterThan(-1)
    expect(fileToFileEnd).toBeGreaterThan(fileToFileStart)
    expect(fileToFileSource).toContain('recalculateXlsxFileToFileStreamingNative')
    expect(fileToFileSource).toContain("legacyOptions.engine === 'workpaper'")
    expect(fileToFileSource).not.toContain('WorkPaper.buildFromSnapshot')
    expect(fileToFileSource).not.toContain('withPreparedRecalculatedXlsxOutput')
    expect(publicXlsxBarrel).toContain('recalculateXlsxFileToFile')
  })

  it('keeps bilig-workpaper legacy bytes recalc small-workbook only', () => {
    const legacyWorkPaper = readFileSync(join(repoRoot, 'packages/bilig/src/xlsx-recalc.ts'), 'utf8')

    expect(legacyWorkPaper).toContain('const legacyWorkPaperBytesApiLimit = 1_000_000')
    expect(legacyWorkPaper).toContain("assertLegacyWorkPaperBytesApiWithinLimit(bytes, 'recalculateXlsx')")
    expect(legacyWorkPaper).toContain("assertLegacyWorkPaperBytesApiWithinLimit(bytes, 'recalculateXlsxToFile')")
    expect(legacyWorkPaper).toContain("assertLegacyWorkPaperBytesApiWithinLimit(bytes, 'inspectXlsxCache')")
    expect(legacyWorkPaper).toContain('Use recalculateXlsxFileToFile() for file-backed streaming-native XLSX jobs')
  })

  it('keeps headless MCP XLSX file import on the file-backed importer boundary', () => {
    const mcpXlsxFile = readFileSync(join(repoRoot, 'packages/headless/src/work-paper-mcp-xlsx-file.ts'), 'utf8')
    const guardIndex = mcpXlsxFile.indexOf('assertWorkPaperMcpXlsxImportWithinSmallWorkbookLimit(xlsxPath)')
    const importIndex = mcpXlsxFile.indexOf('importXlsxFile(xlsxPath')
    const workPaperIndex = mcpXlsxFile.indexOf('WorkPaper.buildFromSnapshot')

    expect(mcpXlsxFile).toContain("import { importXlsxFile } from './xlsx.js'")
    expect(mcpXlsxFile).toContain('const workPaperMcpFromXlsxBytesLimit = 1_000_000')
    expect(mcpXlsxFile).toContain('statSync(xlsxPath).size')
    expect(mcpXlsxFile).toContain('small-workbook WorkPaper materialization path')
    expect(mcpXlsxFile).toContain('importXlsxFile(xlsxPath')
    expect(guardIndex).toBeGreaterThan(-1)
    expect(importIndex).toBeGreaterThan(guardIndex)
    expect(workPaperIndex).toBeGreaterThan(guardIndex)
    expect(mcpXlsxFile).not.toContain('readFileSync(xlsxPath)')
    expect(mcpXlsxFile).not.toContain('importXlsx(new Uint8Array')
  })

  it('keeps formula clinic large-file mode on native preflight before WorkPaper import', () => {
    const formulaClinic = readFileSync(join(repoRoot, 'packages/headless/src/formula-clinic-cli.ts'), 'utf8')
    const largeFileGuardIndex = formulaClinic.indexOf('fileSizeBytes > formulaClinicLegacyWorkPaperBytesApiLimit')
    const readBytesIndex = formulaClinic.indexOf('new Uint8Array(readFileSync(filePath))')
    const workPaperIndex = formulaClinic.indexOf('WorkPaper.buildFromSnapshot')

    expect(formulaClinic).toContain('const formulaClinicLegacyWorkPaperBytesApiLimit = 1_000_000')
    expect(formulaClinic).toContain("from '@bilig/xlsx/workbook-compatibility-report'")
    expect(formulaClinic).toContain('buildWorkbookCompatibilityReportFromFile(input.filePath')
    expect(formulaClinic).toContain("status: 'native-preflight'")
    expect(formulaClinic).toContain('maxObservedRssBytes: report.diagnostics.maxObservedRssBytes')
    expect(formulaClinic).toContain('phaseRssPeaks: report.diagnostics.phaseRssPeaks')
    expect(formulaClinic).toContain('formulaCounts: report.diagnostics.formulaCounts')
    expect(formulaClinic).toContain('patchedCacheCount: report.diagnostics.patchedCacheCount')
    expect(formulaClinic).toContain('unsupportedReason: report.diagnostics.unsupportedReason')
    expect(largeFileGuardIndex).toBeGreaterThan(-1)
    expect(readBytesIndex).toBeGreaterThan(largeFileGuardIndex)
    expect(workPaperIndex).toBeGreaterThan(readBytesIndex)
  })

  it('keeps public corpus workers on native file-backed scanners before materialized fallback', () => {
    const workerCommands = readFileSync(join(repoRoot, 'scripts/public-workbook-corpus-worker-commands.ts'), 'utf8')
    const workbookHelpers = readFileSync(join(repoRoot, 'scripts/public-workbook-corpus-workbook.ts'), 'utf8')
    const publicCorpusProofHelpers = [
      'scripts/public-workbook-corpus-large-simple-compact.ts',
      'scripts/public-workbook-corpus-verify.ts',
      'scripts/public-workbook-corpus-verify-worker.ts',
      'scripts/public-workbook-corpus-worker-commands.ts',
      'scripts/public-workbook-corpus-workbook.ts',
      'scripts/public-workbook-corpus-xlsx-byte-source.ts',
      'scripts/public-workbook-corpus-xlsx-footprint.ts',
      'scripts/public-workbook-corpus-xlsx-worksheet-footprint.ts',
    ]
    const fallbackGuardIndex = workerCommands.indexOf('assertMaterializedWorkbookFallbackWithinLimit(filePath')
    const fingerprintReadIndex = workerCommands.indexOf('fingerprintWorkbookBytes(readFileSync(filePath)')
    const footprintReadIndex = workerCommands.indexOf('const bytes = filePath ? readFileSync(filePath) : readFileSync(0)')

    expect(workerCommands).toContain('const publicWorkbookCorpusWorkerMaterializedBytesFallbackLimit = 1_000_000')
    expect(workerCommands).toContain("from '@bilig/xlsx/formula-cache-reader'")
    expect(workerCommands).toContain("from '@bilig/xlsx/workbook-compatibility-report'")
    expect(workerCommands).toContain("from '@bilig/xlsx/zip-reader'")
    expect(workbookHelpers).toContain("from '@bilig/xlsx/zip-reader'")
    expect(workerCommands).toContain('tryFingerprintFormulaWorkbookFromFile(filePath, fileName)')
    expect(workerCommands).toContain('tryInspectNativeCompatibilityFootprintFromFile(filePath, fileName)')
    expect(fallbackGuardIndex).toBeGreaterThan(-1)
    expect(fingerprintReadIndex).toBeGreaterThan(fallbackGuardIndex)
    expect(footprintReadIndex).toBeGreaterThan(fallbackGuardIndex)
    expect(workbookHelpers).toContain('nativeOnly: true')
    expect(workbookHelpers).toContain('SheetJS fallback is disabled for corpus verification.')
    expect(workbookHelpers).not.toContain('createRequire(import.meta.url)')
    expect(workbookHelpers).not.toContain("requireModule('xlsx')")
    expect(workbookHelpers).not.toContain('loadOptionalSheetJs')
    for (const path of publicCorpusProofHelpers) {
      expect(readFileSync(join(repoRoot, path), 'utf8')).not.toContain('../packages/excel-import/src/xlsx-zip.js')
    }
  })

  it('keeps external XLSX stress public import small-workbook only before materialized reads', () => {
    const stressWorker = readFileSync(join(repoRoot, 'scripts/external-xlsx-memory-stress-worker.ts'), 'utf8')
    const guardIndex = stressWorker.indexOf('assertExternalXlsxStressPublicImportWithinSmallWorkbookLimit(filePath)')
    const readIndex = stressWorker.indexOf('return readFileSync(filePath)')

    expect(stressWorker).toContain('const externalXlsxStressPublicImportBytesLimit = 1_000_000')
    expect(stressWorker).toContain("from '@bilig/xlsx/zip-reader'")
    expect(stressWorker).toContain('Use default file-backed external XLSX stress mode for large workbook memory gates.')
    expect(stressWorker).not.toContain("requireModule('../packages/excel-import/src/xlsx-zip.js')")
    expect(guardIndex).toBeGreaterThan(-1)
    expect(readIndex).toBeGreaterThan(guardIndex)
  })

  it('keeps external XLSX stress source resolution file-backed for cached workbooks', () => {
    const stress = readFileSync(join(repoRoot, 'scripts/external-xlsx-memory-stress.ts'), 'utf8')
    const zipBranchIndex = stress.indexOf("if (source.fileName.toLowerCase().endsWith('.zip'))")
    const ensureCachedIndex = stress.indexOf('await ensureSourceFileCached(source, sourceCachePath, args)')
    const extractArchiveIndex = stress.indexOf(
      'extractExternalXlsxStressWorkbookEntriesFromArchiveFile(source, sourceCachePath, args.cacheDir)',
    )

    expect(stress).toContain('async function ensureSourceFileCached(')
    expect(stress).toContain("from '@bilig/xlsx'")
    expect(stress).toContain('readXlsxZipEntriesLazyFromByteSource(archiveSource)')
    expect(stress).toContain('forEachInflatedXlsxZipEntryChunkAsync(')
    expect(stress).toContain('statSync(input.path).size')
    expect(stress).toContain('hashExternalXlsxStressWorkbookFileSha256(input.path)')
    expect(stress).toContain('statSync(outputPath).size !== expectedEntryByteLength')
    expect(stress).not.toContain('const bytes = readFileSync(input.path)')
    expect(stress).not.toContain('readFileSync(outputPath).byteLength')
    expect(stress).not.toContain('readOrFetchSourceBytes')
    expect(stress).not.toContain('unzipSync')
    expect(zipBranchIndex).toBeGreaterThan(-1)
    expect(ensureCachedIndex).toBeGreaterThan(zipBranchIndex)
    expect(extractArchiveIndex).toBeGreaterThan(ensureCachedIndex)
  })

  it('keeps WorkPaper evaluator doors owned by WorkPaper packages', () => {
    const xlsxEvaluator = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/evaluator-cli.ts'), 'utf8')
    const unscopedWorkPaperBin = readFileSync(join(repoRoot, 'packages/bilig/bin/bilig-evaluate.js'), 'utf8')
    const scopedWorkPaperBin = readFileSync(join(repoRoot, 'packages/workpaper/bin/bilig-evaluate.js'), 'utf8')
    const unscopedWorkPaperEvaluator = readFileSync(join(repoRoot, 'packages/bilig/src/evaluator.ts'), 'utf8')

    expect(xlsxEvaluator).not.toContain("import('@bilig/headless')")
    expect(xlsxEvaluator).not.toContain('workpaper-service')
    expect(unscopedWorkPaperBin).toContain("await import('../dist/evaluator-bin.js')")
    expect(scopedWorkPaperBin).toContain("await import('../dist/evaluator-bin.js')")
    expect(unscopedWorkPaperBin).not.toContain('@bilig/xlsx-formula-recalc/evaluator')
    expect(scopedWorkPaperBin).not.toContain('@bilig/xlsx-formula-recalc/evaluator')
    expect(unscopedWorkPaperEvaluator).toContain("export type BiligEvaluatorDoor = 'workpaper-service' | 'agent-mcp'")
    expect(unscopedWorkPaperEvaluator).not.toContain('xlsx-cache')
  })

  it('keeps bilig-workpaper off the xlsx-formula-recalc package boundary', () => {
    const manifest = packageManifest('packages/bilig')
    const dependencies = objectField(manifest, 'dependencies')
    const scripts = objectField(manifest, 'scripts')
    const xlsxRiskTool = readFileSync(join(repoRoot, 'packages/bilig/src/work-paper-mcp-xlsx-risk-tool.ts'), 'utf8')
    const sourceViolations = sourceFiles(join(repoRoot, 'packages/bilig/src')).flatMap((sourceFile) =>
      sourceSpecifierViolations(relativePath(sourceFile), ['@bilig/xlsx-formula-recalc']),
    )

    expect(dependencies).not.toHaveProperty('@bilig/xlsx-formula-recalc')
    expect(dependencies).not.toHaveProperty('xlsx-formula-recalc')
    expect(stringField(scripts, 'build')).not.toContain('@bilig/xlsx-formula-recalc')
    expect(sourceViolations).toEqual([])
    expect(xlsxRiskTool).toContain("from '@bilig/xlsx/workbook-compatibility-report'")
    expect(xlsxRiskTool).toContain('buildWorkbookCompatibilityReportFromFile(xlsxPath')
    expect(xlsxRiskTool).toContain('const defaultXlsxWorkbookRiskInspectLimit = 2000')
    expect(xlsxRiskTool).toContain("args['inspectLimit'] ?? defaultXlsxWorkbookRiskInspectLimit")
    expect(xlsxRiskTool).not.toContain("args['inspectLimit'] ?? 'all'")
    expect(xlsxRiskTool).not.toContain('buildWorkbookCompatibilityReport(readFileSync')
    expect(xlsxRiskTool).not.toContain('readFileSync(xlsxPath)')
    expect(xlsxRiskTool).not.toContain('@bilig/xlsx-formula-recalc')
  })

  it('keeps native file recalc CLI and public file types off static headless imports', () => {
    const violations = nativeXlsxFormulaRecalcPathBoundarySources.flatMap((path) => sourceSpecifierViolations(path, ['@bilig/headless']))

    expect(violations).toEqual([])
  })

  it('keeps the xlsx-formula-recalc native package install and build path off @bilig/headless', () => {
    const manifest = packageManifest('packages/xlsx-formula-recalc')
    const dependencies = objectField(manifest, 'dependencies')
    const devDependencies = objectField(manifest, 'devDependencies')
    const peerDependencies = objectField(manifest, 'peerDependencies')
    const scripts = objectField(manifest, 'scripts')

    expect(dependencies).not.toHaveProperty('@bilig/headless')
    expect(peerDependencies).not.toHaveProperty('@bilig/headless')
    expect(devDependencies).not.toHaveProperty('@bilig/headless')
    expect(stringField(scripts, 'build')).not.toContain('@bilig/headless')
  })

  it('keeps published Bilig XLSX runtime packages free of SheetJS xlsx dependencies', () => {
    const violations = publishedNativeXlsxRuntimePackages.flatMap((path) =>
      packageDependencyViolations(path, ['xlsx', 'xlsx-js-style', '@sheetjs/xlsx']),
    )

    expect(violations).toEqual([])
  })

  it('keeps published XLSX formula recalculation CLI entrypoints on the file-backed async recalc path', () => {
    const violations = fileBackedXlsxFormulaRecalcCliEntrypoints.filter((path) => {
      const source = readFileSync(join(repoRoot, path), 'utf8')
      return !source.includes('runXlsxFormulaRecalcCliAsync') || source.includes('runXlsxFormulaRecalcCli(')
    })

    expect(violations).toEqual([])
  })
})
