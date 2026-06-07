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
const nativeXlsxExampleScripts = [
  'examples/recalc-bridge-workflows/smoke.mjs',
  'examples/recalc-bridge-workflows/stackoverflow-sheetjs-63085785.mjs',
] as const
const nativeXlsxHeadlessFixtureTests = ['packages/headless/src/__tests__/work-paper-source-preserving-xlsx-export.test.ts'] as const
const nativeXlsxExcelImportFixtureTests = [
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
  'packages/excel-import/src/__tests__/xlsx-worksheet-relationship-path-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-worksheet-dimensions-roundtrip.test.ts',
] as const
const nativeXlsxFormulaRecalcPackages = ['packages/xlsx-formula-recalc', 'packages/bilig-xlsx-formula-recalc'] as const
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
  'packages/xlsx/src/streaming-native-external-cache.ts',
  'packages/xlsx/src/streaming-native-inspect.ts',
  'packages/xlsx/src/streaming-native-recalc.ts',
  'packages/xlsx/src/streaming-native-row-chain-wasm.ts',
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
    const violations = nativeXlsxCorpusProofScripts.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps recalculation bridge examples on @bilig/xlsx instead of SheetJS fixture edits', () => {
    const violations = nativeXlsxExampleScripts.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

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

  it('keeps file-backed xlsx-recalc CLI on @bilig/xlsx without primary WorkPaper fallback', () => {
    const cliApi = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/cli-api.ts'), 'utf8')
    const fileRecalc = readFileSync(join(repoRoot, 'packages/xlsx-formula-recalc/src/file-recalc.ts'), 'utf8')
    const scopedIndex = readFileSync(join(repoRoot, 'packages/bilig-xlsx-formula-recalc/src/index.ts'), 'utf8')
    const sheetjsAdapter = readFileSync(join(repoRoot, 'packages/sheetjs-formula-recalc/src/index.ts'), 'utf8')
    const exceljsAdapter = readFileSync(join(repoRoot, 'packages/exceljs-formula-recalc/src/index.ts'), 'utf8')
    const unscopedManifest = packageManifest('packages/xlsx-formula-recalc')
    const scopedManifest = packageManifest('packages/bilig-xlsx-formula-recalc')

    expect(cliApi).toContain("from './file-recalc.js'")
    expect(cliApi).toContain("from '@bilig/xlsx'")
    expect(cliApi).not.toMatch(/from\s+['"]\.\/index\.js['"]/u)
    expect(cliApi).not.toContain("import('./legacy-workpaper.js')")
    expect(fileRecalc).toContain("from '@bilig/xlsx'")
    expect(fileRecalc).not.toContain("import('./legacy-workpaper.js')")
    expect(scopedIndex).not.toContain('legacy-workpaper')
    expect(objectField(unscopedManifest, 'exports')).not.toHaveProperty('./legacy-workpaper')
    expect(objectField(scopedManifest, 'exports')).not.toHaveProperty('./legacy-workpaper')
    expect(sheetjsAdapter).toContain("from 'bilig-workpaper/xlsx'")
    expect(sheetjsAdapter).not.toContain('xlsx-formula-recalc/legacy-workpaper')
    expect(exceljsAdapter).toContain("from 'bilig-workpaper/xlsx'")
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
