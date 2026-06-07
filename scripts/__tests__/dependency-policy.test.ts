import { readFileSync, readdirSync } from 'node:fs'
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
const nativeXlsxExcelImportFixtureTests = [
  'packages/excel-import/src/__tests__/xlsx-alignment-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-byte-source-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-calculation-properties-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-data-validations-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-date1904-system.test.ts',
  'packages/excel-import/src/__tests__/xlsx-defined-names-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-error-cell-import.test.ts',
  'packages/excel-import/src/__tests__/xlsx-external-defined-names.test.ts',
  'packages/excel-import/src/__tests__/xlsx-export-large-simple.test.ts',
  'packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-formula-cache-text-normalization.test.ts',
  'packages/excel-import/src/__tests__/xlsx-hyperlink-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/excel-import-sparse-range.test.ts',
  'packages/excel-import/src/__tests__/xlsx-axis-visibility-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-properties-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-name-whitespace.test.ts',
  'packages/excel-import/src/__tests__/xlsx-sheet-visibility-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-view-state-roundtrip.test.ts',
  'packages/excel-import/src/__tests__/xlsx-workbook-sheet-paths.test.ts',
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

function sourceImportViolations(path: string, forbiddenImports: readonly string[]): string[] {
  return sourceFiles(join(repoRoot, path, 'src')).flatMap((sourceFile) => {
    const source = readFileSync(sourceFile, 'utf8')
    return forbiddenImports
      .filter((specifier) => new RegExp(`from\\s+['"]${specifier.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')}['"]`, 'u').test(source))
      .map((specifier) => `${relativePath(sourceFile)}: ${specifier}`)
  })
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

  it('keeps simple excel-import fixture builders on @bilig/xlsx instead of SheetJS', () => {
    const violations = nativeXlsxExcelImportFixtureTests.filter((path) => hasRuntimeXlsxImport(readFileSync(join(repoRoot, path), 'utf8')))

    expect(violations).toEqual([])
  })

  it('keeps XLSX formula recalculation on @bilig/xlsx instead of direct ZIP or SheetJS dependencies', () => {
    const dependencyViolations = nativeXlsxFormulaRecalcPackages.flatMap((path) => packageDependencyViolations(path, ['xlsx', 'fflate']))
    const importViolations = nativeXlsxFormulaRecalcPackages.flatMap((path) => sourceImportViolations(path, ['xlsx', 'fflate']))

    expect([...dependencyViolations, ...importViolations]).toEqual([])
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
