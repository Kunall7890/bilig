import { readdirSync, readFileSync, statSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

interface PackageManifest {
  readonly dependencies?: Record<string, string>
  readonly devDependencies?: Record<string, string>
  readonly peerDependencies?: Record<string, string>
  readonly optionalDependencies?: Record<string, string>
  readonly exports?: Record<string, unknown>
  readonly files?: readonly string[]
}

const sourceRoot = fileURLToPath(new URL('../', import.meta.url))
const packageJsonPath = new URL('../../package.json', import.meta.url)
const readmePath = new URL('../../README.md', import.meta.url)
const examplePath = new URL('../../../../examples/workbook-agent-model/named-range-formula.ts', import.meta.url)

const bannedRuntimeDependencies = Object.freeze([
  '@bilig/core',
  '@bilig/headless',
  '@bilig/agent-api',
  '@bilig/web',
  '@bilig/grid',
  '@bilig/renderer',
  'zod',
  'effect',
] as const)

const businessExampleTerms = Object.freeze([
  'findTable',
  'findRows',
  'Item',
  'Quantity',
  'Rate',
  'Status',
  'Total',
  'ready',
  'revenue',
  'prepaid',
  'forecast',
  'quote',
] as const)

function readPackageManifest(): PackageManifest {
  const parsed: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('package.json did not parse as an object')
  }
  return parsed as PackageManifest
}

function walkSourceFiles(root: string): readonly string[] {
  const entries = readdirSync(root, { withFileTypes: true })
  const files: string[] = []
  for (const entry of entries) {
    if (entry.name === '__tests__' || entry.name.endsWith('.test.ts')) {
      continue
    }
    const path = `${root}/${entry.name}`
    if (entry.isDirectory()) {
      files.push(...walkSourceFiles(path))
      continue
    }
    if (entry.isFile() && entry.name.endsWith('.ts')) {
      files.push(path)
    }
  }
  return files
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function importPattern(packageName: string): RegExp {
  const escaped = escapeRegExp(packageName)
  return new RegExp(`(?:from\\s+['"]${escaped}(?:/[^'"]*)?['"]|import\\(\\s*['"]${escaped}(?:/[^'"]*)?['"]\\s*\\))`)
}

function shapeExample(readme: string): string {
  const sectionStart = readme.indexOf('## The Shape')
  const sectionEnd = readme.indexOf('## Which Package')
  expect(sectionStart).toBeGreaterThanOrEqual(0)
  expect(sectionEnd).toBeGreaterThan(sectionStart)

  const section = readme.slice(sectionStart, sectionEnd)
  const match = /```ts\n([\s\S]*?)\n```/.exec(section)
  expect(match?.[1]).toBeDefined()
  return match?.[1] ?? ''
}

describe('@bilig/workbook package boundary', () => {
  it('keeps runtime dependencies boring and workbook-local', () => {
    const manifest = readPackageManifest()

    expect(manifest.dependencies).toEqual({
      '@bilig/formula': 'workspace:*',
      '@bilig/protocol': 'workspace:*',
    })

    const dependencyFields = [manifest.dependencies, manifest.devDependencies, manifest.peerDependencies, manifest.optionalDependencies]
    for (const dependencies of dependencyFields) {
      const names = new Set(Object.keys(dependencies ?? {}))
      for (const banned of bannedRuntimeDependencies) {
        expect(names.has(banned)).toBe(false)
      }
    }
  })

  it('keeps public source imports out of app, engine, ui, and validator packages', () => {
    const files = walkSourceFiles(sourceRoot)
    expect(files.length).toBeGreaterThan(0)

    const importPatterns = bannedRuntimeDependencies.map((dependency) => ({
      dependency,
      pattern: importPattern(dependency),
    }))

    for (const file of files) {
      expect(statSync(file).isFile()).toBe(true)
      const source = readFileSync(file, 'utf8')
      for (const { dependency, pattern } of importPatterns) {
        expect(pattern.test(source), `${file} imports ${dependency}`).toBe(false)
      }
      expect(source.includes('apps/'), `${file} imports app code`).toBe(false)
    }
  })

  it('publishes layered subpath exports without replacing the root barrel', () => {
    const exportsMap = readPackageManifest().exports

    expect(exportsMap).toMatchObject({
      '.': expect.any(Object),
      './model': expect.any(Object),
      './find': expect.any(Object),
      './check': expect.any(Object),
      './formula': expect.any(Object),
      './verify': expect.any(Object),
      './runtime': expect.any(Object),
      './command': expect.any(Object),
      './schema': expect.any(Object),
    })

    expect(readPackageManifest().files).toContain('fixtures')
  })

  it('keeps the first README path neutral and strict-proof oriented', () => {
    const example = shapeExample(readFileSync(readmePath, 'utf8'))

    expect(example).toContain("workbook.findName('input')")
    expect(example).toContain("workbook.findName('factor')")
    expect(example).toContain("workbook.findName('result')")
    expect(example).toContain('runWorkbookPlan(transportedPlan, adapter, { strict: true })')
    for (const term of businessExampleTerms) {
      expect(example.includes(term), `README first example contains ${term}`).toBe(false)
    }
  })

  it('keeps the runnable example neutral and strict-proof oriented', () => {
    const example = readFileSync(examplePath, 'utf8')

    expect(example).toContain("workbook.findName('input')")
    expect(example).toContain("workbook.findName('factor')")
    expect(example).toContain("workbook.findName('result')")
    expect(example).toContain('runWorkbookPlan(transportedPlan, adapter, { strict: true })')
    expect(example).toContain('commandReceipts')
    expect(example).toContain('resolvedRefs')
    for (const term of businessExampleTerms) {
      expect(example.includes(term), `runnable example contains ${term}`).toBe(false)
    }
  })
})
