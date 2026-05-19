import { readdirSync, readFileSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

const requiredFeatureGroups = [
  { label: 'engine history and mutation model', patterns: ['core/history/', 'engine-history.fuzz.test.ts'] },
  { label: 'engine metadata model', patterns: ['core/metadata/', 'engine-metadata.fuzz.test.ts'] },
  { label: 'engine replica replay', patterns: ['core/replica/local-batch-replay-parity'] },
  { label: 'engine snapshot restore', patterns: ['core/snapshot/restore-roundtrip-parity'] },
  { label: 'engine structural inverse replay', patterns: ['core/structure/inverse-replay'] },
  { label: 'headless WorkPaper action history', patterns: ['headless/work-paper/action-sequence-save-load'] },
  { label: 'headless tracked cell refs', patterns: ['headless/work-paper-tracked-change-reducer'] },
  { label: 'headless core parity', patterns: ['headless/core-parity/'] },
  { label: 'formula parser', patterns: ['formula/parse/canonicalization', 'formula/parse/invalid-input'] },
  { label: 'formula evaluator', patterns: ['formula/evaluation/canonicalization-stability'] },
  { label: 'formula translation and rename', patterns: ['formula/translation/', 'formula/rename/'] },
  { label: 'formula lookup/reference family', patterns: ['formula/lookup-reference/'] },
  { label: 'formula wasm/runtime differential', patterns: ['core/formula-runtime/generated-differential'] },
  { label: 'xlsx import/export semantics', patterns: ['excel-import/xlsx/'] },
  { label: 'xlsx byte boundary', patterns: ['xlsx-import-container.mjs'] },
  { label: 'csv import semantics', patterns: ['excel-import/csv/', 'core/csv/'] },
  { label: 'csv byte boundary', patterns: ['csv-import.mjs'] },
  { label: 'binary protocol frames', patterns: ['binary-protocol/frame-roundtrip', 'binary-protocol-frame.mjs'] },
  { label: 'agent api frames', patterns: ['agent-api/frame-codec/', 'agent-stdio-frame.mjs'] },
  { label: 'contracts and protocol guards', patterns: ['contracts/schema/', 'protocol/guards/'] },
  { label: 'codex app tool normalization', patterns: ['bilig/codex-app/tool-input/'] },
  { label: 'codex app message parsing', patterns: ['bilig/codex-app/message-parsers/'] },
  { label: 'codex app verification receipts', patterns: ['bilig/codex-app/verification-status/'] },
  { label: 'web runtime reconnect', patterns: ['web/runtime-sync/reconnect-convergence'] },
  { label: 'web projected viewport cache', patterns: ['web/projected-viewport/'] },
  { label: 'web selection/app model', patterns: ['web/app-model/', 'web/selection-command/'] },
  { label: 'grid geometry and selection', patterns: ['grid/geometry/', 'grid/selection/'] },
  { label: 'grid clipboard', patterns: ['grid/clipboard/'] },
  { label: 'grid renderer cache', patterns: ['grid/renderer-v3/'] },
  { label: 'workbook runtime snapshot/session', patterns: ['bilig/workbook-runtime/'] },
  { label: 'zero sync projection', patterns: ['zero-sync/snapshot/', 'bilig/sync-relay/'] },
  {
    label: 'storage and worker transport',
    patterns: ['storage-server/in-memory-persistence-model', 'worker-transport/request-response-parity'],
  },
] as const

const forbiddenMarkers = [
  '@fuzz-browser',
  'BILIG_FUZZ_SKIP_BROWSER',
  'BILIG_BROWSER_INCLUDE_FUZZ',
  'BILIG_FUZZ_BROWSER',
  'test:fuzz:main',
  'test:fuzz:nightly',
  'test:fuzz:replay',
  'test:fuzz:promote',
] as const

describe('fuzz inventory guardrails', () => {
  it('keeps exactly one package fuzz entrypoint', () => {
    const packageJson = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8')) as unknown
    const scripts = isRecord(packageJson) && isRecord(packageJson.scripts) ? packageJson.scripts : {}

    expect(Object.keys(scripts).filter((scriptName) => scriptName === 'test:fuzz' || scriptName.startsWith('test:fuzz:'))).toEqual([
      'test:fuzz',
    ])
  })

  it('keeps browser fuzz and fuzz variants out of the repo wiring', () => {
    const files = [
      join(repoRoot, 'package.json'),
      ...listFiles(join(repoRoot, 'scripts')),
      ...listFiles(join(repoRoot, 'e2e/tests')),
      ...listFiles(join(repoRoot, 'packages/test-fuzz')),
    ].filter(isTextInventoryFile)

    expect(findForbiddenMarkers(files)).toEqual([])
  })

  it('keeps every critical correctness feature group under direct fuzz coverage', () => {
    const files = [...listFuzzInventoryFiles(join(repoRoot, 'packages')), ...listFuzzInventoryFiles(join(repoRoot, 'apps'))]

    expect(findMissingFeatureGroups(files)).toEqual([])
  })
})

// Helpers

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isTextInventoryFile(file: string): boolean {
  return /\.(?:ts|tsx|js|mjs|json|dict)$/u.test(file) && !file.endsWith('scripts/__tests__/fuzz-inventory.test.ts')
}

function findForbiddenMarkers(files: readonly string[]): string[] {
  const found = new Set<string>()
  for (const file of files) {
    const text = readFileSync(file, 'utf8')
    for (const marker of forbiddenMarkers) {
      if (text.includes(marker)) {
        found.add(marker)
      }
    }
  }
  return [...found].toSorted((left, right) => left.localeCompare(right))
}

function findMissingFeatureGroups(files: readonly string[]): string[] {
  const remaining = new Map(requiredFeatureGroups.map((group) => [group.label, [...group.patterns]]))
  for (const file of files) {
    if (remaining.size === 0) {
      break
    }
    const text = `${file}\n${readFileSync(file, 'utf8')}`
    for (const [label, patterns] of remaining) {
      if (patterns.some((pattern) => text.includes(pattern))) {
        remaining.delete(label)
      }
    }
  }
  return [...remaining.keys()]
}

function listFiles(root: string): string[] {
  const files: string[] = []
  for (const entry of readdirSync(root)) {
    if (entry === 'node_modules' || entry === 'dist' || entry === 'coverage' || entry.startsWith('.')) {
      continue
    }
    const path = join(root, entry)
    const stat = statSync(path)
    if (stat.isDirectory()) {
      files.push(...listFiles(path))
    } else if (stat.isFile()) {
      files.push(path)
    }
  }
  return files
}

function listFuzzInventoryFiles(root: string): string[] {
  return listFiles(root).filter((file) => {
    if (!isTextInventoryFile(file)) {
      return false
    }
    if (file.includes('/packages/test-fuzz/byte-targets/') || file.includes('/packages/test-fuzz/dictionaries/')) {
      return true
    }
    return /\.fuzz\.test\.tsx?$/u.test(file)
  })
}
