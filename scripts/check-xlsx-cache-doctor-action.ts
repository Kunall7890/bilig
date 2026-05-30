import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const rootActionPath = join(repoRoot, 'action.yml')
const nestedActionPath = join(repoRoot, 'actions', 'xlsx-cache-doctor', 'action.yml')
const resolveScriptPath = join(repoRoot, 'actions', 'xlsx-cache-doctor', 'resolve-workbooks.mjs')
const inspectScriptPath = join(repoRoot, 'actions', 'xlsx-cache-doctor', 'inspect-workbooks.mjs')

const [rootAction, nestedAction] = await Promise.all([readFile(rootActionPath, 'utf8'), readFile(nestedActionPath, 'utf8')])

if (rootAction !== nestedAction) {
  throw new Error(
    [
      'Root action.yml must stay byte-for-byte aligned with actions/xlsx-cache-doctor/action.yml.',
      'Edit both files together so Marketplace and subdirectory action users get the same inputs, outputs, and behavior.',
    ].join('\n'),
  )
}

if (!existsSync(resolveScriptPath) || !existsSync(inspectScriptPath)) {
  throw new Error('XLSX Cache Doctor action helper scripts must exist next to the nested action.yml.')
}

if (!rootAction.includes('workbooks:') || !rootAction.includes('changed-files-only:')) {
  throw new Error('XLSX Cache Doctor action.yml must expose repo-scale workbooks and changed-files-only inputs.')
}

const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-cache-doctor-action-'))
try {
  writeFileSync(join(tempDir, 'a.xlsx'), '')
  writeFileSync(join(tempDir, 'notes.txt'), '')
  const outputPath = join(tempDir, 'github-output.txt')
  const result = spawnSync(process.execPath, [resolveScriptPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      GITHUB_WORKSPACE: tempDir,
      GITHUB_OUTPUT: outputPath,
      BILIG_WORKBOOKS: '**/*.xlsx',
      BILIG_CHANGED_FILES_ONLY: 'false',
    },
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`XLSX Cache Doctor resolver smoke failed:\n${result.stdout}\n${result.stderr}`)
  }
  const output = await readFile(outputPath, 'utf8')
  if (!output.includes('workbook-count=1') || !output.includes('workbooks-json=["a.xlsx"]')) {
    throw new Error(`Unexpected resolver output:\n${output}`)
  }
} finally {
  rmSync(tempDir, { recursive: true, force: true })
}

const changedFilesTempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-cache-doctor-changed-files-'))
try {
  runGit(changedFilesTempDir, ['init'])
  runGit(changedFilesTempDir, ['config', 'user.email', 'cache-doctor@example.com'])
  runGit(changedFilesTempDir, ['config', 'user.name', 'XLSX Cache Doctor'])
  writeFileSync(join(changedFilesTempDir, 'unchanged.xlsx'), '')
  runGit(changedFilesTempDir, ['add', 'unchanged.xlsx'])
  runGit(changedFilesTempDir, ['commit', '-m', 'base'])
  writeFileSync(join(changedFilesTempDir, 'changed.xlsx'), '')
  runGit(changedFilesTempDir, ['add', 'changed.xlsx'])
  runGit(changedFilesTempDir, ['commit', '-m', 'change workbook'])

  const outputPath = join(changedFilesTempDir, 'github-output.txt')
  const result = spawnSync(process.execPath, [resolveScriptPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      GITHUB_WORKSPACE: changedFilesTempDir,
      GITHUB_OUTPUT: outputPath,
      BILIG_WORKBOOKS: '**/*.xlsx',
      BILIG_CHANGED_FILES_ONLY: 'true',
    },
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`XLSX Cache Doctor changed-file resolver smoke failed:\n${result.stdout}\n${result.stderr}`)
  }
  const output = await readFile(outputPath, 'utf8')
  if (!output.includes('workbook-count=1') || !output.includes('workbooks-json=["changed.xlsx"]')) {
    throw new Error(`Unexpected changed-file resolver output:\n${output}`)
  }
} finally {
  rmSync(changedFilesTempDir, { recursive: true, force: true })
}

function runGit(cwd: string, args: readonly string[]): void {
  const result = spawnSync('git', args, {
    cwd,
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`git ${args.join(' ')} failed:\n${result.stdout}\n${result.stderr}`)
  }
}
