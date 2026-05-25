#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'

type PackageManifest = Readonly<Record<string, unknown>>

const repoRoot = resolve(import.meta.dir, '..')
const packageDir = join(repoRoot, 'packages', 'create-workpaper')
const packageName = '@bilig/create-workpaper'
const packDir = join(repoRoot, 'build', 'create-workpaper-package')
const generatedDir = join(packDir, 'generated')
const args = new Set(process.argv.slice(2))
const requirePublished = args.has('--require-published')

function fail(message: string): never {
  throw new Error(message)
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    fail(message)
  }
}

function readJson(path: string): PackageManifest {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  assert(isRecord(parsed), `${path} must contain a JSON object`)
  return parsed
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function stringArray(value: unknown): readonly string[] {
  assert(Array.isArray(value) && value.every((entry) => typeof entry === 'string'), 'expected string array')
  return value
}

function run(command: string, commandArgs: readonly string[], options: { readonly allowFailure?: boolean } = {}) {
  const result = spawnSync(command, commandArgs, {
    cwd: repoRoot,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  if (result.status !== 0 && !options.allowFailure) {
    const stderr = result.stderr.trim()
    const stdout = result.stdout.trim()
    fail(`${command} ${commandArgs.join(' ')} failed\n${stderr || stdout}`)
  }
  return result
}

function assertManifest(manifest: PackageManifest): string {
  assert(manifest.name === packageName, `packages/create-workpaper/package.json.name must be ${packageName}`)
  assert(typeof manifest.version === 'string' && /^\d+\.\d+\.\d+$/u.test(manifest.version), 'package version must be stable semver')
  assert(
    manifest.description === 'Create a runnable Bilig WorkPaper starter for Node services and agent tools.',
    'package description must keep the starter promise specific',
  )
  assert(manifest.homepage === 'https://proompteng.github.io/bilig/', 'package homepage must point to public docs')
  assert(manifest.license === 'MIT', 'package license must be MIT')
  assert(manifest.type === 'module', 'package type must be module')

  const keywords = stringArray(manifest.keywords)
  for (const keyword of ['formula-engine', 'spreadsheet-formulas', 'typescript', 'workbook-api', 'workpaper', 'xlsx-recalculation']) {
    assert(keywords.includes(keyword), `package keywords must include ${keyword}`)
  }

  assert(isRecord(manifest.bugs), 'package bugs must be an object')
  assert(manifest.bugs['url'] === 'https://github.com/proompteng/bilig/issues', 'package bugs.url must point to GitHub issues')

  assert(isRecord(manifest.repository), 'package repository must be an object')
  assert(manifest.repository['type'] === 'git', 'package repository.type must be git')
  assert(manifest.repository['url'] === 'git+https://github.com/proompteng/bilig.git', 'package repository.url must point to GitHub')
  assert(manifest.repository['directory'] === 'packages/create-workpaper', 'package repository.directory must be packages/create-workpaper')

  assert(isRecord(manifest.bin), 'package bin must be an object')
  assert(manifest.bin['create-workpaper'] === 'bin/create-bilig-workpaper.js', 'package bin must expose create-workpaper')

  const files = stringArray(manifest.files)
  for (const included of ['agent-overlay', 'bin', 'template', 'README.md']) {
    assert(files.includes(included), `package files must include ${included}`)
  }

  assert(isRecord(manifest.publishConfig), 'package publishConfig must be an object')
  assert(manifest.publishConfig['access'] === 'public', 'package publishConfig.access must be public')

  assert(isRecord(manifest.engines), 'package engines must be an object')
  assert(manifest.engines['node'] === '>=22.0.0', 'package engines.node must stay Node 22 compatible')

  assert(isRecord(manifest.scripts), 'package scripts must be an object')
  assert(manifest.scripts['smoke'] === 'node ./bin/create-bilig-workpaper.js --help', 'package smoke script must exercise the CLI')

  return manifest.version
}

function assertDocs(): void {
  const readme = readFileSync(join(packageDir, 'README.md'), 'utf8')
  const docs = readFileSync(join(repoRoot, 'docs', 'create-bilig-workpaper.md'), 'utf8')
  const rootReadme = readFileSync(join(repoRoot, 'README.md'), 'utf8')
  const templateSource = readFileSync(join(packageDir, 'template', 'src', 'index.ts'), 'utf8')
  for (const [label, source] of [
    ['packages/create-workpaper/README.md', readme],
    ['docs/create-bilig-workpaper.md', docs],
    ['README.md', rootReadme],
  ] as const) {
    assert(source.includes('npm create @bilig/workpaper@latest'), `${label} must document the scoped npm create path`)
    assert(source.includes('@bilig/create-workpaper'), `${label} must include the published package name`)
    assert(source.includes('--agent'), `${label} must document the agent-ready starter path`)
  }
  assert(readme.includes('verified: true') || readme.includes('"verified": true'), 'starter README must show the verification output')
  assert(docs.includes('verified: true') || docs.includes('"verified": true'), 'starter docs must show the verification output')
  assert(
    templateSource.includes('assertSmokeOutput(output)') &&
      templateSource.indexOf('assertSmokeOutput(output)') < templateSource.indexOf('console.log(JSON.stringify(output, null, 2))'),
    'starter smoke output must be verified before it is printed',
  )
  assert(
    templateSource.includes('https://github.com/proompteng/bilig/discussions/new?category=general'),
    'starter smoke output must include the adoption-feedback link after verification',
  )
  assert(templateSource.includes('https://github.com/proompteng/bilig/stargazers'), 'starter smoke output must include the star link')
  assert(
    templateSource.includes('https://github.com/proompteng/bilig/subscription'),
    'starter smoke output must include the release-watch link',
  )
  assert(
    docs.includes('https://github.com/proompteng/bilig/discussions/new?category=general'),
    'starter docs must include the adoption-feedback link',
  )
  assert(docs.includes('https://github.com/proompteng/bilig/stargazers'), 'starter docs must include the star link')
  assert(docs.includes('https://github.com/proompteng/bilig/subscription'), 'starter docs must include the release-watch link')
  assert(readme.includes('agent:verify'), 'starter README must document the agent verification script')
  assert(docs.includes('agent:verify'), 'starter docs must document the agent verification script')
}

function assertPackedTarball(): void {
  rmSync(packDir, { force: true, recursive: true })
  mkdirSync(packDir, { recursive: true })
  run('pnpm', ['--filter', packageName, 'pack', '--pack-destination', packDir])

  const tarballs = readdirSync(packDir).filter((entry) => entry.endsWith('.tgz'))
  assert(tarballs.length === 1, `expected one packed tarball in ${packDir}, found ${String(tarballs.length)}`)
  const tarball = join(packDir, tarballs[0] ?? fail('missing tarball'))
  const tarOutput = run('tar', ['-tf', tarball]).stdout
  const entries = new Set(tarOutput.trim().split('\n').filter(Boolean))
  for (const entry of [
    'package/package.json',
    'package/README.md',
    'package/agent-overlay/.cursor/mcp.json',
    'package/agent-overlay/.vscode/mcp.json',
    'package/agent-overlay/AGENTS.md',
    'package/agent-overlay/CLAUDE.md',
    'package/agent-overlay/README.md',
    'package/agent-overlay/mcp/bilig-workpaper.mcp.json',
    'package/agent-overlay/package.json',
    'package/bin/create-bilig-workpaper.js',
    'package/template/package.json',
    'package/template/README.md',
    'package/template/src/index.ts',
    'package/template/tsconfig.json',
  ]) {
    assert(entries.has(entry), `${basename(tarball)} is missing ${entry}`)
  }

  for (const entry of entries) {
    assert(!entry.includes('node_modules/'), `${basename(tarball)} must not include node_modules`)
    assert(!entry.includes('.cache/'), `${basename(tarball)} must not include .cache output`)
  }
}

function assertGeneratedStarters(): void {
  rmSync(generatedDir, { force: true, recursive: true })
  mkdirSync(generatedDir, { recursive: true })

  const cliPath = join(packageDir, 'bin', 'create-bilig-workpaper.js')
  const serviceDir = join(generatedDir, 'service-demo')
  const agentDir = join(generatedDir, 'agent-demo')

  run('node', [cliPath, serviceDir])
  run('node', [cliPath, agentDir, '--agent'])

  const serviceManifest = readJson(join(serviceDir, 'package.json'))
  assert(isRecord(serviceManifest.scripts), 'generated service package scripts must be an object')
  assert(serviceManifest.scripts['smoke'] === 'tsx src/index.ts', 'generated service starter must keep the smoke script')
  assert(serviceManifest.scripts['agent:verify'] === undefined, 'generated service starter must not include agent-only scripts')

  const agentManifest = readJson(join(agentDir, 'package.json'))
  assert(isRecord(agentManifest.scripts), 'generated agent package scripts must be an object')
  assert(
    agentManifest.scripts['agent:verify'] === 'npm run smoke && npm run mcp:challenge',
    'generated agent starter must verify API and MCP paths',
  )
  assert(
    agentManifest.scripts['mcp:server'] === 'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable',
    'generated agent starter must include the file-backed MCP server script',
  )

  for (const expected of ['AGENTS.md', 'CLAUDE.md', '.cursor/mcp.json', '.vscode/mcp.json', 'mcp/bilig-workpaper.mcp.json']) {
    assert(existsSync(join(agentDir, expected)), `generated agent starter is missing ${expected}`)
  }
}

function assertPublishedVersion(version: string): void {
  const result = run('npm', ['view', `${packageName}@${version}`, 'version', '--json'], { allowFailure: true })
  if (result.status === 0 && result.stdout.includes(version)) {
    console.log(`${packageName}@${version} is published on npm`)
    return
  }

  const message = `${packageName}@${version} is not published on npm yet`
  if (requirePublished) {
    fail(message)
  }
  console.warn(`${message}; run with --require-published in release verification after the npm package exists.`)
}

const manifest = readJson(join(packageDir, 'package.json'))
const version = assertManifest(manifest)

assert(existsSync(join(packageDir, 'bin', 'create-bilig-workpaper.js')), 'CLI entrypoint must exist')
assert(existsSync(join(packageDir, 'template', 'src', 'index.ts')), 'starter template source must exist')
assert(existsSync(join(packageDir, 'agent-overlay', 'AGENTS.md')), 'agent starter overlay must exist')
assertDocs()
assertGeneratedStarters()
assertPackedTarball()
assertPublishedVersion(version)

console.log(`${packageName}@${version} package checks passed`)
