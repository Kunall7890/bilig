#!/usr/bin/env bun

import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { loadRuntimePackages, parseStableSemver } from './runtime-package-set.ts'

export interface SyncRuntimePackageVersionsOptions {
  rootDir: string
  version: string
}

export interface SyncRuntimePackageVersionsResult {
  version: string
  updatedFiles: string[]
  updatedPackages: string[]
}

export function syncRuntimePackageVersions(options: SyncRuntimePackageVersionsOptions): SyncRuntimePackageVersionsResult {
  const version = options.version.trim()
  parseStableSemver(version)

  const updatedFiles: string[] = []
  const runtimePackages = loadRuntimePackages(options.rootDir)

  for (const runtimePackage of runtimePackages) {
    const packageJsonPath = join(options.rootDir, runtimePackage.dir, 'package.json')
    const manifest = readJsonRecord(packageJsonPath)
    manifest['version'] = version
    if (writeJsonIfChanged(packageJsonPath, manifest)) {
      updatedFiles.push(packageJsonPath)
    }
  }

  syncReleasePleaseManifestVersion(options.rootDir, version, updatedFiles)
  syncMcpServerVersions(options.rootDir, version, runtimePackages, updatedFiles)
  syncDockerfileWorkpaperVersion(options.rootDir, version, updatedFiles)
  syncGeminiExtensionVersion(options.rootDir, version, updatedFiles)
  syncXlsxCacheDoctorActionVersion(options.rootDir, version, updatedFiles)
  syncAgentEvaluatorDocVersions(options.rootDir, version, updatedFiles)
  syncHuggingFaceWorkpaperSpaceVersion(options.rootDir, version, updatedFiles)
  syncMcpDirectoryDocVersion(options.rootDir, version, updatedFiles)

  return {
    version,
    updatedFiles,
    updatedPackages: runtimePackages.map((runtimePackage) => runtimePackage.name),
  }
}

function syncHuggingFaceWorkpaperSpaceVersion(rootDir: string, version: string, updatedFiles: string[]): void {
  const packageJsonPath = join(rootDir, 'examples/huggingface-workpaper-space/package.json')
  const packageJson = readJsonRecord(packageJsonPath)
  const dependencies = packageJson['dependencies']
  if (!isRecord(dependencies)) {
    throw new Error(`${packageJsonPath} must define object dependencies`)
  }
  dependencies['@bilig/workpaper'] = version
  if (writeJsonIfChanged(packageJsonPath, packageJson)) {
    updatedFiles.push(packageJsonPath)
  }

  const readmePath = join(rootDir, 'examples/huggingface-workpaper-space/README.md')
  const readme = readFileSync(readmePath, 'utf8')
  const nextReadme = readme
    .replace(/@bilig\/workpaper@\d+\.\d+\.\d+/gu, `@bilig/workpaper@${version}`)
    .replace(/"packageVersion": "\d+\.\d+\.\d+"/gu, `"packageVersion": "${version}"`)
  writeTextIfChanged(readmePath, readme, nextReadme, updatedFiles)

  const proofPath = join(rootDir, 'examples/huggingface-workpaper-space/workpaper_proof.mjs')
  const proof = readFileSync(proofPath, 'utf8')
  const nextProof = replaceRequired(
    proof,
    /const workpaperPackageVersion = '\d+\.\d+\.\d+'/u,
    `const workpaperPackageVersion = '${version}'`,
    `${proofPath} must define workpaperPackageVersion`,
  )
  writeTextIfChanged(proofPath, proof, nextProof, updatedFiles)

  const checkPath = join(rootDir, 'examples/huggingface-workpaper-space/scripts/check-space.py')
  const check = readFileSync(checkPath, 'utf8')
  const nextCheck = replaceRequired(
    check,
    /payload\.get\("packageVersion"\) != "\d+\.\d+\.\d+"/u,
    `payload.get("packageVersion") != "${version}"`,
    `${checkPath} must validate packageVersion`,
  )
  writeTextIfChanged(checkPath, check, nextCheck, updatedFiles)
}

function syncMcpDirectoryDocVersion(rootDir: string, version: string, updatedFiles: string[]): void {
  const docPath = join(rootDir, 'docs/mcp-spreadsheet-server-directory.md')
  const currentContent = readFileSync(docPath, 'utf8')
  const nextContent = replaceRequired(
    currentContent,
    /current\s+repo\s+package\s+version\s+is\s+`\d+\.\d+\.\d+`/u,
    `current repo package version is \`${version}\``,
    `${docPath} must include the current repo package version`,
  )
  writeTextIfChanged(docPath, currentContent, nextContent, updatedFiles)
}

function syncAgentEvaluatorDocVersions(rootDir: string, version: string, updatedFiles: string[]): void {
  const docPaths = [join(rootDir, 'docs/agent-adoption-kit.md'), join(rootDir, 'docs/eval-agent-mcp.md')]
  for (const docPath of docPaths) {
    const currentContent = readFileSync(docPath, 'utf8')
    const nextContent = replaceRequired(
      replaceRequired(
        currentContent,
        /("@bilig\/workpaper":\s*")\d+\.\d+\.\d+(")/u,
        `$1${version}$2`,
        `${docPath} must include a @bilig/workpaper evaluator package version`,
      ),
      /("xlsx-formula-recalc":\s*")\d+\.\d+\.\d+(")/u,
      `$1${version}$2`,
      `${docPath} must include an xlsx-formula-recalc evaluator package version`,
    )
    writeTextIfChanged(docPath, currentContent, nextContent, updatedFiles)
  }
}

function syncXlsxCacheDoctorActionVersion(rootDir: string, version: string, updatedFiles: string[]): void {
  const actionPaths = [join(rootDir, 'action.yml'), join(rootDir, 'actions/xlsx-cache-doctor/action.yml')]
  for (const actionPath of actionPaths) {
    const currentContent = readFileSync(actionPath, 'utf8')
    const nextContent = replaceRequired(
      currentContent,
      /(package-version:\n\s+description: npm version or dist-tag for @bilig\/xlsx-formula-recalc\. Pin this for production workflows\.\n\s+required: false\n\s+default: )'[^']+'/u,
      `$1'${version}'`,
      `${actionPath} must include a package-version default`,
    )
    writeTextIfChanged(actionPath, currentContent, nextContent, updatedFiles)
  }

  const markdownPaths = [join(rootDir, 'docs/xlsx-cache-doctor-github-action.md'), join(rootDir, 'README.md')]
  for (const markdownPath of markdownPaths) {
    const markdownContent = readFileSync(markdownPath, 'utf8')
    const nextMarkdownContent = syncXlsxCacheDoctorMarkdownVersion(markdownPath, markdownContent, version)
    writeTextIfChanged(markdownPath, markdownContent, nextMarkdownContent, updatedFiles)
  }
}

function syncXlsxCacheDoctorMarkdownVersion(path: string, content: string, version: string): string {
  const nextContent = content.replaceAll(/package-version: '\d+\.\d+\.\d+'/g, `package-version: '${version}'`)

  if (path.endsWith('docs/xlsx-cache-doctor-github-action.md')) {
    return replaceRequired(
      nextContent,
      /(\| `package-version`\s+\|\s+)\d+\.\d+\.\d+(\s+\| npm version or dist-tag for `@bilig\/xlsx-formula-recalc`\. Pin this in production\. \|)/u,
      `$1${version}$2`,
      `${path} must include the package-version input table row`,
    )
  }

  return nextContent
}

function syncDockerfileWorkpaperVersion(rootDir: string, version: string, updatedFiles: string[]): void {
  const dockerfilePath = join(rootDir, 'Dockerfile')
  const currentContent = readFileSync(dockerfilePath, 'utf8')
  const versionArgPattern = /^ARG BILIG_WORKPAPER_VERSION=.*$/mu
  if (!versionArgPattern.test(currentContent)) {
    throw new Error('Dockerfile must include ARG BILIG_WORKPAPER_VERSION')
  }

  const nextContent = currentContent.replace(versionArgPattern, `ARG BILIG_WORKPAPER_VERSION=${version}`)
  if (currentContent === nextContent) {
    return
  }
  writeFileSync(dockerfilePath, nextContent)
  updatedFiles.push(dockerfilePath)
}

function syncGeminiExtensionVersion(rootDir: string, version: string, updatedFiles: string[]): void {
  const geminiExtensionPath = join(rootDir, 'gemini-extension.json')
  const geminiExtension = readJsonRecord(geminiExtensionPath)
  geminiExtension['version'] = version

  if (writeJsonIfChanged(geminiExtensionPath, geminiExtension)) {
    updatedFiles.push(geminiExtensionPath)
  }
}

function syncReleasePleaseManifestVersion(rootDir: string, version: string, updatedFiles: string[]): void {
  const manifestPath = join(rootDir, '.release-please-manifest.json')
  const manifest = readJsonRecord(manifestPath)
  manifest['packages/headless'] = version

  if (writeJsonIfChanged(manifestPath, manifest)) {
    updatedFiles.push(manifestPath)
  }
}

function syncMcpServerVersions(
  rootDir: string,
  version: string,
  runtimePackages: readonly { readonly dir: string; readonly name: string }[],
  updatedFiles: string[],
): void {
  for (const runtimePackage of runtimePackages) {
    const packageJsonPath = join(rootDir, runtimePackage.dir, 'package.json')
    const manifest = readJsonRecord(packageJsonPath)
    if (typeof manifest['mcpName'] !== 'string') {
      continue
    }
    const serverJsonPath = join(rootDir, runtimePackage.dir, 'server.json')
    const serverJson = readJsonRecord(serverJsonPath)
    serverJson['version'] = version

    const npmPackage = findNpmPackageEntry(serverJson, runtimePackage.name)
    if (!npmPackage) {
      throw new Error(`${runtimePackage.dir}/server.json must include an npm package entry for ${runtimePackage.name}`)
    }
    npmPackage['version'] = version

    if (writeJsonIfChanged(serverJsonPath, serverJson)) {
      updatedFiles.push(serverJsonPath)
    }
  }
}

function findNpmPackageEntry(serverJson: Record<string, unknown>, packageName: string): Record<string, unknown> | undefined {
  const packages = serverJson['packages']
  if (!Array.isArray(packages)) {
    return undefined
  }
  return packages.find(
    (entry): entry is Record<string, unknown> => isRecord(entry) && entry['registryType'] === 'npm' && entry['identifier'] === packageName,
  )
}

function readJsonRecord(path: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isRecord(parsed)) {
    throw new Error(`Expected JSON object in ${path}`)
  }
  return parsed
}

function writeJsonIfChanged(path: string, value: Record<string, unknown>): boolean {
  const nextContent = `${JSON.stringify(value, null, 2)}\n`
  if (readFileSync(path, 'utf8') === nextContent) {
    return false
  }
  writeFileSync(path, nextContent)
  return true
}

function writeTextIfChanged(path: string, currentContent: string, nextContent: string, updatedFiles: string[]): void {
  if (currentContent === nextContent) {
    return
  }
  writeFileSync(path, nextContent)
  updatedFiles.push(path)
}

function replaceRequired(content: string, pattern: RegExp, replacement: string, errorMessage: string): string {
  if (!pattern.test(content)) {
    throw new Error(errorMessage)
  }
  return content.replace(pattern, replacement)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readRequiredStringArg(args: Map<string, string | true>, name: string): string {
  const value = args.get(name)
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`--${name} is required`)
  }
  return value.trim()
}

function parseArgs(argv: readonly string[]): Map<string, string | true> {
  const args = new Map<string, string | true>()
  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index]
    if (!value || value === '--' || !value.startsWith('--')) {
      continue
    }
    const key = value.slice(2)
    const nextValue = argv[index + 1]
    if (!nextValue || nextValue.startsWith('--')) {
      args.set(key, true)
      continue
    }
    args.set(key, nextValue)
    index += 1
  }
  return args
}

function isDirectInvocation(): boolean {
  const scriptPath = process.argv[1]
  return Boolean(scriptPath) && import.meta.url === pathToFileURL(resolve(scriptPath)).href
}

if (isDirectInvocation()) {
  const args = parseArgs(process.argv.slice(2))
  const version = readRequiredStringArg(args, 'version')
  const rootDir = resolve(new URL('..', import.meta.url).pathname)
  const result = syncRuntimePackageVersions({ rootDir, version })
  console.log(JSON.stringify(result, null, 2))
}
