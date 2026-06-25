#!/usr/bin/env node
import { cp, mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)))
const templateRoot = join(packageRoot, 'template')
const agentOverlayRoot = join(packageRoot, 'agent-overlay')
const starterWorkpaperPath = './pricing.workpaper.json'
const existingRepoWorkpaperPath = './.bilig/pricing.workpaper.json'

const args = process.argv.slice(2)
let options

try {
  options = parseCliArgs(args)
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error))
  console.error('')
  printHelp()
  process.exit(1)
}

if (options.help) {
  printHelp()
  process.exit(0)
}

const defaultTargetDirectory = options.addAgent ? '.' : 'bilig-workpaper-starter'
const targetDirectory = resolve(process.cwd(), options.targetDirectory ?? defaultTargetDirectory)
const generatedWorkpaperPackageVersion = await readPackageVersion()
const projectName = options.addAgent
  ? await resolveExistingProjectName(targetDirectory, options.targetDirectory ?? defaultTargetDirectory)
  : normalizePackageName(options.targetDirectory ?? defaultTargetDirectory)

if (options.addAgent) {
  await ensureDirectory(targetDirectory)
  const overlayResult = await copyAgentOverlay(
    agentOverlayRoot,
    targetDirectory,
    projectName,
    existingRepoWorkpaperPath,
    generatedWorkpaperPackageVersion,
    options.force,
  )

  const targetLabel = relative(process.cwd(), targetDirectory) || '.'
  console.log(`Added Bilig MCP and host integration files to ${targetLabel}`)
  printChangeList('Wrote', overlayResult.written)
  printChangeList('Skipped existing', overlayResult.skipped)
  if (overlayResult.installSummary !== undefined) {
    console.log(`Review skipped-file handoff: ${overlayResult.installSummary}`)
  }
  console.log('')
  console.log('Next:')
  if (targetLabel !== '.') {
    console.log(`  cd ${targetLabel}`)
  }
  console.log('  npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json')
  console.log(
    `  npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ${existingRepoWorkpaperPath} --init-demo-workpaper --writable`,
  )
  console.log('')
  console.log('Expected proof output includes:')
  console.log('  "verified": true')
} else {
  await ensureWritableTarget(targetDirectory, options.force)
  await copyTemplate(templateRoot, targetDirectory, projectName, starterWorkpaperPath, generatedWorkpaperPackageVersion)

  if (options.template === 'agent') {
    await copyTemplate(agentOverlayRoot, targetDirectory, projectName, starterWorkpaperPath, generatedWorkpaperPackageVersion)
  }

  console.log(`Created ${relative(process.cwd(), targetDirectory) || '.'}`)
  console.log('')
  console.log('Next:')
  console.log(`  cd ${relative(process.cwd(), targetDirectory) || '.'}`)
  console.log('  npm install')
  console.log(options.template === 'agent' ? '  npm run agent:verify' : '  npm run smoke')
  if (options.template === 'agent') {
    console.log('  npm run mcp:server')
  }
  console.log('')
  console.log('Expected smoke output includes:')
  console.log('  "verified": true')
}

function parseCliArgs(inputArgs) {
  const parsed = {
    addAgent: false,
    force: false,
    help: false,
    targetDirectory: undefined,
    template: 'service',
  }

  for (let index = 0; index < inputArgs.length; index += 1) {
    const arg = inputArgs[index]
    if (arg === '--help' || arg === '-h') {
      parsed.help = true
      continue
    }
    if (arg === '--force') {
      parsed.force = true
      continue
    }
    if (arg === '--add-agent' || arg === '--overlay-only') {
      parsed.addAgent = true
      parsed.template = 'agent'
      continue
    }
    if (arg === '--agent') {
      parsed.template = 'agent'
      continue
    }
    if (arg === '--service') {
      parsed.template = 'service'
      continue
    }
    if (arg === '--template') {
      const nextValue = inputArgs[index + 1]
      if (nextValue === undefined || nextValue.startsWith('-')) {
        throw new Error('--template requires service or agent')
      }
      parsed.template = normalizeTemplate(nextValue)
      index += 1
      continue
    }
    if (arg.startsWith('--template=')) {
      parsed.template = normalizeTemplate(arg.slice('--template='.length))
      continue
    }
    if (arg.startsWith('-')) {
      throw new Error(`Unknown option: ${arg}`)
    }
    if (parsed.targetDirectory !== undefined) {
      throw new Error(`Unexpected extra directory argument: ${arg}`)
    }
    parsed.targetDirectory = arg
  }

  return parsed
}

function normalizeTemplate(value) {
  if (value === 'service' || value === 'agent') {
    return value
  }
  throw new Error(`Unknown template: ${value}. Expected service or agent.`)
}

function printHelp() {
  console.log(`@bilig/create-workpaper

Usage:
  npm create @bilig/workpaper@latest <directory>
  npm create @bilig/workpaper@latest <directory> -- --agent
  npm create @bilig/workpaper@latest . -- --add-agent
  npm exec @bilig/create-workpaper@latest <directory>

Options:
  --agent                    Add MCP client configs, host files, and verification scripts.
  --add-agent                Add only MCP and host files to an existing repo.
  --template service|agent   Choose the starter shape. Defaults to service.
  --force                    Allow overwriting existing generated/overlay files.
  -h, --help                 Print this help text.
`)
}

async function ensureDirectory(directory) {
  try {
    const existing = await stat(directory)
    if (!existing.isDirectory()) {
      throw new Error(`${directory} exists and is not a directory`)
    }
  } catch (error) {
    if (error?.code === 'ENOENT') {
      await mkdir(directory, { recursive: true })
      return
    }
    throw error
  }
}

async function ensureWritableTarget(directory, allowExisting) {
  try {
    const existing = await stat(directory)
    if (!existing.isDirectory()) {
      throw new Error(`${directory} exists and is not a directory`)
    }
    const entries = await readdir(directory)
    if (entries.length > 0 && !allowExisting) {
      throw new Error(`${directory} is not empty. Re-run with --force to write into it.`)
    }
  } catch (error) {
    if (error?.code === 'ENOENT') {
      await mkdir(directory, { recursive: true })
      return
    }
    throw error
  }
}

async function copyTemplate(sourceRoot, outputDirectory, packageName, workpaperPath, workpaperPackageVersion) {
  await cp(sourceRoot, outputDirectory, {
    recursive: true,
    filter: async (source) => {
      const sourceStat = await stat(source)
      if (sourceStat.isDirectory()) {
        return true
      }
      const relativePath = relative(sourceRoot, source)
      const targetPath = join(outputDirectory, relativePath)
      const text = await readFile(source, 'utf8')
      await mkdir(dirname(targetPath), { recursive: true })
      await writeFile(targetPath, renderTemplate(text, packageName, workpaperPath, workpaperPackageVersion))
      return false
    },
  })
}

async function copyAgentOverlay(sourceRoot, outputDirectory, packageName, workpaperPath, workpaperPackageVersion, force) {
  const files = await listFiles(sourceRoot)
  const results = await Promise.all(
    files.map(async (sourcePath) => {
      const relativePath = relative(sourceRoot, sourcePath)
      if (relativePath === 'package.json') {
        return { path: relativePath, status: 'ignored' }
      }

      const targetRelativePath = relativePath === 'README.md' ? 'BILIG_WORKPAPER.md' : relativePath
      const targetPath = join(outputDirectory, targetRelativePath)

      if (!force && (await pathExists(targetPath))) {
        return { path: targetRelativePath, status: 'skipped' }
      }

      const text = await readFile(sourcePath, 'utf8')
      await mkdir(dirname(targetPath), { recursive: true })
      await writeFile(targetPath, renderTemplate(text, packageName, workpaperPath, workpaperPackageVersion))
      return { path: targetRelativePath, status: 'written' }
    }),
  )
  const skipped = results.filter((result) => result.status === 'skipped').map((result) => result.path)
  const written = results.filter((result) => result.status === 'written').map((result) => result.path)
  let installSummary
  if (!force && skipped.length > 0) {
    installSummary = 'BILIG_WORKPAPER_INSTALL.md'
    const summaryPath = join(outputDirectory, installSummary)
    await writeFile(summaryPath, renderInstallSummary(packageName, workpaperPath, skipped))
    if (!written.includes(installSummary)) {
      written.push(installSummary)
    }
  }

  return {
    installSummary,
    skipped,
    written,
  }
}

function renderTemplate(text, packageName, workpaperPath, workpaperPackageVersion) {
  return text
    .replaceAll('__PROJECT_NAME__', packageName)
    .replaceAll('__WORKPAPER_PATH__', workpaperPath)
    .replaceAll('__BILIG_WORKPAPER_VERSION__', workpaperPackageVersion)
}

function renderInstallSummary(packageName, workpaperPath, skipped) {
  const skippedList = skipped.map((path) => `- ${path}`).join('\n')
  return `# Bilig WorkPaper Install Notes

Project: \`${packageName}\`

The agent overlay installed the files it could without overwriting existing
project instructions.

## Skipped Existing Files

${skippedList}

## Add To Existing Agent Instructions

\`\`\`text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Read
BILIG_WORKPAPER.md first. For formula or workbook tasks, edit one exact input
or formula, read the recalculated dependent value, and report persistence
evidence. Do not claim success from a write call alone.
\`\`\`

## Verify

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ${workpaperPath} --init-demo-workpaper --writable
\`\`\`

Expected evaluator output includes \`"verified": true\`.
`
}

async function listFiles(directory) {
  const entries = await readdir(directory, { withFileTypes: true })
  const nestedEntries = await Promise.all(
    entries.map(async (entry) => {
      const entryPath = join(directory, entry.name)
      if (entry.isDirectory()) {
        return listFiles(entryPath)
      }
      return entry.isFile() ? [entryPath] : []
    }),
  )

  return nestedEntries.flat()
}

async function pathExists(path) {
  try {
    await stat(path)
    return true
  } catch (error) {
    if (error?.code === 'ENOENT') {
      return false
    }
    throw error
  }
}

async function resolveExistingProjectName(directory, fallbackName) {
  const manifestPath = join(directory, 'package.json')
  try {
    const parsed = JSON.parse(await readFile(manifestPath, 'utf8'))
    if (typeof parsed?.name === 'string' && parsed.name.trim() !== '') {
      return normalizePackageName(parsed.name)
    }
  } catch (error) {
    if (error?.code !== 'ENOENT') {
      throw new Error(`Could not read ${manifestPath}: ${error instanceof Error ? error.message : String(error)}`, {
        cause: error,
      })
    }
  }

  if (fallbackName === '.' || fallbackName === './') {
    return normalizePackageName(directory)
  }
  return normalizePackageName(fallbackName)
}

async function readPackageVersion() {
  const manifestPath = join(packageRoot, 'package.json')
  const parsed = JSON.parse(await readFile(manifestPath, 'utf8'))
  if (typeof parsed?.version !== 'string' || !/^\d+\.\d+\.\d+$/u.test(parsed.version)) {
    throw new Error(`Expected ${manifestPath} to contain a stable package version`)
  }
  return parsed.version
}

function printChangeList(label, values) {
  if (values.length === 0) {
    return
  }
  console.log(`${label}:`)
  for (const value of values) {
    console.log(`  ${value}`)
  }
}

function normalizePackageName(name) {
  const parts = name.split(/[\\/]/)
  let base = 'bilig-workpaper-starter'
  for (let index = parts.length - 1; index >= 0; index -= 1) {
    if (parts[index] !== '') {
      base = parts[index]
      break
    }
  }
  const normalized = base
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '')

  return normalized || 'bilig-workpaper-starter'
}
