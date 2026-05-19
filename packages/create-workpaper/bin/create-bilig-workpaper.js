#!/usr/bin/env node
import { cp, mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)))
const templateRoot = join(packageRoot, 'template')
const agentOverlayRoot = join(packageRoot, 'agent-overlay')

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

const targetDirectory = resolve(process.cwd(), options.targetDirectory ?? 'bilig-workpaper-starter')
const projectName = normalizePackageName(options.targetDirectory ?? 'bilig-workpaper-starter')

await ensureWritableTarget(targetDirectory, options.force)
await copyTemplate(templateRoot, targetDirectory, projectName)

if (options.template === 'agent') {
  await copyTemplate(agentOverlayRoot, targetDirectory, projectName)
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

function parseCliArgs(inputArgs) {
  const parsed = {
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
  npm exec @bilig/create-workpaper@latest <directory>

Options:
  --agent                    Add AGENTS.md, MCP client configs, and agent smoke scripts.
  --template service|agent   Choose the starter shape. Defaults to service.
  --force                    Allow writing into an existing directory.
  -h, --help                 Print this help text.
`)
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

async function copyTemplate(sourceRoot, outputDirectory, packageName) {
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
      await writeFile(targetPath, text.replaceAll('__PROJECT_NAME__', packageName))
      return false
    },
  })
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
