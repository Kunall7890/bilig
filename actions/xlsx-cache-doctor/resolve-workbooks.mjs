#!/usr/bin/env node

import { appendFileSync, existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { relative, sep } from 'node:path'
import { spawnSync } from 'node:child_process'

const workspace = process.env.GITHUB_WORKSPACE || process.cwd()
process.chdir(workspace)

const changedFilesOnly = process.env.BILIG_CHANGED_FILES_ONLY === 'true'
const patterns = parsePatterns(process.env.BILIG_WORKBOOKS || process.env.BILIG_WORKBOOK || '')

if (patterns.length === 0) {
  throw new Error('Set either workbook or workbooks for XLSX Cache Doctor.')
}

const candidates = unique(patterns.flatMap((pattern) => resolvePattern(pattern)))
if (candidates.length === 0) {
  throw new Error(`No XLSX workbooks matched: ${patterns.join(', ')}`)
}

const changedFiles = changedFilesOnly ? collectChangedFiles() : null
const workbooks = changedFiles ? candidates.filter((file) => changedFiles.has(file)) : candidates
const output = JSON.stringify(workbooks)

console.log(`Resolved ${workbooks.length.toString()} XLSX workbook(s).`)
for (const workbook of workbooks) {
  console.log(`- ${workbook}`)
}

writeGithubOutput('workbook-count', String(workbooks.length))
writeGithubOutput('workbooks-json', output)

function parsePatterns(raw) {
  return raw
    .split(/[\n,]/u)
    .map((item) => item.trim())
    .filter(Boolean)
}

function resolvePattern(pattern) {
  const normalizedPattern = normalizePath(pattern)
  if (!hasGlob(normalizedPattern)) {
    return existsSync(normalizedPattern) && normalizedPattern.endsWith('.xlsx') ? [normalizedPattern] : []
  }
  const matcher = globToRegExp(normalizedPattern)
  return listFiles(workspace)
    .map((file) => normalizePath(relative(workspace, file)))
    .filter((file) => file.endsWith('.xlsx') && matcher.test(file))
}

function collectChangedFiles() {
  const changed = collectChangedFilesFromGit()
  if (changed.size === 0) {
    console.log('No changed files were found for changed-files-only mode.')
  }
  return changed
}

function collectChangedFilesFromGit() {
  const event = readEvent()
  const baseSha = event?.pull_request?.base?.sha
  const headSha = event?.pull_request?.head?.sha || process.env.GITHUB_SHA || 'HEAD'
  if (baseSha) {
    const changed = diffChangedXlsxFiles(baseSha, headSha)
    if (changed) {
      return changed
    }
    fetchRef(baseSha)
    const changedAfterFetch = diffChangedXlsxFiles(baseSha, headSha)
    if (changedAfterFetch) {
      return changedAfterFetch
    }
  }
  if (process.env.GITHUB_BASE_REF) {
    fetchRef(process.env.GITHUB_BASE_REF)
    const changed = diffChangedXlsxFiles(`origin/${process.env.GITHUB_BASE_REF}`, 'HEAD')
    if (changed) {
      return changed
    }
  }
  const previousCommitChanged = diffChangedXlsxFiles('HEAD~1', 'HEAD')
  if (previousCommitChanged) {
    return previousCommitChanged
  }

  throw new Error('Could not determine changed XLSX files. Set checkout fetch-depth: 0 or changed-files-only: "false".')
}

function diffChangedXlsxFiles(base, head) {
  const result = spawnSync('git', ['diff', '--name-only', '--diff-filter=ACMRT', `${base}...${head}`], {
    cwd: workspace,
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    return undefined
  }
  return new Set(
    result.stdout
      .split('\n')
      .map((line) => normalizePath(line.trim()))
      .filter((line) => line.endsWith('.xlsx')),
  )
}

function readEvent() {
  const eventPath = process.env.GITHUB_EVENT_PATH
  if (!eventPath || !existsSync(eventPath)) {
    return undefined
  }
  return JSON.parse(readFileSync(eventPath, 'utf8'))
}

function fetchRef(ref) {
  spawnSync('git', ['fetch', '--no-tags', '--depth=1', 'origin', ref], {
    cwd: workspace,
    encoding: 'utf8',
  })
}

function listFiles(root) {
  const files = []
  const stack = [root]
  while (stack.length > 0) {
    const current = stack.pop()
    if (!current) {
      continue
    }
    for (const entry of readdirSync(current)) {
      if (entry === '.git' || entry === 'node_modules') {
        continue
      }
      const path = `${current}${sep}${entry}`
      const stat = statSync(path)
      if (stat.isDirectory()) {
        stack.push(path)
      } else if (stat.isFile()) {
        files.push(path)
      }
    }
  }
  return files
}

function hasGlob(pattern) {
  return /[*?[\]]/u.test(pattern)
}

function globToRegExp(pattern) {
  const segments = pattern.split('/')
  let source = '^'
  for (const segment of segments) {
    if (segment === '**') {
      source += '(?:[^/]+/)*'
      continue
    }
    source += `${globSegmentToRegExp(segment)}/`
  }
  const trimmedSource = source.endsWith('/') ? source.slice(0, -1) : source
  return new RegExp(`${trimmedSource}$`, 'u')
}

function globSegmentToRegExp(segment) {
  let source = ''
  for (const char of segment) {
    if (char === '*') {
      source += '[^/]*'
    } else if (char === '?') {
      source += '[^/]'
    } else {
      source += char.replace(/[|\\{}()[\]^$+*?.]/gu, '\\$&')
    }
  }
  return source
}

function normalizePath(path) {
  return path.replaceAll('\\', '/').replace(/^\.\//u, '')
}

function unique(items) {
  return [...new Set(items)].toSorted((left, right) => left.localeCompare(right))
}

function writeGithubOutput(name, value) {
  appendFileSync(process.env.GITHUB_OUTPUT, `${name}=${value}\n`)
}
