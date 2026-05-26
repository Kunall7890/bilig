#!/usr/bin/env node

import { spawnSync } from 'node:child_process'
import process from 'node:process'

const defaultWorkflowName = 'github-ci'
const defaultBranch = 'main'
const defaultTimeoutSeconds = 30 * 60
const defaultPollSeconds = 15

const options = parseArgs(process.argv.slice(2))
const targetRepo = requiredOption(options.repo, '--repo')
const targetSha = requiredOption(options.sha, '--sha')
const workflowName = options.workflow ?? defaultWorkflowName
const branchName = options.branch ?? defaultBranch
const timeoutSeconds = parsePositiveInteger(options.timeoutSeconds ?? process.env.GITHUB_CI_WAIT_TIMEOUT_SECONDS, defaultTimeoutSeconds)
const pollSeconds = parsePositiveInteger(options.pollSeconds ?? process.env.GITHUB_CI_WAIT_POLL_SECONDS, defaultPollSeconds)
const deadline = Date.now() + timeoutSeconds * 1000

let attempt = 1
while (Date.now() <= deadline) {
  const runsResult = listRuns({ repo: targetRepo, sha: targetSha, workflow: workflowName, branch: branchName })
  if (!runsResult.ok) {
    console.log(`${workflowName} run lookup failed for ${targetSha}: ${runsResult.error}; attempt ${attempt}`)
    sleep(pollSeconds)
    attempt += 1
    continue
  }

  const run = latestMatchingRun(runsResult.runs, targetSha, workflowName)
  if (!run) {
    console.log(`No ${workflowName} run found yet for ${targetSha}; attempt ${attempt}`)
    sleep(pollSeconds)
    attempt += 1
    continue
  }

  if (run.status === 'completed' && run.conclusion === 'success') {
    console.log(`${workflowName} succeeded for ${targetSha}: ${run.url}`)
    process.exit(0)
  }

  if (run.status === 'completed') {
    fail(`${workflowName} did not pass for ${targetSha}: conclusion=${run.conclusion ?? 'unknown'} url=${run.url}`)
  }

  console.log(`${workflowName} is ${run.status} for ${targetSha}: ${run.url}; attempt ${attempt}`)
  sleep(pollSeconds)
  attempt += 1
}

fail(`Timed out after ${timeoutSeconds}s waiting for ${workflowName} to pass for ${targetSha}`)

function parseArgs(argv) {
  const parsed = {}
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    if (!arg.startsWith('--')) {
      fail(`Unexpected positional argument: ${arg}`)
    }

    const [key, inlineValue] = arg.slice(2).split('=', 2)
    const nextValue = inlineValue ?? argv[index + 1]
    if (!nextValue || nextValue.startsWith('--')) {
      fail(`Missing value for --${key}`)
    }

    parsed[toCamelCase(key)] = nextValue
    if (inlineValue === undefined) {
      index += 1
    }
  }

  return parsed
}

function toCamelCase(value) {
  return value.replaceAll(/-([a-z])/g, (_match, letter) => letter.toUpperCase())
}

function requiredOption(value, name) {
  if (!value) {
    fail(`${name} is required`)
  }
  return value
}

function parsePositiveInteger(value, fallback) {
  if (!value) {
    return fallback
  }

  const parsed = Number.parseInt(value, 10)
  if (!Number.isSafeInteger(parsed) || parsed <= 0) {
    fail(`Expected a positive integer, received: ${value}`)
  }

  return parsed
}

function listRuns({ repo, sha, branch }) {
  const [owner, repoName] = repo.split('/', 2)
  if (!owner || !repoName) {
    fail(`Expected --repo in owner/name form, received: ${repo}`)
  }

  const path = `repos/${encodeURIComponent(owner)}/${encodeURIComponent(repoName)}/actions/runs`
  const query = `branch=${encodeURIComponent(branch)}&head_sha=${encodeURIComponent(sha)}&per_page=20`
  const result = spawnSync(
    'gh',
    [
      'api',
      `${path}?${query}`,
      '--jq',
      '.workflow_runs | map({databaseId: .id, status, conclusion, url: .html_url, headSha: .head_sha, workflowName: .name, displayTitle: .display_title, createdAt: .created_at})',
    ],
    {
      encoding: 'utf8',
    },
  )

  if (result.error) {
    return { ok: false, error: `failed to execute gh: ${result.error.message}` }
  }

  if (result.status !== 0) {
    return { ok: false, error: `gh api failed with exit ${result.status}: ${result.stderr.trim()}` }
  }

  try {
    const runs = JSON.parse(result.stdout)
    if (!Array.isArray(runs)) {
      return { ok: false, error: 'gh api returned a non-array JSON payload' }
    }
    return { ok: true, runs }
  } catch (error) {
    return { ok: false, error: `could not parse gh api JSON: ${error instanceof Error ? error.message : String(error)}` }
  }
}

function latestMatchingRun(runs, sha, workflow) {
  return runs.find((run) => run && typeof run === 'object' && run.headSha === sha && run.workflowName === workflow)
}

function sleep(seconds) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, seconds * 1000)
}

function fail(message) {
  console.error(message)
  process.exit(1)
}
