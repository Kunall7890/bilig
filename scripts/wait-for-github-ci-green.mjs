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
  const run = latestMatchingRun(listRuns({ repo: targetRepo, sha: targetSha, workflow: workflowName, branch: branchName }), targetSha)
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

function listRuns({ repo, sha, workflow, branch }) {
  const result = spawnSync(
    'gh',
    [
      'run',
      'list',
      '--repo',
      repo,
      '--workflow',
      workflow,
      '--commit',
      sha,
      '--branch',
      branch,
      '--limit',
      '20',
      '--json',
      'databaseId,status,conclusion,url,headSha,workflowName,displayTitle,createdAt',
    ],
    {
      encoding: 'utf8',
    },
  )

  if (result.error) {
    fail(`Failed to execute gh: ${result.error.message}`)
  }

  if (result.status !== 0) {
    fail(`gh run list failed with exit ${result.status}: ${result.stderr.trim()}`)
  }

  try {
    const runs = JSON.parse(result.stdout)
    if (!Array.isArray(runs)) {
      fail('gh run list returned a non-array JSON payload')
    }
    return runs
  } catch (error) {
    fail(`Could not parse gh run list JSON: ${error instanceof Error ? error.message : String(error)}`)
  }
}

function latestMatchingRun(runs, sha) {
  return runs.find((run) => run && typeof run === 'object' && run.headSha === sha)
}

function sleep(seconds) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, seconds * 1000)
}

function fail(message) {
  console.error(message)
  process.exit(1)
}
