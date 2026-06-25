import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

const currentStarterIssueNumbers: readonly number[] = [452, 453]

const closedStarterIssueNumbers = [
  137, 138, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 154, 224, 231, 199, 200, 201, 202, 203, 204, 205, 228, 229, 246,
  266, 282, 294, 160, 161, 164, 165, 166, 168, 169, 170, 171, 172, 173, 174, 175, 176, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187,
  188, 189, 190, 191, 192, 276, 227, 247, 256, 315, 316, 317, 318, 319, 336, 341, 343, 344, 345, 346, 347, 354, 364, 374, 377, 380, 285,
  273, 283, 300, 360, 361, 362, 363, 366, 367, 368, 369, 371, 358, 334, 454,
] as const

interface StarterIssueDiscoveryInputs {
  readonly contributing: string
  readonly llms: string
  readonly newContributorGuide: string
  readonly starterIssues: string
}

export function requireStarterIssueDiscovery(starterIssues: string, llms: string): void
export function requireStarterIssueDiscovery(inputs: StarterIssueDiscoveryInputs): void
export function requireStarterIssueDiscovery(inputsOrStarterIssues: StarterIssueDiscoveryInputs | string, maybeLlms?: string): void {
  if (typeof inputsOrStarterIssues === 'string') {
    requireStarterIssueUrlDiscovery(inputsOrStarterIssues, maybeLlms ?? '')
    return
  }

  const { contributing, llms, newContributorGuide, starterIssues } = inputsOrStarterIssues

  requireStarterIssueUrlDiscovery(starterIssues, llms)
  requireIncludes(newContributorGuide, '## First-Time Command Checklist', 'docs/new-contributor-guide.md')
  requireIncludes(newContributorGuide, 'pnpm docs:discovery:check', 'docs/new-contributor-guide.md')
  requireIncludes(newContributorGuide, 'pnpm format:check', 'docs/new-contributor-guide.md')
  requireIncludes(newContributorGuide, 'pnpm lint', 'docs/new-contributor-guide.md')
  requireIncludes(newContributorGuide, 'first-time contributor review happens on GitHub.', 'docs/new-contributor-guide.md')
  requireIncludes(contributing, 'pull requests on GitHub are welcome; maintainers', 'CONTRIBUTING.md')
  requireIncludes(starterIssues, 'new-contributor-guide.md#first-time-command-checklist', 'docs/starter-issues.md')
  requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/blob/main/CONTRIBUTING.md', 'docs/starter-issues.md')
  requireIncludes(starterIssues, 'Current starter queue as of June 25, 2026:', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '2 open `good first issue` issues.', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '2 open `first-timers-only` issues.', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '2 open `help wanted` issues.', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '1 starter issue is a code or test task.', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '1 starter issue is a focused docs or integration transcript task.', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '0 starter issues are currently under active review.', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '## Start Here This Week', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '#452 docs(starter): add current WorkPaper service proof transcript', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '#453 examples(starter): add Hono WorkPaper route smoke', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '## Code And Test Starters', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '#453 examples(starter): add Hono WorkPaper route smoke', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '## Integration Docs Starters', 'docs/starter-issues.md')
  requireIncludes(starterIssues, '#452 docs(starter): add current WorkPaper service proof transcript', 'docs/starter-issues.md')
  requireIncludes(starterIssues, 'Add `help wanted` only when an external contributor can make progress', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, '115 open `first-timers-only` issues.', 'docs/starter-issues.md')
  requireNotIncludes(
    starterIssues,
    '#370: test(examples): add malformed CSV fixture check to the csv-shaped smoke',
    'docs/starter-issues.md',
  )
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/373', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/271', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/291', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/295', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/251', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/349', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/351', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/352', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/353', 'docs/starter-issues.md')
  requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/365', 'docs/starter-issues.md')
  requireIncludes(contributing, 'new-contributor-guide.md#first-time-command-checklist', 'CONTRIBUTING.md')
  requireIncludes(llms, 'first-patch list capped at 15 scoped issues.', 'docs/llms.txt')
  requireNotIncludes(llms, 'https://github.com/proompteng/bilig/issues/334', 'docs/llms.txt')
  requireNotIncludes(llms, 'https://github.com/proompteng/bilig/issues/272', 'docs/llms.txt')
  requireNotIncludes(llms, 'https://github.com/proompteng/bilig/issues/277', 'docs/llms.txt')
  requireNotIncludes(llms, 'https://github.com/proompteng/bilig/issues/281', 'docs/llms.txt')
}

function requireStarterIssueUrlDiscovery(starterIssues: string, llms: string): void {
  for (const issueNumber of currentStarterIssueNumbers) {
    const required = `https://github.com/proompteng/bilig/issues/${issueNumber}`
    requireIncludes(starterIssues, required, 'docs/starter-issues.md')
    requireIncludes(llms, required, 'docs/llms.txt')
  }

  for (const issueNumber of closedStarterIssueNumbers) {
    const issueUrl = `https://github.com/proompteng/bilig/issues/${issueNumber}`
    requireNotIncludes(starterIssues, issueUrl, 'docs/starter-issues.md')
    requireNotIncludes(llms, issueUrl, 'docs/llms.txt')
  }
}
