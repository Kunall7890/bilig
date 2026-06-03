---
title: Starter Issues
published: true
description: Current small Bilig starter issues for contributors, focused on credible code, test, example, and integration tasks.
tags: contributors, starter-issues, open-source, workpaper
canonical_url: https://proompteng.github.io/bilig/starter-issues.html
---

# Starter Issues

This page is the stable contributor on-ramp for small public `bilig` tasks. It
intentionally stays short: GitHub's `good first issue`, `first-timers-only`, and
`help wanted` labels should point to work that is current, scoped, and credible
for someone opening the repository cold.

Current starter queue as of June 3, 2026:

- 1 open `good first issue` issue.
- 1 open `first-timers-only` issue.
- 1 open `help wanted` issue.
- 0 starter issues are code or test tasks.
- 1 starter issue is a focused docs or integration transcript task.
- 0 starter issues are currently under active review.

## Start Here This Week

If you are opening the queue cold, pick one of these before browsing the full
issue list. They are small, current, and map to the public adoption path for
`@bilig/headless`.

- [#334: docs(agent): add OpenAI Responses streaming tool-call transcript](https://github.com/proompteng/bilig/issues/334)
  helps agent builders see the tool-call loop.

## Code And Test Starters

No code or test starter issues are currently open. Open a scoped code/test
ticket before the next contributor push if the queue should include one.

## Integration Docs Starters

- [#334: docs(agent): add OpenAI Responses streaming tool-call transcript](https://github.com/proompteng/bilig/issues/334)

## Claim A Starter Issue

Comment on the issue before opening a pull request. If the issue is unassigned,
a maintainer can assign it to you and keep the scope reserved while you work.
If it already has an assignee, pick another starter ticket or ask whether the
current assignee still wants help.

For a first patch, keep the pull request focused on the issue's acceptance
proof. Include the command you ran, mention the issue number, and open a draft
pull request early if any requirement is unclear. The
[new contributor guide](new-contributor-guide.md) gives the shortest setup,
code-map, and
[first-time command checklist](new-contributor-guide.md#first-time-command-checklist).
Read
[CONTRIBUTING.md](https://github.com/proompteng/bilig/blob/main/CONTRIBUTING.md)
before opening the pull request.

Useful filters:

- [`good first issue`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22good%20first%20issue%22)
- [`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
- [`help wanted`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22help%20wanted%22)

GitHub surfaces issues labeled `good first issue` in contributor discovery
paths, per
[GitHub's label guidance](https://docs.github.com/articles/helping-new-contributors-find-your-project-with-labels),
so starter tickets should stay genuinely scoped and current. Do not use that
label for cross-cutting formula, import/export, or runtime changes that require
broad architectural context.

Use `first-timers-only` only for issues that are ready for someone making their
first contribution to this repository. Those issues should name the expected
files, a copyable validation command, and a narrow acceptance proof in the issue
body.

Add `help wanted` only when an external contributor can make progress without
private context or maintainer-only systems.
