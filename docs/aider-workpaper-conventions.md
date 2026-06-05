---
title: Aider WorkPaper conventions
published: true
description: Load Bilig WorkPaper conventions into Aider so spreadsheet-shaped coding tasks use formula readback, exported WorkPaper state, and explicit limitations instead of spreadsheet UI automation.
tags: aider, coding agents, spreadsheet automation, workpaper, formulas
canonical_url: https://proompteng.github.io/bilig/aider-workpaper-conventions.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Aider WorkPaper Conventions

Use this when Aider is editing a repo that contains workbook-shaped business
logic: pricing, approvals, payouts, budgets, forecasts, import validation, stale
XLSX formula caches, or spreadsheet formula readback after changing cells.

Aider's official convention flow is a good fit for this: load a small
`CONVENTIONS.md` file as read-only context with `/read CONVENTIONS.md` or
`aider --read CONVENTIONS.md`, or configure `.aider.conf.yml` so Aider loads the
file automatically from the repo. Bilig keeps that path boring and explicit:

- `.aider.conf.yml` reads `CONVENTIONS.md`.
- `CONVENTIONS.md` tells Aider to prefer WorkPaper state before Excel,
  LibreOffice, Google Sheets, browser grids, screenshots, or cached XLSX values
  when the workflow can run through code.
- The conventions require before/after formula readback, serialized or exported
  WorkPaper evidence, and explicit limitations before Aider reports success.

## First Check

Run the package-owned proof before trusting an agent workflow:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The result must include `verified: true`, the edited cell, before value, after
formula readback, exported or persisted WorkPaper state, and restore or restart
readback. A write call alone is not proof.

If the workbook contains provider-backed formulas such as `IMPORTRANGE`, run the
boundary case too:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
```

## Load Aider

In this repo, Aider loads the conventions automatically because
`.aider.conf.yml` contains:

```yaml
read:
  - CONVENTIONS.md
```

In another repo, either copy the same two files or explicitly start Aider with
the conventions:

```sh
aider --read CONVENTIONS.md
```

Keep `CONVENTIONS.md` focused on Aider's workbook proof policy. Broad repo
policy belongs in `AGENTS.md`; project MCP wiring belongs in an MCP config that
the host actually reads.

## WorkPaper Path

When state must persist, run the local file-backed MCP server from the
conventions:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
```

When the source is an `.xlsx`, start with the risk tool instead:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
```

That XLSX mode reports workbook risk indicators before an agent trusts the
imported WorkPaper. It does not certify Excel compatibility.

## Required Aider Response

Before saying a workbook is updated, Aider should return:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

If any readback step fails, report the blocker. Do not treat a write call,
terminal exit code, screenshot, or cached XLSX value as the final result.

## Current Source Boundary

The checked-in Bilig source contains the Aider convention path through
`CONVENTIONS.md` and `.aider.conf.yml`. Public `@latest` command surfaces should
still be checked with `bilig-evaluate --door agent-mcp --json` before use,
because npm publishing can lag the repository.

## Official Aider Docs Checked

- [Specifying coding conventions](https://aider.chat/docs/usage/conventions.html)
- [YAML config file](https://aider.chat/docs/config/aider_conf.html)

## Related

- [Coding agent rule chooser](agent-rule-chooser.md)
- [Agent adoption kit](agent-adoption-kit.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
