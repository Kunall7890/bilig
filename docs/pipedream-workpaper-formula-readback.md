---
title: Pipedream WorkPaper formula readback
published: true
description: Use Bilig WorkPaper from Pipedream to verify formula-backed workflow fields without Excel, LibreOffice, or browser automation.
tags: pipedream, automation, spreadsheet, workpaper, formulas
canonical_url: https://proompteng.github.io/bilig/pipedream-workpaper-formula-readback.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Pipedream WorkPaper Formula Readback

Bilig can be used from Pipedream as a workflow action. The action writes one
forecast input cell, asks Bilig WorkPaper to recalculate dependent formulas, and
fails unless Bilig returns verified readback proof.

The public upstream review is:

- <https://github.com/PipedreamHQ/pipedream/pull/20972>
- Head branch: `gregkonush:bilig-workpaper-verify-readback`
- Pipedream app slug: `bilig_workpaper`
- Action key: `bilig_workpaper-verify-formula-readback`

Use this path when a Pipedream workflow needs spreadsheet-style business logic
without opening Excel, LibreOffice, Google Sheets, or a browser UI.

## Proof Shape

The action calls:

```text
POST /api/workpaper/n8n/forecast
```

with:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

and requires:

```json
{
  "verified": true,
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

The default hosted proof writes `0.4` to `Inputs!B3`, changes expected ARR from
`60000` to `96000`, and returns the edited cell plus before/after computed
values.

## Pipedream Contract

The upstream action uses the connected account's `base_url` from `$auth`. It
does not expose Base URL as an action prop. This matches Pipedream maintainer
feedback and keeps the runtime input focused on workbook data:

- sheet name
- input cell address
- numeric input value
- value divisor for percentage-style inputs

The local staging artifact is under:

```text
integrations/pipedream-bilig-workpaper
```

The upstream PR is the source of truth for the Pipedream registry review state
and final app slug.

## Boundary

Pipedream should own workflow triggers, schedules, branching, and downstream
steps. Bilig should own workbook formula execution and proof that recalculated
values survived export/restore. For private workbook data, run your own Bilig
service and set the Pipedream connected account Base URL to that host.
