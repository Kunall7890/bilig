---
title: Bilig product surface notes
published: false
description: Internal product-surface notes, assets, and proof links for presenting @bilig/workpaper without canned launch copy.
tags: product, proof, typescript, node, workbook, workpaper
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/product-hunt-thumbnail.png
image: /assets/product-hunt-thumbnail.png
---

# Bilig Product Surface Notes

Use this page when a product or directory surface needs short, factual Bilig
copy plus the current proof links and assets. Do not use canned comments,
generic launch copy, or vague tool-host positioning.

## Product Copy

Name:

```text
Bilig
```

Tagline:

```text
WorkPaper formulas for TypeScript services.
```

Short description:

```text
Bilig runs workbook-shaped business rules in Node: edit inputs, recalculate
formulas, read outputs, and save WorkPaper JSON without opening a spreadsheet
UI.
```

## Proof Links

- Homepage: <https://proompteng.github.io/bilig/>
- npm package: <https://www.npmjs.com/package/@bilig/workpaper>
- Repository: <https://github.com/proompteng/bilig>
- Node service evaluator:
  <https://proompteng.github.io/bilig/eval-workpaper-service.html>
- MCP tool evaluator: <https://proompteng.github.io/bilig/eval-agent-mcp.html>
- Benchmark scope:
  <https://proompteng.github.io/bilig/what-workpaper-benchmark-proves.html>
- Compatibility limits:
  <https://proompteng.github.io/bilig/where-bilig-is-not-excel-compatible-yet.html>
- MCP setup: <https://proompteng.github.io/bilig/mcp-client-setup.html>
- Starter issues:
  <https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only>

## Assets

Thumbnail:

![Product thumbnail](../../docs/assets/product-hunt-thumbnail.png)

Gallery:

![Workbook API gallery image](../../docs/assets/product-hunt-gallery-01-workbook-api.png)

![Formula readback gallery image](../../docs/assets/product-hunt-gallery-02-agent-readback.png)

![Node service gallery image](../../docs/assets/product-hunt-gallery-03-node-service.png)

Video:

<video controls src="../../docs/assets/product-hunt-demo.webm" title="Bilig product demo"></video>

The WebM is for docs and social previews. If a launch form needs video, upload
the demo to YouTube first or omit the video field.

## Fit Check

These checks follow Product Hunt's launch prep guidance:
<https://www.producthunt.com/launch/preparing-for-launch> and
<https://www.producthunt.com/launch/>.

- Availability: the npm evaluator is public and runnable before any submission.
- Account: submit from a personal maker account, not a company account.
- Timing: a Product Hunt launch day starts at midnight PST.
- Ask: invite people to check it out and leave feedback. Do not ask for
  upvotes.
- Thumbnail: `240x240`, below Product Hunt's 2 MB image limit.
- Gallery images: `1270x760`, below Product Hunt's 5 MB image limit.
- Video: use a YouTube link in Product Hunt; the local WebM is not the
  launch-form video field.

## Proof To Lead With

- The evaluator installs from npm and does not clone the monorepo.
- The public package is `@bilig/workpaper`.
- A useful run returns `verified: true` after edit, recalculation, save, and
  restore.
- The benchmark page documents the narrow comparable-row claim and the worst
  p95 row, currently `sheet-rename-dependencies`.
- The compatibility page states what is not Excel-compatible yet before a user
  tries to import a real workbook.

## Launch Checklist

1. Link to a proof page, not only the repository.
2. Upload the thumbnail and three gallery images.
3. Use a YouTube link only if the demo has been uploaded there; otherwise omit
   video from the form.
4. Submit from a personal maker account.
5. Ask for feedback, not upvotes.
6. Stay online to answer questions about compatibility, import/export, MCP
   setup, and benchmark scope.
7. If somebody asks for a missing workflow, turn it into a small issue or a
   focused example.
