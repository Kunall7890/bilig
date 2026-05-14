---
title: Product Hunt launch kit for bilig
published: true
description: Public launch copy, assets, proof points, and first-comment text for introducing @bilig/headless as a TypeScript WorkPaper runtime for Node services and agents.
tags: producthunt, launch, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/product-hunt-launch-kit.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/product-hunt-thumbnail.png
image: /assets/product-hunt-thumbnail.png
---

# Product Hunt Launch Kit For Bilig

Use this page when a launch surface needs the short version of what `bilig`
does, the proof links, and the assets in one place.

Do not launch with a vague "AI spreadsheet" pitch. The useful claim is smaller:
`@bilig/headless` runs workbook formulas from TypeScript, edits cells through an
API, reads the calculated value, and saves the WorkPaper as JSON.

## Product Copy

Name:

```text
bilig
```

Tagline:

```text
Workbook formulas for TypeScript services and agents.
```

Short description:

```text
bilig is a headless WorkPaper runtime for Node.js code that needs spreadsheet
formulas without opening a browser grid. Build sheets, write an input, read the
recalculated value, and save the workbook state as JSON.
```

First comment:

```text
I maintain bilig. The shortest way to judge it is the npm-only smoke test:
start from an empty Node project, install @bilig/headless, run eval.ts, edit an
input, read the recalculated value, save WorkPaper JSON, restore it, and check
the value again.

It is for backend and agent workflows where formulas are product logic, not for
manual spreadsheet editing. The benchmark and compatibility gaps are public:
46/46 comparable mean-latency rows are faster in the checked WorkPaper vs
HyperFormula artifact, one duplicate-lookup p95 row is slower, and UI rendering
is out of scope.

Useful feedback would be concrete: which formula family, persistence shape, MCP
client, or import/export path would block you from trying this in a real
service?
```

## Links

- Homepage: <https://proompteng.github.io/bilig/>
- npm smoke test: <https://proompteng.github.io/bilig/try-bilig-headless-in-node.html>
- Repository: <https://github.com/proompteng/bilig>
- npm package: <https://www.npmjs.com/package/@bilig/headless>
- Benchmark notes: <https://proompteng.github.io/bilig/what-workpaper-benchmark-proves.html>
- Compatibility gaps: <https://proompteng.github.io/bilig/where-bilig-is-not-excel-compatible-yet.html>
- MCP setup: <https://proompteng.github.io/bilig/mcp-client-setup.html>
- Starter issues:
  <https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only>

## Assets

Thumbnail:

![Product Hunt thumbnail](assets/product-hunt-thumbnail.png)

Gallery:

![Workbook API gallery image](assets/product-hunt-gallery-01-workbook-api.png)

![Agent readback gallery image](assets/product-hunt-gallery-02-agent-readback.png)

![Node service gallery image](assets/product-hunt-gallery-03-node-service.png)

Video:

<video controls src="assets/product-hunt-demo.webm" title="bilig Product Hunt launch demo"></video>

## Proof To Lead With

- The smoke test installs from npm and does not clone the monorepo.
- The example is TypeScript: `eval.ts` is maintained in
  `examples/headless-workpaper/npm-eval.ts`.
- The output must show `verified: true` after save and restore.
- The public benchmark page documents the narrow `46/46` comparable mean-row
  claim and the slower duplicate-lookup p95 caveat.
- The compatibility page says what is not Excel-compatible yet before a user
  tries to import a real workbook.

## Launch Checklist

1. Link to the npm smoke test, not only the repository.
2. Upload the thumbnail, three gallery images, and the WebM demo.
3. Pin the first comment above.
4. Stay online to answer questions about Excel compatibility, XLSX import,
   MCP setup, and benchmark scope.
5. If somebody asks for a missing workflow, turn it into a small
   `first-timers-only` issue or a focused example.
