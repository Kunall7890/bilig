# Bilig WorkPaper Formula Readback For n8n

This example is an importable n8n workflow for the spreadsheet formula problem
n8n users keep hitting: write an input value, recalculate formulas, and verify
the computed output without opening Excel, LibreOffice, Google Sheets, or a
browser spreadsheet UI.

The workflow uses only built-in n8n nodes:

- Manual Trigger
- Code
- HTTP Request
- Code

There are two importable workflow files:

| File | Default endpoint | Use it when |
| --- | --- | --- |
| `bilig-workpaper-formula-readback.n8n.json` | `https://bilig.proompteng.ai` | You want the fastest hosted demo before deploying anything. |
| `bilig-workpaper-formula-readback.self-hosted.n8n.json` | `http://host.docker.internal:4321`, then `http://localhost:4321` | You want the formula readback step to stay inside your own local or self-hosted Bilig app. |

The hosted workflow calls this public demo endpoint:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

The self-hosted workflow calls:

```text
POST http://host.docker.internal:4321/api/workpaper/n8n/forecast
```

and falls back to:

```text
POST http://localhost:4321/api/workpaper/n8n/forecast
```

Use `host.docker.internal` when n8n runs in Docker and Bilig runs on the host.
Use `localhost` when n8n and Bilig run in the same host network. Change
`baseUrl` in the `Choose forecast input` node if your Bilig app has a different
internal URL.

The route edits one input cell in a demo forecast WorkPaper, recalculates the
summary formulas, exports and restores the WorkPaper JSON, and returns proof
that the formula output changed and survived restore.

## Import

1. Open n8n.
2. Choose Import from File.
3. Select `bilig-workpaper-formula-readback.n8n.json` for the hosted demo or
   `bilig-workpaper-formula-readback.self-hosted.n8n.json` for a local Bilig
   route.
4. Run the workflow manually.

n8n documents workflow import/export as JSON:
<https://docs.n8n.io/workflows/export-import/>.

## Expected Proof

The final node returns a compact object like:

```json
{
  "verdict": "verified",
  "editedCell": "Inputs!B3",
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000,
  "targetGap": 5600,
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

Change the input in the `Choose forecast input` node if you want to test a
different editable cell:

- `B2`: qualified opportunities
- `B3`: win rate
- `B4`: average ARR
- `B5`: expansion multiplier

## Data Boundary

The hosted demo is only for quick inspection. The request sends the selected
sheet name, cell address, and value to `bilig.proompteng.ai`.

For private workflow data, use the self-hosted workflow and point `baseUrl` at
your own Bilig app. In that mode n8n still uses only built-in nodes, but the
calculation and verification happen on infrastructure you control.

## Why not just use formula.js in a Code node?

Use `formulajs` or plain JavaScript inside n8n when the workflow only needs a
small scalar formula and you are comfortable owning the formula expression in
the Code node.

Use this Bilig workflow when the calculation is workbook-shaped: multiple
sheets, cell addresses, stored formulas, restore proof, and a need to verify
that the value read by the next node came from recalculation rather than a stale
cache.

n8n Cloud does not allow arbitrary external npm modules in the Code node. In
self-hosted n8n, external modules require the `NODE_FUNCTION_ALLOW_EXTERNAL`
setting and the module must be available to the n8n runtime. The workflow here
avoids that Code-node module setup by keeping the workbook runtime behind one
HTTP step.

This is intentionally not a custom n8n node yet. It is the smallest reproducible
workflow that proves the formula-workbook value path in n8n before asking users
to install anything.
