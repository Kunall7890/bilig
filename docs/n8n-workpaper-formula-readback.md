# n8n WorkPaper Formula Readback

Use this when an n8n workflow needs spreadsheet formulas but the important
operation is not editing a visible Excel grid. The workflow writes one input,
recalculates dependent formulas, reads the computed outputs, and checks that the
WorkPaper JSON restores to the same result.

## Importable Workflow

The example workflows live in:

```text
examples/n8n-workpaper-formula-readback/bilig-workpaper-formula-readback.n8n.json
examples/n8n-workpaper-formula-readback/bilig-workpaper-formula-readback.self-hosted.n8n.json
```

It uses only built-in n8n nodes:

- Manual Trigger
- Code
- HTTP Request
- Code

n8n imports workflows as JSON, so the file can be imported directly from the
editor. See the n8n workflow import/export docs:
<https://docs.n8n.io/workflows/export-import/>.

## Hosted Demo vs Self-Hosted Route

The hosted demo workflow defaults to the public Bilig endpoint so someone can
import it and run the proof before deploying Bilig:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

The self-hosted workflow defaults to a local Bilig app:

```text
POST http://host.docker.internal:4321/api/workpaper/n8n/forecast
```

and falls back to:

```text
POST http://localhost:4321/api/workpaper/n8n/forecast
```

Use `host.docker.internal` when n8n runs in Docker and Bilig runs on the host.
Use `localhost` when n8n and Bilig run in the same host network. Change
`baseUrl` in the `Choose local forecast input` node if your Bilig app has a
different internal URL.

Request:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

Response shape:

```json
{
  "verified": true,
  "editedCell": "Inputs!B3",
  "before": {
    "expectedArr": 60000
  },
  "after": {
    "expectedArr": 96000,
    "targetGap": 5600
  },
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

Editable inputs in the demo forecast WorkPaper:

| Cell | Meaning |
| --- | --- |
| `B2` | Qualified opportunities |
| `B3` | Win rate |
| `B4` | Average ARR |
| `B5` | Expansion multiplier |

## Why This Fits n8n

n8n should orchestrate the workflow. Bilig owns the formula workbook step:

1. receive one spreadsheet-shaped input edit;
2. recalculate formulas in Node;
3. return the computed readback;
4. export and restore WorkPaper JSON as proof.

That keeps the n8n surface small and reproducible before a custom community node
exists.

## Privacy and Dependency Boundary

The hosted workflow is a demo. It sends the selected sheet name, cell address,
and value to `bilig.proompteng.ai`.

For production data, use the self-hosted workflow and keep the Bilig route on
your own network. That removes the public hosted dependency while keeping the
n8n workflow inspectable: Manual Trigger, Code, HTTP Request, Code.

## Formula.js Boundary

For a small scalar formula that lives entirely in a Code node, `formulajs` or
plain JavaScript can be a better fit.

Bilig is for workbook-shaped state: formulas stored in cells, range reads, cell
edits, recalculation, JSON persistence, restore proof, and a result the next n8n
node can trust.

n8n Cloud does not allow arbitrary external npm modules in Code nodes.
Self-hosted n8n can allow external modules with `NODE_FUNCTION_ALLOW_EXTERNAL`,
but the module must still be installed in the n8n runtime. The Bilig workflow
keeps that dependency outside the Code node and makes the workbook calculation a
single local HTTP step.
