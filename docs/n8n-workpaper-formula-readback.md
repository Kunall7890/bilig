# n8n WorkPaper Formula Readback

Use this when an n8n workflow needs spreadsheet formulas but the important
operation is not editing a visible Excel grid. The workflow writes one input,
recalculates dependent formulas, reads the computed outputs, and checks that the
WorkPaper JSON restores to the same result.

## Community Node

Use the scoped community node when you want a native n8n node instead of the
zero-install HTTP Request workflow:

```text
@bilig/n8n-nodes-workpaper
```

Install it from **Settings** -> **Community nodes** in n8n, or install the same
package from npm in a self-hosted deployment:

```sh
npm install @bilig/n8n-nodes-workpaper
```

The node is a thin HTTP integration around the same formula-readback endpoint.
It has no credentials for the hosted demo path; point `Bilig Base URL` at your
own Bilig deployment for production data.

The community node also has a `WorkPaper JSON` -> `Evaluate Document` operation
for user-owned workbook state. That operation posts a WorkPaper JSON document,
cell edits, and readback cells to:

```text
POST /api/workpaper/n8n/evaluate
```

Use it when the workflow already owns the workbook model and needs the next n8n
node to receive both formula readback proof and the updated WorkPaper JSON.

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

Start the local formula-readback server from npm; no Bilig checkout is needed:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-n8n-formula-server --port 4321
```

The same server exposes both endpoints:

```text
POST /api/workpaper/n8n/forecast
POST /api/workpaper/n8n/evaluate
```

The self-hosted workflow defaults to that local Bilig endpoint:

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

That keeps the n8n surface small and reproducible. Use the community node when
you want a native n8n package; use the workflow JSON when you want the most
inspectable proof with only built-in nodes.

## Privacy and Dependency Boundary

The hosted workflow is a demo. It sends the selected sheet name, cell address,
and value to `bilig.proompteng.ai`.

The generic `Evaluate Document` operation sends the provided WorkPaper JSON
document to the configured Bilig base URL. For private workbook data, run the
server inside your own network and point the n8n node at that internal URL.

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
