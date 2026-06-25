# Hono WorkPaper Route Example

A minimal, runnable smoke that mounts the shared WorkPaper request handler
inside a [Hono](https://hono.dev) app and proves the route shape end-to-end:

1. **GET** `/api/workpaper/summary` — reads formula results before any edit.
2. **POST** `/api/workpaper/revenue` — edits revenue records, lets the
   WorkPaper recalculate, and reads back the dependent `Summary!B2`
   (`Total revenue`) formula result.
3. **GET** `/api/workpaper/summary` again — confirms the workbook persisted
   correctly between the write and the subsequent read.

The handler uses web-standard `Request`/`Response` and is adapted to Hono
via `c.req.raw` — the same two-line pattern documented in
[`docs/node-framework-workpaper-adapters.md`](../../docs/node-framework-workpaper-adapters.md).

## Install and run

From a **clean checkout** of this repository:

```sh
pnpm --dir examples/hono-workpaper-route install --ignore-workspace
pnpm --dir examples/hono-workpaper-route run smoke
```

## Expected output

```json
{
  "route": "Hono WorkPaper Route",
  "inputCell": "Revenue!D2:Dn",
  "readbackCell": "Summary!B2",
  "before": {
    "totalRevenue": 36900,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "edit": {
    "records": 4,
    "after": {
      "totalRevenue": 48600,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "checks": {
      "totalRevenueChanged": true,
      "formulasPersisted": true,
      "serializedBytes": 1194
    }
  },
  "after": {
    "totalRevenue": 48600,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "success": true
}
```

The output confirms:

| Field | Meaning |
|---|---|
| `route` | Hono route name |
| `inputCell` | Cell range written by the POST body |
| `readbackCell` | Formula cell read back to prove recalculation |
| `before` | Summary values before the edit |
| `edit.records` | Number of revenue records written |
| `edit.after` | Summary values immediately after the edit |
| `edit.checks.totalRevenueChanged` | Formula result changed after the write |
| `edit.checks.formulasPersisted` | `=SUM(...)` survives serialization |
| `edit.checks.serializedBytes` | Workbook JSON byte size (non-zero) |
| `after` | Summary values read back from persisted workbook |
| `success` | `true` when all assertions pass |

`serializedBytes` can change as the persisted document schema evolves.
Treat it as a positive persistence signal, not a golden value.

## How it works

`hono-workpaper-route.ts` creates one `createInMemoryWorkbookStorage` instance
and builds a Hono app with two routes:

```ts
const honoHandler = createHonoWorkPaperHandler()
app.get('/api/workpaper/summary', honoHandler)
app.post('/api/workpaper/revenue', honoHandler)
```

The adapter is two lines:

```ts
function createHonoWorkPaperHandler() {
  return (c: { req: { raw: Request } }): Promise<Response> => handler(c.req.raw)
}
```

Requests are dispatched through `app.fetch(new Request(...))` — the same
boundary Hono exposes on Cloudflare Workers, Deno Deploy, Bun, and Node.js
`@hono/node-server`. No HTTP server is started; the smoke is purely
in-process.

## Type checking

```sh
pnpm --dir examples/hono-workpaper-route run typecheck
```
