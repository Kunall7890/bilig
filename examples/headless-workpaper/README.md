# Headless WorkPaper Example

This example shows `@bilig/headless` running as a Node library with no browser
UI. It builds a small revenue workbook, evaluates formulas, applies an
agent-style edit, persists the workbook, restores it, and prints the verified
summary.

Run it outside the monorepo with the published package:

```sh
npm install
npm start
```

Expected output:

```json
{
  "initial": {
    "totalRevenue": 27300,
    "westCustomers": 30
  },
  "afterAgentEdit": {
    "totalRevenue": 36900,
    "westCustomers": 38,
    "enterpriseArpa": 1200,
    "qualifiedCustomerCounts": [20, 30, 18]
  },
  "persistedSheets": ["Deals", "Summary"]
}
```

The repository smoke test runs this same example against packed local runtime
packages through `pnpm workpaper:smoke:external`.
