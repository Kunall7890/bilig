# __PROJECT_NAME__

Formula-backed quote approval API built with `@bilig/headless`.

```sh
npm install
npm run smoke
```

The smoke run writes quote inputs, recalculates workbook formulas, persists the
WorkPaper as JSON, restores it, and checks that the restored formula output
matches the live output.

Run a local API:

```sh
npm run dev
curl http://localhost:8788/api/quote/approval
curl -X POST http://localhost:8788/api/quote/approval \
  -H 'content-type: application/json' \
  -d '{"units":40,"listPrice":1200,"discount":0.05,"unitCost":760,"minimumMargin":0.3}'
```

Learn more: <https://github.com/proompteng/bilig>
