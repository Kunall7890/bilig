# Generic Workbook Agent Model

This example shows the intended `@bilig/workbook` shape without a built-in
business model. The consumer defines the model. Bilig supplies generic selectors,
formula helpers, static verification, runtime requirements, and a proof-shaped
run result.

```sh
npm install
npm start
```

The model finds a table by headers, narrows rows by a generic predicate, writes a
formula to a result column, and declares checks. The adapter in this folder is a
tiny stand-in for `@bilig/core` or an app runtime: it applies the plan, reads
back the formula target, and verifies generic checks.

The output answers the questions an agent needs before it trusts a workbook
mutation:

- what model and action were selected
- which refs the selectors bound
- which commands and low-level ops were planned
- which adapter capabilities are required
- which checks passed
- what proof supported each check

No revenue, quote, forecast, or other domain-specific model is built into
`@bilig/workbook`.
